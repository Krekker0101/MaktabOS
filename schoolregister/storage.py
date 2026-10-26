"""Reliable SQLite persistence and private OS-user data location."""
from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import sqlite3
import sys
import tempfile
import uuid
from contextlib import contextmanager
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import quote


COLUMNS = (
    "id", "surname", "given_name", "patronymic", "date_of_birth", "grade",
    "section", "shift", "address", "guardian_name", "guardian_phone",
    "extra_phone", "admitted_on", "status", "notes", "from_school", "to_school", "left_on",
)
WRITABLE = tuple(x for x in COLUMNS if x != "id")
PASSWORD_ITERATIONS = 650_000


def data_directory() -> Path:
    override = os.getenv("SCHOOLREGISTER_DATA_DIR")
    if override:
        return Path(override).expanduser().resolve()
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    target = base / "MaktabOS"
    legacy_database = base / "SchoolRegister" / "school.db"
    new_database = target / "school.db"
    if not new_database.exists() and legacy_database.is_file():
        _migrate_legacy_database(legacy_database, new_database)
    return target


def _migrate_legacy_database(source_path: Path, target_path: Path) -> None:
    target_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=".maktabos-migration-", suffix=".sqlite", dir=target_path.parent, delete=False
    ) as temp_file:
        temporary_path = Path(temp_file.name)
    source = destination = None
    try:
        source_uri = f"file:{quote(source_path.as_posix(), safe='/:')}?mode=ro"
        source = sqlite3.connect(source_uri, uri=True)
        destination = sqlite3.connect(temporary_path)
        source.backup(destination)
        if destination.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise sqlite3.DatabaseError("Migrated database failed integrity check")
        destination.close()
        destination = None
        os.replace(temporary_path, target_path)
    finally:
        if destination is not None:
            destination.close()
        if source is not None:
            source.close()
        temporary_path.unlink(missing_ok=True)


def timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _secure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    # Never chmod a user's arbitrary home/external-backup folder.
    private_root = data_directory()
    if os.name != "nt" and (path == private_root or private_root in path.parents):
        path.chmod(0o700)


def _secure_file(path: Path) -> None:
    if os.name != "nt" and path.exists():
        path.chmod(0o600)


def school_year(today: date | None = None, start_month: int = 9, start_day: int = 1) -> int:
    """Academic year beginning year. Default 2026-09-01 -> 2026."""
    today = today or date.today()
    return today.year if (today.month, today.day) >= (start_month, start_day) else today.year - 1


class Database:
    def __init__(self, path: Path | None = None):
        self.path = Path(path) if path else data_directory() / "school.db"
        _secure_dir(self.path.parent)
        self.con = sqlite3.connect(self.path, timeout=20)
        self.con.row_factory = sqlite3.Row
        self.con.create_function("unicode_casefold",1,lambda text: (text or "").casefold(),deterministic=True)
        self.con.execute("PRAGMA journal_mode=WAL")
        self.con.execute("PRAGMA synchronous=FULL")
        self.con.execute("PRAGMA foreign_keys=ON")
        self.con.execute("PRAGMA busy_timeout=20000")
        self._schema()
        _secure_file(self.path)

    def _schema(self):
        if self.con.execute("PRAGMA user_version").fetchone()[0] > 2:
            raise RuntimeError("База создана более новой версией приложения")
        with self.con:
            self.con.executescript('''
                CREATE TABLE IF NOT EXISTS students (
                    id TEXT PRIMARY KEY,
                    surname TEXT NOT NULL,
                    given_name TEXT NOT NULL,
                    patronymic TEXT NOT NULL DEFAULT '',
                    date_of_birth TEXT NOT NULL,
                    grade INTEGER NOT NULL CHECK(grade BETWEEN 1 AND 12),
                    section TEXT NOT NULL DEFAULT '',
                    shift TEXT NOT NULL DEFAULT '',
                    address TEXT NOT NULL,
                    guardian_name TEXT NOT NULL DEFAULT '',
                    guardian_phone TEXT NOT NULL DEFAULT '',
                    extra_phone TEXT NOT NULL DEFAULT '',
                    admitted_on TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'Учится',
                    notes TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    deleted_at TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_students_class ON students(grade,section,deleted_at);
                CREATE INDEX IF NOT EXISTS idx_students_deleted ON students(deleted_at);
                CREATE INDEX IF NOT EXISTS idx_students_name ON students(surname,given_name);
                CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS audit_log(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    at TEXT NOT NULL,
                    action TEXT NOT NULL,
                    record_id TEXT,
                    detail TEXT NOT NULL DEFAULT ''
                );
            ''')
        # Incremental migration retains v1 records, login, audit and backups.
        existing = {r[1] for r in self.con.execute("PRAGMA table_info(students)")}
        with self.con:
            for name in ("from_school", "to_school", "left_on"):
                if name not in existing:
                    self.con.execute(f"ALTER TABLE students ADD COLUMN {name} TEXT NOT NULL DEFAULT ''")
            if "grade_year" not in existing:
                self.con.execute("ALTER TABLE students ADD COLUMN grade_year INTEGER")
                self.con.execute("UPDATE students SET grade_year=? WHERE grade_year IS NULL", (school_year(),))
            self.con.execute('''CREATE TABLE IF NOT EXISTS transfer_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id TEXT NOT NULL REFERENCES students(id),
                direction TEXT NOT NULL CHECK(direction IN ('in','out')),
                school_name TEXT NOT NULL,
                event_date TEXT NOT NULL DEFAULT '',
                recorded_at TEXT NOT NULL,
                source TEXT NOT NULL DEFAULT 'manual'
            )''')
            self.con.execute("CREATE INDEX IF NOT EXISTS idx_transfer_student ON transfer_events(student_id,id)")
            self.con.execute('''CREATE TABLE IF NOT EXISTS grade_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id TEXT NOT NULL REFERENCES students(id),
                old_grade INTEGER NOT NULL,
                new_grade INTEGER NOT NULL,
                school_year INTEGER NOT NULL,
                recorded_at TEXT NOT NULL,
                source TEXT NOT NULL
            )''')
            self.con.execute('''CREATE TABLE IF NOT EXISTS import_fingerprints (
                fingerprint TEXT PRIMARY KEY,
                student_id TEXT NOT NULL REFERENCES students(id),
                imported_at TEXT NOT NULL
            )''')
            self.con.execute("PRAGMA user_version=2")

    def close(self):
        self.con.close()

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str):
        with self.con:
            self.con.execute("INSERT INTO settings VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))

    def is_initialized(self) -> bool:
        return bool(self.get_setting("pw_hash"))

    @staticmethod
    def _validate_password(password: str) -> None:
        if len(password) < 10:
            raise ValueError("Пароль должен быть длиной не менее 10 символов")

    @staticmethod
    def _password_digest(password: str, salt: bytes) -> bytes:
        return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PASSWORD_ITERATIONS)

    def _store_password(self, password: str) -> None:
        self._validate_password(password)
        salt = secrets.token_bytes(16)
        digest = self._password_digest(password, salt)
        self.con.execute(
            "INSERT INTO settings(key,value) VALUES('pw_salt',?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (salt.hex(),),
        )
        self.con.execute(
            "INSERT INTO settings(key,value) VALUES('pw_hash',?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (digest.hex(),),
        )

    @staticmethod
    def _recovery_digest(recovery_key: str, salt: bytes) -> bytes:
        return hashlib.pbkdf2_hmac("sha256", recovery_key.encode("utf-8"), salt, PASSWORD_ITERATIONS)

    def set_password_initial(self, password: str, recovery_key: str | None = None):
        if self.is_initialized():
            raise ValueError("Пароль уже создан")
        self._validate_password(password)
        if recovery_key is not None:
            self._validate_recovery_key(recovery_key)
        with self.con:
            self._store_password(password)
            if recovery_key is not None:
                recovery_salt = secrets.token_bytes(16)
                recovery_hash = self._recovery_digest(recovery_key, recovery_salt)
                self.con.execute("INSERT INTO settings VALUES(?,?)", ("recovery_salt", recovery_salt.hex()))
                self.con.execute("INSERT INTO settings VALUES(?,?)", ("recovery_hash", recovery_hash.hex()))
            self.con.execute("INSERT INTO settings VALUES(?,?)", ("failed_attempts", "0"))
            self.con.execute("INSERT INTO settings VALUES(?,?)", ("locked_until", "0"))

    def authenticate(self, password: str) -> bool:
        if not self.is_initialized():
            return False
        locked = float(self.get_setting("locked_until", "0"))
        import time
        if time.time() < locked:
            raise PermissionError(f"Временная блокировка. Повторите через {int(locked-time.time())+1} сек.")
        salt = bytes.fromhex(self.get_setting("pw_salt"))
        expected = bytes.fromhex(self.get_setting("pw_hash"))
        received = self._password_digest(password, salt)
        ok = hmac.compare_digest(expected, received)
        with self.con:
            if ok:
                self.con.execute("UPDATE settings SET value='0' WHERE key IN ('failed_attempts','locked_until')")
            else:
                failures = int(self.get_setting("failed_attempts", "0")) + 1
                self.con.execute("UPDATE settings SET value=? WHERE key='failed_attempts'", (str(failures),))
                if failures >= 5:
                    self.con.execute("UPDATE settings SET value=? WHERE key='locked_until'", (str(time.time()+60),))
                    self.con.execute("UPDATE settings SET value='0' WHERE key='failed_attempts'")
        return ok

    @staticmethod
    def _validate_recovery_key(recovery_key: str) -> None:
        if len(recovery_key) < 10:
            raise ValueError("Ключ восстановления должен быть длиной не менее 10 символов")

    def has_recovery_key(self) -> bool:
        salt = self.get_setting("recovery_salt")
        digest = self.get_setting("recovery_hash")
        if bool(salt) != bool(digest):
            raise RuntimeError("Настройки ключа восстановления повреждены")
        return bool(salt and digest)

    def set_recovery_key(self, recovery_key: str) -> None:
        self._validate_recovery_key(recovery_key)
        if self.has_recovery_key():
            raise ValueError("Ключ восстановления уже настроен")
        salt = secrets.token_bytes(16)
        digest = self._recovery_digest(recovery_key, salt)
        with self.con:
            self.con.execute("INSERT INTO settings VALUES(?,?)", ("recovery_salt", salt.hex()))
            self.con.execute("INSERT INTO settings VALUES(?,?)", ("recovery_hash", digest.hex()))

    def change_password(self, current_password: str, new_password: str) -> None:
        self._validate_password(new_password)
        if not self.authenticate(current_password):
            raise ValueError("Неверный текущий пароль")
        with self.con:
            self._store_password(new_password)
            self.con.execute("UPDATE settings SET value='0' WHERE key IN ('failed_attempts','locked_until')")

    def reset_password_with_recovery_key(self, recovery_key: str, new_password: str) -> None:
        self._validate_password(new_password)
        if not self.has_recovery_key():
            raise ValueError("Для этой базы ещё не настроен ключ восстановления")

        import time
        locked = float(self.get_setting("recovery_locked_until", "0"))
        if time.time() < locked:
            raise PermissionError(
                f"Временная блокировка. Повторите через {int(locked-time.time())+1} сек."
            )
        salt = bytes.fromhex(self.get_setting("recovery_salt"))
        expected = bytes.fromhex(self.get_setting("recovery_hash"))
        received = self._recovery_digest(recovery_key, salt)
        if not hmac.compare_digest(expected, received):
            failures = int(self.get_setting("recovery_failed_attempts", "0")) + 1
            with self.con:
                if failures >= 5:
                    self.con.execute(
                        "INSERT INTO settings(key,value) VALUES('recovery_locked_until',?) "
                        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                        (str(time.time() + 60),),
                    )
                    failures = 0
                self.con.execute(
                    "INSERT INTO settings(key,value) VALUES('recovery_failed_attempts',?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (str(failures),),
                )
            raise ValueError("Неверный ключ восстановления")

        with self.con:
            self._store_password(new_password)
            self.con.execute(
                "UPDATE settings SET value='0' "
                "WHERE key IN ('failed_attempts','locked_until','recovery_failed_attempts','recovery_locked_until')"
            )

    def _audit(self, action: str, record_id: str = "", detail: str = ""):
        self.con.execute("INSERT INTO audit_log(at,action,record_id,detail) VALUES(?,?,?,?)", (timestamp(), action, record_id, detail))

    def _record_transfers(self, student_id: str, new: dict, old: dict | None, source: str) -> None:
        for field, direction, event_date in (
            ("from_school", "in", new.get("admitted_on", "")),
            ("to_school", "out", new.get("left_on", "")),
        ):
            school = (new.get(field) or "").strip()
            if school and (old is None or school != (old.get(field) or "").strip()
                           or (event_date and event_date != (old.get('admitted_on' if direction=='in' else 'left_on') or ''))):
                self.con.execute('''INSERT INTO transfer_events
                    (student_id,direction,school_name,event_date,recorded_at,source)
                    VALUES(?,?,?,?,?,?)''',
                    (student_id, direction, school, event_date, timestamp(), source))
                self._audit("transfer_" + direction, student_id, school)

    def _insert_student(self, values: dict, record_id: str, now: str, source: str) -> None:
        cols = ",".join(COLUMNS) + ",grade_year,created_at,updated_at"
        marks = ",".join("?" for _ in range(len(COLUMNS) + 3))
        year = school_year()
        self.con.execute(f"INSERT INTO students({cols}) VALUES({marks})",
                         (record_id, *(values.get(c, "") for c in WRITABLE), year, now, now))
        self._record_transfers(record_id, values, None, source)
        self._audit("create" if source == "manual" else "import_create", record_id)

    def _update_student(self, values: dict, record_id: str, now: str, source: str) -> None:
        row = self.con.execute("SELECT * FROM students WHERE id=?", (record_id,)).fetchone()
        if not row or row["deleted_at"]:
            raise ValueError("Невозможно изменить отсутствующую или удалённую запись")
        old = dict(row)
        assignments = ",".join(f"{key}=?" for key in WRITABLE)
        grade_year = school_year() if old["grade"] != values["grade"] else old["grade_year"]
        self.con.execute(f"UPDATE students SET {assignments},grade_year=?,updated_at=? WHERE id=?",
                         (*(values.get(c, "") for c in WRITABLE), grade_year, now, record_id))
        if old["grade"] != values["grade"]:
            self.con.execute('''INSERT INTO grade_events
                (student_id,old_grade,new_grade,school_year,recorded_at,source)
                VALUES(?,?,?,?,?,?)''',
                (record_id, old["grade"], values["grade"], school_year(), now, source))
        self._record_transfers(record_id, values, old, source)
        self._audit("update" if source == "manual" else "import_update", record_id)

    def save_student(self, values: dict, record_id: str | None = None) -> str:
        from .validation import clean_record
        values = clean_record(values, allow_incomplete=True)
        now = timestamp()
        with self.con:
            if record_id:
                self._update_student(values, record_id, now, "manual")
                return record_id
            new_id = str(uuid.uuid4())
            self._insert_student(values, new_id, now, "manual")
            return new_id

    def transfer_history(self, record_id: str) -> list[dict]:
        rows = self.con.execute("SELECT * FROM transfer_events WHERE student_id=? ORDER BY id DESC", (record_id,)).fetchall()
        return [dict(row) for row in rows]

    def grade_history(self, record_id: str) -> list[dict]:
        rows = self.con.execute("SELECT * FROM grade_events WHERE student_id=? ORDER BY id DESC", (record_id,)).fetchall()
        return [dict(row) for row in rows]

    def final_grade(self) -> int:
        return int(self.get_setting('final_grade','11'))

    def set_final_grade(self, value: int) -> None:
        if value not in (11,12):raise ValueError('Выпускной класс: 11 или 12')
        self.set_setting('final_grade',str(value))

    def promote_for_school_year(self, today: date | None = None) -> int:
        '''One promotion for each missed academic year, only pupils still studying.
        Never automatically graduate or advance beyond the configured final grade.
        The operation is fully idempotent (including after restarts).
        '''
        target = school_year(today)
        ceiling=self.final_grade()
        rows = self.con.execute('''SELECT id,grade,grade_year FROM students
            WHERE deleted_at IS NULL AND status='Учится' AND grade<? AND grade_year<?''', (ceiling,target)).fetchall()
        count = 0
        with self.con:
            for row in rows:
                years = target - row["grade_year"]
                new_grade = min(ceiling, row["grade"] + years)
                if new_grade <= row["grade"]:
                    continue
                self.con.execute("UPDATE students SET grade=?,grade_year=?,updated_at=? WHERE id=?",
                                 (new_grade, target, timestamp(), row["id"]))
                self.con.execute('''INSERT INTO grade_events
                    (student_id,old_grade,new_grade,school_year,recorded_at,source)
                    VALUES(?,?,?,?,?,?)''',
                    (row["id"], row["grade"], new_grade, target, timestamp(), "automatic"))
                self._audit("auto_promotion", row["id"], f"{row['grade']} -> {new_grade}; {target}")
                count += 1
        return count

    def get_student(self, record_id: str):
        row = self.con.execute("SELECT * FROM students WHERE id=?", (record_id,)).fetchone()
        return dict(row) if row else None

    def soft_delete(self, record_id: str) -> bool:
        with self.con:
            cursor = self.con.execute("UPDATE students SET deleted_at=?,updated_at=? WHERE id=? AND deleted_at IS NULL", (timestamp(), timestamp(), record_id))
            if cursor.rowcount:
                self._audit("trash", record_id)
            return cursor.rowcount > 0

    def restore(self, record_id: str) -> bool:
        with self.con:
            cursor = self.con.execute("UPDATE students SET deleted_at=NULL, updated_at=? WHERE id=? AND deleted_at IS NOT NULL", (timestamp(), record_id))
            if cursor.rowcount:
                self._audit("restore", record_id)
            return cursor.rowcount > 0

    def list_students(self, *, search: str = "", grade: int | None = None, status: str | None = None, deleted: bool = False,
                      page: int = 0, page_size: int = 100) -> tuple[list[dict], int]:
        terms = ["deleted_at IS NOT NULL" if deleted else "deleted_at IS NULL"]
        args: list = []
        if search.strip():
            s = search.strip().casefold().replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            terms.append("(unicode_casefold(surname) LIKE ? ESCAPE '\\' OR unicode_casefold(given_name) LIKE ? ESCAPE '\\' OR unicode_casefold(patronymic) LIKE ? ESCAPE '\\' OR unicode_casefold(address) LIKE ? ESCAPE '\\' OR unicode_casefold(guardian_name) LIKE ? ESCAPE '\\')")
            args.extend([f"%{s}%"] * 5)
        if grade:
            terms.append("grade=?")
            args.append(grade)
        if status:
            if status not in ('Учится','Переведён','Выпущен'):
                raise ValueError('Недопустимый фильтр статуса')
            terms.append('status=?')
            args.append(status)
        where = " AND ".join(terms)
        total = self.con.execute(f"SELECT COUNT(*) FROM students WHERE {where}", args).fetchone()[0]
        rows = self.con.execute(f"SELECT * FROM students WHERE {where} ORDER BY surname COLLATE NOCASE,given_name COLLATE NOCASE,id LIMIT ? OFFSET ?", (*args, page_size, max(0,page) * page_size)).fetchall()
        return [dict(row) for row in rows], total

    def stats(self) -> dict:
        where = "deleted_at IS NULL"
        studying = self.con.execute(f"SELECT COUNT(*) FROM students WHERE {where} AND status='Учится'").fetchone()[0]
        transferred = self.con.execute(f"SELECT COUNT(*) FROM students WHERE {where} AND status='Переведён'").fetchone()[0]
        graduated = self.con.execute(f"SELECT COUNT(*) FROM students WHERE {where} AND status='Выпущен'").fetchone()[0]
        total = studying + transferred + graduated
        classes = self.con.execute(f"SELECT COUNT(DISTINCT grade||'-'||section) FROM students WHERE {where} AND status='Учится'").fetchone()[0]
        trashed = self.con.execute("SELECT COUNT(*) FROM students WHERE deleted_at IS NOT NULL").fetchone()[0]
        arriving = self.con.execute(f"SELECT COUNT(*) FROM students WHERE {where} AND from_school<>''").fetchone()[0]
        review = self.con.execute(f"SELECT COUNT(*) FROM students WHERE {where} AND status='Учится' AND grade>=? AND grade_year<?", (self.final_grade(),school_year())).fetchone()[0]
        return dict(active=total, studying=studying, transferred=transferred, graduated=graduated,
                    classes=classes, trashed=trashed, arriving=arriving, review=review)

    def grade_distribution(self) -> list[dict]:
        rows = self.con.execute('''SELECT grade,section,COUNT(*) AS total FROM students
            WHERE deleted_at IS NULL AND status='Учится'
            GROUP BY grade,section ORDER BY grade,section''').fetchall()
        return [dict(r) for r in rows]

    def apply_import(self, records: list[dict], mode: str = "skip") -> tuple[int, int]:
        '''Transactional import with safe legacy-format duplicate detection.
        Fingerprints of legacy rows remain valid after a year's promotion.
        '''
        from .validation import clean_record
        if mode not in ("skip", "update"):
            raise ValueError("Неизвестный режим")
        inserted = updated = 0
        now = timestamp()
        with self.con:
            for raw in records:
                record = clean_record(raw, allow_incomplete=True)
                fingerprint = raw.get("_fingerprint", "")
                if fingerprint:
                    seen = self.con.execute("SELECT student_id FROM import_fingerprints WHERE fingerprint=?", (fingerprint,)).fetchone()
                    if seen:
                        continue
                record_id = record.get("id") or str(uuid.uuid4())
                existing = self.con.execute("SELECT * FROM students WHERE id=?", (record_id,)).fetchone()
                if not existing and not record.get("id"):
                    # DOB is optional in the six-column Tajik template: never match solely on blank DOB.
                    if record["date_of_birth"]:
                        existing = self.con.execute('''SELECT * FROM students WHERE deleted_at IS NULL
                            AND unicode_casefold(surname)=? AND unicode_casefold(given_name)=?
                            AND date_of_birth=?''',
                            (record["surname"].casefold(), record["given_name"].casefold(), record["date_of_birth"])).fetchone()
                    # Never infer identity from a phone number alone. A shared family
                    # phone can belong to multiple children. Without a UUID or DOB,
                    # only an exact import fingerprint makes reruns idempotent.
                if existing:
                    if existing["deleted_at"]:
                        raise ValueError("Файл ссылается на запись в корзине. Сначала восстановите её.")
                    record_id = existing["id"]
                    if mode == "update" and record.get("id"):
                        if raw.get('_provided_fields'):
                            for key in WRITABLE:
                                if key not in raw['_provided_fields']:
                                    record[key]=existing[key]
                        self._update_student(record, record_id, now, "excel")
                        updated += 1
                else:
                    self._insert_student(record, record_id, now, "excel")
                    inserted += 1
                if fingerprint:
                    self.con.execute("INSERT INTO import_fingerprints(fingerprint,student_id,imported_at) VALUES(?,?,?)", (fingerprint, record_id, now))
        return inserted, updated

    def sqlite_snapshot(self, target: Path) -> None:
        """Online consistent snapshot, includes committed WAL transactions."""
        dest = sqlite3.connect(target)
        try:
            self.con.backup(dest)
            result = dest.execute("PRAGMA integrity_check").fetchone()[0]
            if result != "ok":
                raise RuntimeError("Повреждение копии базы")
        finally:
            dest.close()

    def replace_from_snapshot(self, snapshot: Path) -> None:
        source = sqlite3.connect(snapshot)
        try:
            if source.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                raise ValueError("Некорректная база")
            tables = {r[0] for r in source.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if not {"students", "settings", "audit_log"}.issubset(tables):
                raise ValueError("Это не резервная копия MaktabOS")
            source.backup(self.con)
            self.con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        finally:
            source.close()
