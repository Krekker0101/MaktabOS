"""Owner-operated recovery for a local MaktabOS database."""
from __future__ import annotations

import hashlib
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path

from .storage import PASSWORD_ITERATIONS


def _digest(value: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", value.encode("utf-8"), salt, PASSWORD_ITERATIONS)


def _upsert_setting(connection: sqlite3.Connection, key: str, value: str) -> None:
    connection.execute(
        "INSERT INTO settings(key,value) VALUES(?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )


def recover_local_account(
    database_path: Path,
    new_password: str,
    new_recovery_key: str,
) -> Path:
    """Back up a MaktabOS database, then replace its login and recovery credentials."""
    database_path = Path(database_path).expanduser().resolve()
    if len(new_password) < 10:
        raise ValueError("Новый пароль должен быть длиной не менее 10 символов")
    if len(new_recovery_key) < 10:
        raise ValueError("Новый ключ восстановления должен быть длиной не менее 10 символов")
    if not database_path.is_file():
        raise FileNotFoundError(f"Файл базы не найден: {database_path}")

    connection = sqlite3.connect(database_path, timeout=20)
    try:
        has_settings = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='settings'"
        ).fetchone()
        has_students = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='students'"
        ).fetchone()
        if not has_settings or not has_students:
            raise ValueError("Выбранный файл не похож на базу MaktabOS")

        settings = dict(connection.execute("SELECT key,value FROM settings"))
        if not settings.get("pw_hash") or not settings.get("pw_salt"):
            raise ValueError("В выбранной базе нет настроенного пароля для сброса")

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_path = database_path.with_name(
            f"{database_path.stem}.before-recovery-{stamp}{database_path.suffix}"
        )
        counter = 1
        while backup_path.exists():
            backup_path = database_path.with_name(
                f"{database_path.stem}.before-recovery-{stamp}-{counter}{database_path.suffix}"
            )
            counter += 1

        backup = sqlite3.connect(backup_path)
        try:
            connection.backup(backup)
            integrity = backup.execute("PRAGMA integrity_check").fetchone()
            if not integrity or integrity[0] != "ok":
                raise sqlite3.DatabaseError("Проверка целостности копии базы не пройдена")
        except (OSError, sqlite3.Error):
            backup.close()
            backup_path.unlink(missing_ok=True)
            raise
        else:
            backup.close()

        password_salt = secrets.token_bytes(16)
        recovery_salt = secrets.token_bytes(16)
        with connection:
            _upsert_setting(connection, "pw_salt", password_salt.hex())
            _upsert_setting(connection, "pw_hash", _digest(new_password, password_salt).hex())
            _upsert_setting(connection, "recovery_salt", recovery_salt.hex())
            _upsert_setting(
                connection,
                "recovery_hash",
                _digest(new_recovery_key, recovery_salt).hex(),
            )
            _upsert_setting(connection, "failed_attempts", "0")
            _upsert_setting(connection, "locked_until", "0")
            _upsert_setting(connection, "recovery_failed_attempts", "0")
            _upsert_setting(connection, "recovery_locked_until", "0")
        return backup_path
    finally:
        connection.close()
