"""Password-encrypted snapshots. No plaintext backup persists on disk after completion."""
from __future__ import annotations
import hashlib
import hmac
import os
import sqlite3
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from .storage import Database, _secure_dir, _secure_file

MAGIC = b"SRBAK1\x00"


def _key(password: str, salt: bytes) -> bytes:
    return hashlib.scrypt(password.encode("utf-8"), salt=salt, n=2**14, r=8, p=1, dklen=32)


def create_backup(db: Database, target: Path, password: str) -> Path:
    if not password:
        raise ValueError("Требуется пароль")
    target = Path(target)
    _secure_dir(target.parent)
    temp = None
    try:
        with tempfile.NamedTemporaryFile(prefix="sr-snapshot-",suffix=".sqlite", delete=False) as f:
            temp = Path(f.name)
        _secure_file(temp)
        db.sqlite_snapshot(temp)
        payload = temp.read_bytes()
        salt, nonce = os.urandom(16), os.urandom(12)
        data = MAGIC + salt + nonce + AESGCM(_key(password, salt)).encrypt(nonce,payload,MAGIC)
        # Temp file and atomic replacement, preserving old backups if write fails.
        with tempfile.NamedTemporaryFile(dir=target.parent,prefix=".writing-",suffix=".tmp",delete=False) as f:
            stage = Path(f.name)
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        _secure_file(stage)
        os.replace(stage,target)
        _secure_file(target)
        return target
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)


def _decrypt(source: Path, password: str) -> bytes:
    source = Path(source)
    if source.stat().st_size > 256 * 1024 * 1024:
        raise ValueError("Резервная копия слишком велика")
    data = source.read_bytes()
    if not data.startswith(MAGIC) or len(data) < len(MAGIC)+16+12+16:
        raise ValueError("Неверный формат резервной копии")
    offset=len(MAGIC)
    salt=data[offset:offset+16]
    nonce=data[offset+16:offset+28]
    try:
        return AESGCM(_key(password,salt)).decrypt(nonce,data[offset+28:],MAGIC)
    except InvalidTag as exc:
        raise ValueError("Неверный пароль или повреждена резервная копия") from exc


def restore_backup(db: Database, source: Path, password: str) -> None:
    payload = _decrypt(source,password)
    temp = None
    try:
        with tempfile.NamedTemporaryFile(prefix="sr-restore-",suffix=".sqlite",delete=False) as f:
            temp = Path(f.name)
            f.write(payload)
        _secure_file(temp)
        con = sqlite3.connect(f"file:{temp}?mode=ro",uri=True)
        try:
            valid = con.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
            settings = dict(con.execute("SELECT key,value FROM settings WHERE key IN ('pw_hash','pw_salt')"))
            salt=bytes.fromhex(settings["pw_salt"])
            actual=hashlib.pbkdf2_hmac("sha256",password.encode("utf-8"),salt,650_000)
            good=hmac.compare_digest(actual,bytes.fromhex(settings["pw_hash"]))
        except (sqlite3.Error,KeyError,ValueError) as exc:
            raise ValueError("Недопустимая структура резервной копии") from exc
        finally:
            con.close()
        if not valid or not good:
            raise ValueError("Резервная копия не прошла проверку пароля и целостности")
        db.replace_from_snapshot(temp)
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)


def suggested_backup_name():
    return "maktabos_" + datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%SZ") + ".srbak"
