import tempfile
import unittest
import sqlite3
from pathlib import Path

from schoolregister.recovery import recover_local_account
from schoolregister.storage import Database


class LocalAccountRecoveryTests(unittest.TestCase):
    def test_recovery_creates_backup_preserves_records_and_changes_credentials(self):
        with tempfile.TemporaryDirectory() as temporary:
            database_path = Path(temporary) / "school.db"
            db = Database(database_path)
            try:
                db.set_password_initial("old-password-123", "old-recovery-key")
                student_id = db.save_student({
                    "surname": "Тестов", "given_name": "Алишер", "patronymic": "",
                    "date_of_birth": "2014-04-12", "grade": 6, "section": "Б",
                    "shift": "1", "address": "Душанбе", "guardian_name": "",
                    "guardian_phone": "", "extra_phone": "", "admitted_on": "",
                    "status": "Учится", "notes": "", "from_school": "", "to_school": "",
                    "left_on": "",
                })
            finally:
                db.close()

            backup_path = recover_local_account(
                database_path,
                "new-password-456",
                "new-recovery-key",
            )

            self.assertTrue(backup_path.is_file())
            self.assertNotEqual(backup_path, database_path)
            migrated = Database(database_path)
            try:
                self.assertTrue(migrated.authenticate("new-password-456"))
                self.assertTrue(migrated.has_recovery_key())
                self.assertEqual(migrated.get_student(student_id)["surname"], "Тестов")
                self.assertNotIn(
                    "new-recovery-key",
                    database_path.read_bytes().decode("latin1"),
                )
            finally:
                migrated.close()

    def test_recovery_rejects_unrecognized_database_without_changes(self):
        with tempfile.TemporaryDirectory() as temporary:
            database_path = Path(temporary) / "other.db"
            database_path.write_bytes(b"not a MaktabOS database")
            with self.assertRaises(sqlite3.DatabaseError):
                recover_local_account(database_path, "new-password-456", "new-recovery-key")
            self.assertEqual(database_path.read_bytes(), b"not a MaktabOS database")


if __name__ == "__main__":
    unittest.main()
