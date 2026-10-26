import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from schoolregister.storage import Database, data_directory
from schoolregister.backup import create_backup, restore_backup
from schoolregister.excel import export_excel, parse_excel, asset_template
from schoolregister.validation import clean_record, StudentValidationError


def fixture(**kw):
    row = dict(surname="Тестов",given_name="Алишер",patronymic="",
        date_of_birth="2014-04-12",grade=6,section="Б",shift="1",
        address="Душанбе, примерная 12",guardian_name="Тестова А.",
        guardian_phone="+992900001122",extra_phone="",admitted_on="2020-09-01",
        status="Учится",notes="")
    row.update(kw)
    return clean_record(row)


class SchoolRegisterTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.dir=Path(self.temp.name)
        self.db=Database(self.dir/"persistent"/"school.db")
        self.addCleanup(self.db.close)
        self.password="correct-horse-staple"
        self.db.set_password_initial(self.password)

    def test_login_lockout_and_hash(self):
        self.assertTrue(self.db.authenticate(self.password))
        self.assertFalse(self.db.authenticate("invalid"))
        self.assertNotIn(self.password, self.db.path.read_bytes().decode("latin1"))
        for _ in range(4):self.db.authenticate("invalid")
        with self.assertRaises(PermissionError):self.db.authenticate(self.password)

    def test_change_and_recover_password_with_hashed_key(self):
        recovery_key="local-recovery-key"
        changed_password="new-password-123"
        recovered_password="recovered-password-456"
        self.db.set_recovery_key(recovery_key)
        with self.assertRaises(ValueError):
            self.db.change_password("incorrect-password",changed_password)
        self.assertTrue(self.db.authenticate(self.password))
        self.db.change_password(self.password,changed_password)
        self.assertTrue(self.db.authenticate(changed_password))
        self.assertFalse(self.db.authenticate(self.password))
        with self.assertRaises(ValueError):
            self.db.reset_password_with_recovery_key("incorrect-key",recovered_password)
        self.db.reset_password_with_recovery_key(recovery_key,recovered_password)
        self.assertTrue(self.db.authenticate(recovered_password))
        self.assertNotIn(recovery_key,self.db.path.read_bytes().decode("latin1"))

    def test_initial_password_and_recovery_key_are_set_together(self):
        db=Database(self.dir/"new-account"/"school.db")
        self.addCleanup(db.close)
        with self.assertRaises(ValueError):
            db.set_password_initial(self.password,"short")
        self.assertFalse(db.is_initialized())
        recovery_key="shared-recovery-key"
        db.set_password_initial(self.password,recovery_key)
        self.assertTrue(db.has_recovery_key())
        db.reset_password_with_recovery_key(recovery_key,"replacement-password-789")
        self.assertTrue(db.authenticate("replacement-password-789"))

    def test_save_edit_search_and_restore(self):
        rid=self.db.save_student(fixture())
        item=self.db.get_student(rid)
        self.assertEqual(item["grade"],6)
        self.assertEqual(self.db.list_students(search="Тест",page=0)[1],1)
        self.assertEqual(self.db.list_students(search="тЕсТ",page=0)[1],1)
        self.assertEqual(self.db.list_students(search="%",page=0)[1],0)
        self.db.save_student(fixture(address="Новый адрес"),rid)
        self.assertEqual(self.db.get_student(rid)["address"],"Новый адрес")
        self.assertTrue(self.db.soft_delete(rid))
        self.assertEqual(self.db.list_students()[1],0)
        self.assertEqual(self.db.list_students(deleted=True)[1],1)
        self.assertTrue(self.db.restore(rid))
        self.assertEqual(self.db.list_students()[1],1)

    def test_password_and_persistence_outside_install_dir(self):
        app_dir=self.dir/"application"; app_dir.mkdir()
        with patch.dict(os.environ,{"SCHOOLREGISTER_DATA_DIR":str(self.dir/"stable-data")}):
            self.assertEqual(data_directory(), self.dir/"stable-data")
            other=Database(data_directory()/"school.db")
            other.set_password_initial(self.password)
            other.save_student(fixture())
            other.close()
            (app_dir/"app.exe").write_text("fake")
            (app_dir/"app.exe").unlink()
            reinstalled=Database(data_directory()/"school.db")
            try:self.assertEqual(reinstalled.list_students()[1],1)
            finally:reinstalled.close()

    def test_legacy_data_is_copied_to_maktabos_directory(self):
        base=self.dir/"local-app-data"
        legacy=Database(base/"SchoolRegister"/"school.db")
        legacy.save_student(fixture())
        legacy.close()

        with patch.dict(os.environ,{"LOCALAPPDATA":str(base)},clear=True), \
             patch("schoolregister.storage.sys.platform","win32"):
            new_directory=data_directory()

        self.assertEqual(new_directory,base/"MaktabOS")
        self.assertTrue((base/"SchoolRegister"/"school.db").is_file())
        migrated=Database(new_directory/"school.db")
        try:self.assertEqual(migrated.list_students()[1],1)
        finally:migrated.close()

    def test_xlsx_export_round_trip_and_duplicate_skip(self):
        rid=self.db.save_student(fixture())
        target=self.dir/"students.xlsx"
        export_excel(self.db,target)
        preview=parse_excel(target)
        self.assertFalse(preview.errors)
        self.assertEqual(len(preview.records),1)
        self.assertEqual(preview.records[0]["id"],rid)
        self.assertEqual(self.db.apply_import(preview.records, "skip"),(0,0))
        edited=preview.records[0].copy()
        edited["address"]="Изменён импортом"
        self.assertEqual(self.db.apply_import([edited], "update"),(0,1))
        self.assertEqual(self.db.get_student(rid)["address"],"Изменён импортом")

    def test_official_template_import_and_no_duplicates(self):
        preview=parse_excel(asset_template())
        self.assertEqual(preview.records,[])
        new_record=fixture()
        self.assertEqual(self.db.apply_import([new_record]),(1,0))
        self.assertEqual(self.db.apply_import([new_record]),(0,0))
        self.assertEqual(self.db.list_students()[1],1)

    def test_import_is_atomic(self):
        existing=self.db.save_student(fixture())
        records=[fixture(surname="Новый"),fixture(id=existing,surname="Заменён")]
        self.db.soft_delete(existing)
        with self.assertRaises(ValueError):self.db.apply_import(records,mode="update")
        self.assertEqual(self.db.list_students()[1],0)

    def test_excel_formulas_rejected(self):
        from openpyxl import Workbook
        from schoolregister.constants import HEADERS
        wb=Workbook();ws=wb.active;ws.title="Ученики";ws.append(HEADERS)
        ws.append(["", "=1+1", "Мария", "", "2014-01-02", 4, "А", "", "Адрес"])
        p=self.dir/"unsafe.xlsx";wb.save(p)
        preview=parse_excel(p)
        self.assertIn("формулы запрещены", preview.errors[0])

    def test_validation_fails_bad_data(self):
        for bad in [dict(surname=""),dict(date_of_birth="2099-04-01"),dict(grade=13),dict(guardian_phone="bad phone")]:
            with self.assertRaises(StudentValidationError):fixture(**bad)

    def test_encrypted_backup_restore_and_wrong_password(self):
        self.db.save_student(fixture())
        backup=self.dir/"snapshot.srbak"
        create_backup(self.db,backup,self.password)
        blob=backup.read_bytes()
        self.assertNotIn("Тестов".encode(),blob)
        with self.assertRaises(ValueError):restore_backup(self.db,backup,"wrongpass123")
        self.db.save_student(fixture(surname="Второй"))
        self.assertEqual(self.db.list_students()[1],2)
        restore_backup(self.db,backup,self.password)
        self.assertEqual(self.db.list_students()[1],1)
        self.assertTrue(self.db.authenticate(self.password))


if __name__=='__main__':unittest.main()
