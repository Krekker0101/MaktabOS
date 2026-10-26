"""Small GUI to recover a local MaktabOS account after forgetting its credentials."""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from schoolregister.recovery import recover_local_account
from schoolregister.storage import data_directory


class RecoveryWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MaktabOS — восстановление доступа")
        self.setMinimumWidth(570)

        root = QWidget()
        layout = QVBoxLayout(root)
        self.setCentralWidget(root)

        warning = QLabel(
            "Закройте MaktabOS перед восстановлением. Утилита сбрасывает пароль "
            "выбранной локальной базы без старого пароля или ключа. Используйте её "
            "только для базы, к которой вы уполномочены получать доступ. Перед "
            "изменением будет создана резервная копия."
        )
        warning.setWordWrap(True)
        layout.addWidget(warning)

        db_row = QHBoxLayout()
        default_database = data_directory() / "school.db"
        self.database_path = QLineEdit(str(default_database))
        self.database_path.setReadOnly(True)
        db_row.addWidget(self.database_path, 1)
        choose = QPushButton("Выбрать базу…")
        choose.clicked.connect(self.choose_database)
        db_row.addWidget(choose)
        layout.addLayout(db_row)
        self.database_status = QLabel()
        self.database_status.setWordWrap(True)
        layout.addWidget(self.database_status)
        self.update_database_status()

        form = QFormLayout()
        self.password = self._password_field("Новый пароль")
        self.password_confirm = self._password_field("Повторите пароль")
        self.recovery_key = self._password_field("Новый ключ восстановления")
        self.recovery_confirm = self._password_field("Повторите ключ")
        form.addRow("Новый пароль:", self.password)
        form.addRow("Подтверждение:", self.password_confirm)
        form.addRow("Ключ восстановления:", self.recovery_key)
        form.addRow("Подтверждение:", self.recovery_confirm)
        layout.addLayout(form)

        self.confirm_authority = QCheckBox(
            "Подтверждаю, что я владелец базы или уполномочен восстановить доступ"
        )
        layout.addWidget(self.confirm_authority)

        actions = QHBoxLayout()
        actions.addStretch()
        cancel = QPushButton("Закрыть")
        cancel.clicked.connect(self.close)
        actions.addWidget(cancel)
        recover = QPushButton("Создать копию и сбросить пароль")
        recover.clicked.connect(self.recover)
        actions.addWidget(recover)
        layout.addLayout(actions)

    @staticmethod
    def _password_field(placeholder: str) -> QLineEdit:
        field = QLineEdit()
        field.setEchoMode(QLineEdit.EchoMode.Password)
        field.setPlaceholderText(placeholder + " (не менее 10 символов)")
        return field

    def choose_database(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите базу MaktabOS",
            str(Path(self.database_path.text()).parent),
            "База SQLite (*.db);;Все файлы (*)",
        )
        if filename:
            self.database_path.setText(filename)
            self.update_database_status()

    def update_database_status(self) -> None:
        path = Path(self.database_path.text())
        if path.is_file():
            self.database_status.setText(f"База найдена автоматически: {path}")
        else:
            self.database_status.setText(
                f"Стандартная база не найдена: {path}\n"
                "Если база находится в другом месте, нажмите «Выбрать базу…»."
            )

    def recover(self) -> None:
        if not self.confirm_authority.isChecked():
            QMessageBox.warning(self, "Подтверждение", "Подтвердите право на восстановление этой базы.")
            return
        if self.password.text() != self.password_confirm.text():
            QMessageBox.warning(self, "Новый пароль", "Пароли не совпадают.")
            return
        if self.recovery_key.text() != self.recovery_confirm.text():
            QMessageBox.warning(self, "Ключ восстановления", "Ключи не совпадают.")
            return

        answer = QMessageBox.question(
            self,
            "Подтвердите сброс",
            "Будет создана резервная копия, затем изменятся данные входа. "
            "Карточки учеников сохранятся. Продолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        try:
            backup_path = recover_local_account(
                Path(self.database_path.text()),
                self.password.text(),
                self.recovery_key.text(),
            )
        except (OSError, ValueError, RuntimeError) as exc:
            QMessageBox.critical(self, "Сброс не выполнен", str(exc))
            return
        except sqlite3.Error as exc:
            QMessageBox.critical(self, "Ошибка базы", f"Не удалось безопасно завершить сброс:\n{exc}")
            return

        QMessageBox.information(
            self,
            "Доступ восстановлен",
            "Пароль обновлён, ключ восстановления сохранён в виде хеша.\n"
            f"Резервная копия:\n{backup_path}\n\n"
            "Теперь запустите MaktabOS и войдите с новым паролем.",
        )
        self.close()


def main() -> int:
    app = QApplication(sys.argv)
    window = RecoveryWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
