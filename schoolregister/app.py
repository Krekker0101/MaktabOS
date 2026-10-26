"""Qt 6 desktop application: no network, no remote services."""
from __future__ import annotations
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtGui import QFont, QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDateEdit, QDialog, QFileDialog,
    QFormLayout, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPushButton, QScrollArea, QSpinBox,
    QStackedWidget, QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout,
    QWidget, QAbstractItemView, QSizePolicy, QMenu
)

from .backup import create_backup, restore_backup, suggested_backup_name
from .constants import APP_NAME, STATUSES
from .excel import export_excel, parse_file, save_template
from .word import export_word
from .storage import Database, data_directory
from .validation import clean_record, StudentValidationError
from . import __version__

THEME = """
QWidget { background: #F5F7FB; color: #19253B; font-family: 'Segoe UI', 'Inter', 'Arial'; font-size: 13px; }
QMainWindow { background: #F5F7FB; }
QLabel { background: transparent; }
QFrame#Card, QFrame#TableCard, QFrame#ToolbarCard { background: #FFFFFF; border: 1px solid #E5EAF2; border-radius: 14px; }
QFrame#Sidebar { background: #111C30; border: none; border-radius: 0; }
QFrame#Sidebar QLabel { color: #FFFFFF; }
QLabel#Brand { font-weight: 800; font-size: 22px; color: white; }
QLabel#PageTitle { font-weight: 800; font-size: 28px; color: #182740; }
QLabel#Muted { color: #64748B; font-size: 12px; }
QLabel#Metric { font-size: 27px; font-weight: 800; color: #182740; }
QPushButton { background: #FFFFFF; color: #243652; padding: 10px 15px; border: 1px solid #DAE1ED; border-radius: 9px; font-weight: 600; }
QPushButton:hover { border-color: #A8C1ED; background: #F0F5FF; }
QPushButton:pressed { background: #E4EDFC; }
QPushButton#Primary { background: #2563EB; color: white; border-color: #2563EB; }
QPushButton#Primary:hover { background: #1D4ED8; }
QPushButton#Danger { color: #C24141; background: #FFF7F7; border-color: #F2D5D5; }
QPushButton#Nav { text-align: left; background: transparent; color: #AAB9D4; border: none; padding: 13px 14px; font-size: 13px; font-weight: 600; }
QPushButton#Nav:hover { background: #1D3151; color: white; }
QPushButton#Nav[active="true"] { color: #FFFFFF; background: #264575; border-left: 3px solid #80AFFD; }
QPushButton:disabled { color: #9BA7B8; background: #F0F2F5; border-color: #E1E5EC; }
QLineEdit, QComboBox, QDateEdit, QSpinBox, QTextEdit { background: white; color: #1E293B; border: 1px solid #D9E2EE; border-radius: 8px; padding: 9px 11px; selection-background-color: #BFDBFE; }
QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus, QTextEdit:focus { border: 1px solid #2563EB; }
QComboBox { min-width: 130px; }
QTableWidget { background: white; color: #243247; gridline-color: #F1F3F7; border: none; border-radius: 9px; selection-background-color: #EAF2FF; selection-color: #132B55; alternate-background-color: #F8FAFD; }
QTableWidget::item { padding: 8px; }
QHeaderView::section { background: #F0F5FA; color: #53647F; font-size: 12px; font-weight: 700; border: none; border-bottom: 1px solid #E1E8F0; padding: 12px 8px; }
QScrollBar:vertical { width: 11px; background: transparent; }
QScrollBar::handle:vertical { background: #CBD5E1; border-radius: 5px; min-height: 35px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QDialog { background: #F8FAFD; }
QMessageBox { background: #FFFFFF; }
"""


def btn(text: str, action=None, style: str = "") -> QPushButton:
    b=QPushButton(text)
    if style:
        b.setObjectName(style)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    if action:
        b.clicked.connect(action)
    return b


def label(text: str, style: str = "") -> QLabel:
    v=QLabel(text)
    if style: v.setObjectName(style)
    return v


def message(parent, title, text, kind="info"):
    if kind == "error": QMessageBox.critical(parent,title,text)
    elif kind == "warning": QMessageBox.warning(parent,title,text)
    else: QMessageBox.information(parent,title,text)


class LoginDialog(QDialog):
    def __init__(self, db: Database):
        super().__init__()
        self.db=db
        self.password=None
        self.first=not db.is_initialized()
        self.setWindowTitle(f"Первый запуск · {APP_NAME}" if self.first else f"Вход · {APP_NAME}")
        self.resize(475,520 if self.first else 330)
        root=QVBoxLayout(self); root.setContentsMargins(34,32,34,30); root.setSpacing(14)
        head=label("Школьный реестр"); head.setStyleSheet("font-size:26px;font-weight:800;color:#173D77")
        root.addWidget(head)
        desc=label(
            "Создайте пароль и ключ восстановления (каждый не менее 10 символов)."
            if self.first else "Введите пароль для доступа к сведениям об учениках.",
            "Muted",
        )
        desc.setWordWrap(True); root.addWidget(desc)
        self.pass1=QLineEdit(); self.pass1.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass1.setPlaceholderText("Пароль администратора"); root.addWidget(self.pass1)
        self.pass2=None
        self.recovery1=None
        self.recovery2=None
        if self.first:
            self.pass2=QLineEdit(); self.pass2.setEchoMode(QLineEdit.EchoMode.Password)
            self.pass2.setPlaceholderText("Повторите пароль"); root.addWidget(self.pass2)
            self.recovery1=QLineEdit(); self.recovery1.setEchoMode(QLineEdit.EchoMode.Password)
            self.recovery1.setPlaceholderText("Ключ восстановления"); root.addWidget(self.recovery1)
            self.recovery2=QLineEdit(); self.recovery2.setEchoMode(QLineEdit.EchoMode.Password)
            self.recovery2.setPlaceholderText("Повторите ключ восстановления"); root.addWidget(self.recovery2)
        self.showpass=QCheckBox("Показать пароль"); root.addWidget(self.showpass)
        self.showpass.toggled.connect(self.toggle_password)
        self.note=label(
            "Ключ восстановления хранится только в виде защищённой проверки. "
            "Для другой базы на другом устройстве задайте тот же ключ отдельно.",
            "Muted",
        )
        self.note.setWordWrap(True); root.addWidget(self.note)
        enter=btn("Создать защищённый вход" if self.first else "Войти",self.login,"Primary")
        root.addWidget(enter)
        if not self.first:
            root.addWidget(btn("Забыли пароль?",self.recover))
        root.addStretch()
        self.pass1.returnPressed.connect(self.login)
        if self.pass2: self.pass2.returnPressed.connect(self.login)
        if self.recovery2: self.recovery2.returnPressed.connect(self.login)

    def toggle_password(self,show):
        mode=QLineEdit.EchoMode.Normal if show else QLineEdit.EchoMode.Password
        self.pass1.setEchoMode(mode)
        if self.pass2: self.pass2.setEchoMode(mode)
        if self.recovery1: self.recovery1.setEchoMode(mode)
        if self.recovery2: self.recovery2.setEchoMode(mode)

    def login(self):
        pw=self.pass1.text()
        try:
            if self.first:
                if not self.pass2 or pw != self.pass2.text():
                    raise ValueError("Пароли не совпадают")
                if (not self.recovery1 or not self.recovery2
                        or self.recovery1.text() != self.recovery2.text()):
                    raise ValueError("Ключи восстановления не совпадают")
                self.db.set_password_initial(pw, self.recovery1.text())
            elif not self.db.authenticate(pw):
                raise ValueError("Неверный пароль")
            if not self.first and not self.db.has_recovery_key():
                setup=RecoveryKeySetupDialog(self)
                if setup.exec()==QDialog.DialogCode.Accepted:
                    self.db.set_recovery_key(setup.recovery_key)
        except (ValueError,PermissionError,RuntimeError) as exc:
            message(self,"Вход",str(exc),"warning")
            return
        self.password=pw
        self.accept()


    def recover(self):
        try:
            if not self.db.has_recovery_key():
                message(self,"Восстановление пароля",
                        "Ключ восстановления ещё не настроен. Войдите с текущим паролем, "
                        "чтобы добавить его.","warning")
                return
            dialog=PasswordRecoveryDialog(self)
            if dialog.exec()!=QDialog.DialogCode.Accepted:
                return
            self.db.reset_password_with_recovery_key(dialog.recovery_key,dialog.new_password)
            self.pass1.clear()
            message(self,"Пароль изменён","Теперь войдите с новым паролем.")
        except (ValueError,PermissionError,RuntimeError) as exc:
            message(self,"Восстановление пароля",str(exc),"warning")


class RecoveryKeySetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.recovery_key=""
        self.setWindowTitle("Настройка ключа восстановления")
        self.resize(460,280)
        root=QVBoxLayout(self)
        note=label(
            "Введите один ключ и сохраните его в надёжном месте. "
            "Он позволит задать новый пароль без старого. Не задавайте ключ, "
            "который нельзя будет восстановить.",
            "Muted",
        )
        note.setWordWrap(True);root.addWidget(note)
        self.key1=QLineEdit();self.key1.setEchoMode(QLineEdit.EchoMode.Password)
        self.key1.setPlaceholderText("Ключ восстановления (не менее 10 символов)")
        root.addWidget(self.key1)
        self.key2=QLineEdit();self.key2.setEchoMode(QLineEdit.EchoMode.Password)
        self.key2.setPlaceholderText("Повторите ключ")
        root.addWidget(self.key2)
        self.show_key=QCheckBox("Показать ключ");root.addWidget(self.show_key)
        self.show_key.toggled.connect(self.toggle_key)
        actions=QHBoxLayout();actions.addStretch()
        actions.addWidget(btn("Отмена",self.reject))
        actions.addWidget(btn("Сохранить ключ",self.validate_and_accept,"Primary"))
        root.addLayout(actions)
        self.key2.returnPressed.connect(self.validate_and_accept)

    def toggle_key(self,show):
        mode=QLineEdit.EchoMode.Normal if show else QLineEdit.EchoMode.Password
        self.key1.setEchoMode(mode);self.key2.setEchoMode(mode)

    def validate_and_accept(self):
        key=self.key1.text()
        if len(key)<10:
            message(self,"Ключ восстановления","Ключ должен быть длиной не менее 10 символов.","warning")
        elif key!=self.key2.text():
            message(self,"Ключ восстановления","Ключи не совпадают.","warning")
        else:
            self.recovery_key=key
            self.accept()


class PasswordRecoveryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.recovery_key=""
        self.new_password=""
        self.setWindowTitle("Восстановление пароля")
        self.resize(440,300)
        root=QVBoxLayout(self)
        root.addWidget(label("Введите ключ восстановления, чтобы создать новый пароль.","Muted"))
        self.key=QLineEdit();self.key.setEchoMode(QLineEdit.EchoMode.Password)
        self.key.setPlaceholderText("Ключ восстановления");root.addWidget(self.key)
        self.password=QLineEdit();self.password.setEchoMode(QLineEdit.EchoMode.Password)
        self.password.setPlaceholderText("Новый пароль (не менее 10 символов)");root.addWidget(self.password)
        self.confirm=QLineEdit();self.confirm.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm.setPlaceholderText("Повторите новый пароль");root.addWidget(self.confirm)
        self.show_password=QCheckBox("Показать введённое");root.addWidget(self.show_password)
        self.show_password.toggled.connect(self.toggle_password)
        actions=QHBoxLayout();actions.addStretch()
        actions.addWidget(btn("Отмена",self.reject))
        actions.addWidget(btn("Изменить пароль",self.validate_and_accept,"Primary"))
        root.addLayout(actions)
        self.confirm.returnPressed.connect(self.validate_and_accept)

    def toggle_password(self,show):
        mode=QLineEdit.EchoMode.Normal if show else QLineEdit.EchoMode.Password
        self.key.setEchoMode(mode);self.password.setEchoMode(mode);self.confirm.setEchoMode(mode)

    def validate_and_accept(self):
        if len(self.password.text())<10:
            message(self,"Новый пароль","Пароль должен быть длиной не менее 10 символов.","warning")
        elif self.password.text()!=self.confirm.text():
            message(self,"Новый пароль","Пароли не совпадают.","warning")
        else:
            self.recovery_key=self.key.text()
            self.new_password=self.password.text()
            self.accept()


class PasswordChangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_password=""
        self.new_password=""
        self.setWindowTitle("Изменение пароля")
        self.resize(430,270)
        root=QVBoxLayout(self)
        self.current=QLineEdit();self.current.setEchoMode(QLineEdit.EchoMode.Password)
        self.current.setPlaceholderText("Текущий пароль");root.addWidget(self.current)
        self.password=QLineEdit();self.password.setEchoMode(QLineEdit.EchoMode.Password)
        self.password.setPlaceholderText("Новый пароль (не менее 10 символов)");root.addWidget(self.password)
        self.confirm=QLineEdit();self.confirm.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm.setPlaceholderText("Повторите новый пароль");root.addWidget(self.confirm)
        self.show_password=QCheckBox("Показать введённое");root.addWidget(self.show_password)
        self.show_password.toggled.connect(self.toggle_password)
        actions=QHBoxLayout();actions.addStretch()
        actions.addWidget(btn("Отмена",self.reject))
        actions.addWidget(btn("Изменить пароль",self.validate_and_accept,"Primary"))
        root.addLayout(actions)
        self.confirm.returnPressed.connect(self.validate_and_accept)

    def toggle_password(self,show):
        mode=QLineEdit.EchoMode.Normal if show else QLineEdit.EchoMode.Password
        self.current.setEchoMode(mode);self.password.setEchoMode(mode);self.confirm.setEchoMode(mode)

    def validate_and_accept(self):
        if len(self.password.text())<10:
            message(self,"Новый пароль","Пароль должен быть длиной не менее 10 символов.","warning")
        elif self.password.text()!=self.confirm.text():
            message(self,"Новый пароль","Пароли не совпадают.","warning")
        else:
            self.current_password=self.current.text()
            self.new_password=self.password.text()
            self.accept()


class StudentDialog(QDialog):
    def __init__(self, initial: dict | None = None, parent=None):
        super().__init__(parent)
        self.initial=initial or {}
        self.result_record=None
        self.setWindowTitle("Изменение ученика" if initial else "Добавление ученика")
        self.resize(760,770)
        main=QVBoxLayout(self); main.setContentsMargins(24,20,24,20)
        heading=label("Карточка ученика")
        heading.setStyleSheet("font-size:22px;font-weight:800;color:#173D77")
        main.addWidget(heading)
        sub=label("Обязательны фамилия, имя и класс. Остальные сведения можно добавить позже.","Muted")
        main.addWidget(sub)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.Shape.NoFrame)
        body=QWidget(); form=QFormLayout(body); form.setLabelAlignment(Qt.AlignmentFlag.AlignRight); form.setHorizontalSpacing(18); form.setVerticalSpacing(10)
        scroll.setWidget(body); main.addWidget(scroll)
        self.entries={}
        for key,title,ph in [
            ("surname","Фамилия *","Иванов"),("given_name","Имя *","Иван"),
            ("patronymic","Отчество","Если есть"),("section","Литера","А"),
            ("shift","Смена","Например, 1"),("address","Адрес","Город, улица, дом, квартира"),
            ("guardian_name","Родитель / опекун","ФИО"),
            ("guardian_phone","Телефон родителя","+992 ..."),
            ("extra_phone","Дополнительный телефон","+992 ..."),
            ("admitted_on","Дата поступления","ГГГГ-ММ-ДД; можно оставить пустым"),
            ("left_on","Дата выбытия","ГГГГ-ММ-ДД; можно оставить пустым"),
            ("from_school","Прибыл из школы","Название предыдущей школы"),
            ("to_school","Выбыл в школу","Название школы назначения"),
        ]:
            w=QLineEdit(str(self.initial.get(key, "") or ""))
            w.setPlaceholderText(ph)
            self.entries[key]=w
        for key in ("surname","given_name","patronymic"):
            form.addRow({"surname":"Фамилия *", "given_name":"Имя *", "patronymic":"Отчество"}[key],self.entries[key])
        self.dob=QLineEdit(str(self.initial.get("date_of_birth", "") or ""))
        self.dob.setPlaceholderText("ГГГГ-ММ-ДД; можно заполнить позднее")
        form.addRow("Дата рождения",self.dob)
        self.grade=QSpinBox(); self.grade.setRange(1,12); self.grade.setValue(int(self.initial.get("grade",1)))
        form.addRow("Класс *",self.grade)
        for key in ("section","shift","address","guardian_name","guardian_phone","extra_phone","admitted_on","from_school","to_school","left_on"):
            form.addRow({"section":"Литера", "shift":"Смена", "address":"Адрес", "guardian_name":"Родитель / опекун",
                         "guardian_phone":"Телефон родителя", "extra_phone":"Доп. телефон", "admitted_on":"Дата поступления", "from_school":"Прибыл из школы", "to_school":"Выбыл в школу", "left_on":"Дата выбытия"}[key],self.entries[key])
        self.status=QComboBox(); self.status.addItems(STATUSES)
        self.status.setCurrentText(self.initial.get("status","Учится"))
        form.addRow("Статус",self.status)
        self.notes=QTextEdit(self.initial.get("notes", ""))
        self.notes.setPlaceholderText("Дополнительная служебная информация (без ненужных персональных данных)")
        self.notes.setFixedHeight(84)
        form.addRow("Примечание",self.notes)
        buttons=QHBoxLayout()
        if self.initial.get('id'):
            buttons.addWidget(btn("История переводов и классов",self.show_history))
        buttons.addStretch()
        buttons.addWidget(btn("Отмена",self.reject))
        buttons.addWidget(btn("Сохранить",self.validate_and_accept,"Primary"))
        main.addLayout(buttons)

    def validate_and_accept(self):
        raw={key:widget.text() for key,widget in self.entries.items()}
        raw.update({"id":self.initial.get("id",""),"date_of_birth":self.dob.text().strip(),
                    "grade":self.grade.value(),"status":self.status.currentText(),"notes":self.notes.toPlainText()})
        try:
            self.result_record=clean_record(raw,allow_incomplete=True)
        except StudentValidationError as exc:
            message(self,"Проверьте карточку",str(exc),"warning")
            return
        self.accept()


    def show_history(self):
        database=getattr(self.parent(), 'db', None)
        if database is None:return
        history=database.transfer_history(self.initial['id'])
        grades=database.grade_history(self.initial['id'])
        lines=['ПЕРЕВОДЫ МЕЖДУ ШКОЛАМИ']
        for e in history:
            direction='Прибыл из' if e['direction']=='in' else 'Выбыл в'
            when=e['event_date'] or 'дата события не указана'
            lines.append(f"{direction}: {e['school_name']} | {when} | запись: {e['recorded_at'][:16]+' UTC'}")
        if not history:lines.append('Пока нет событий')
        lines.append('\nИСТОРИЯ КЛАССОВ')
        for e in grades:
            lines.append(f"{e['old_grade']} → {e['new_grade']} | учебный год {e['school_year']} | {e['source']}")
        if not grades:lines.append('Пока нет событий')
        message(self,'История ученика','\n'.join(lines))


class MainWindow(QMainWindow):
    PAGE_SIZE=100
    def __init__(self, db: Database, password: str):
        super().__init__()
        self.db=db
        self.password=password
        self.page=0
        self.deleted=False
        self.promotions=self.db.promote_for_school_year()
        self.dirty=bool(self.promotions)
        self.setWindowTitle(f"Школьный реестр · {APP_NAME} {__version__}")
        self.resize(1360,850)
        self.setMinimumSize(1040,650)
        base=QWidget(); self.setCentralWidget(base)
        h=QHBoxLayout(base); h.setContentsMargins(0,0,0,0); h.setSpacing(0)
        side=QFrame(); side.setObjectName("Sidebar"); side.setFixedWidth(246)
        sl=QVBoxLayout(side); sl.setContentsMargins(19,28,19,20); sl.setSpacing(12)
        sl.addWidget(label(APP_NAME.upper(),"Brand"))
        sl.addWidget(label("Локальная система учёта"))
        sl.addSpacing(30)
        self.nav_buttons=[]
        for title,idx in [("Дашборд",0),("Ученики",1),("Переводы",2),("Корзина",3),("Резервные копии",4),("О программе",5)]:
            b=btn(title,lambda checked=False,i=idx:self.navigate(i),"Nav")
            sl.addWidget(b); self.nav_buttons.append(b)
        sl.addStretch()
        sl.addWidget(btn("Изменить пароль",self.change_password,"Nav"))
        privacy=label("БЕЗ ИНТЕРНЕТА  •  ЛОКАЛЬНО")
        privacy.setStyleSheet("color:#8FB2F4;font-size:10px;font-weight:700")
        sl.addWidget(privacy)
        sl.addWidget(label("Данные остаются на устройстве."))
        h.addWidget(side)
        self.stack=QStackedWidget(); h.addWidget(self.stack,1)
        self.dashboard_page=self.make_dashboard_page()
        self.students_page=self.make_list_page(deleted=False)
        self.transfers_page=self.make_transfers_page()
        self.trash_page=self.make_list_page(deleted=True)
        self.backup_page=self.make_backup_page()
        self.about_page=self.make_about_page()
        for page in (self.dashboard_page,self.students_page,self.transfers_page,self.trash_page,self.backup_page,self.about_page):
            self.stack.addWidget(page)
        self.navigate(0)
        QShortcut(QKeySequence("Ctrl+N"),self,activated=self.add_student)
        QShortcut(QKeySequence("Ctrl+F"),self,activated=lambda:(self.navigate(1),self.search.setFocus()))
        QShortcut(QKeySequence("Ctrl+E"),self,activated=self.export)
        self.refresh()
        self.backup_timer=QTimer(self)
        self.backup_timer.timeout.connect(self.hourly_tasks)
        self.backup_timer.start(60*60*1000)
        QTimer.singleShot(200,self.auto_backup)

    def change_password(self):
        dialog=PasswordChangeDialog(self)
        if dialog.exec()!=QDialog.DialogCode.Accepted:
            return
        try:
            self.db.change_password(dialog.current_password,dialog.new_password)
            self.password=dialog.new_password
            message(self,"Пароль изменён","Новый пароль будет использоваться для входа и резервных копий.")
        except (ValueError,PermissionError) as exc:
            message(self,"Изменение пароля",str(exc),"warning")

    def make_dashboard_page(self):
        page=QWidget();root=QVBoxLayout(page)
        root.setContentsMargins(33,28,33,26);root.setSpacing(19)
        root.addWidget(label('Обзор школы','PageTitle'))
        root.addWidget(label('Актуальные сведения о контингенте на текущий учебный год.','Muted'))
        self.dashboard_metrics={}
        metrics=[('Учатся сейчас','studying'),('Переведены','transferred'),('Выпускники','graduated'),
                 ('Всего карточек','active'),('Прибыли из других школ','arriving'),('Требуют проверки','review')]
        grid=QGridLayout();grid.setSpacing(13)
        for index,(title,key) in enumerate(metrics):
            card=QFrame();card.setObjectName('Card');v=QVBoxLayout(card);v.setContentsMargins(19,17,19,17)
            v.addWidget(label(title,'Muted'))
            n=label('0','Metric');v.addWidget(n);grid.addWidget(card,index//3,index%3)
            self.dashboard_metrics[key]=n
        root.addLayout(grid)
        panel=QFrame();panel.setObjectName('TableCard');layout=QVBoxLayout(panel)
        title=label('Распределение учеников по классам');title.setStyleSheet('font-size:19px;font-weight:750')
        layout.addWidget(title)
        self.year_label=label('','Muted');layout.addWidget(self.year_label)
        self.grade_table=QTableWidget(0,3)
        self.grade_table.setHorizontalHeaderLabels(['Класс','Литера','Учатся'])
        self.grade_table.verticalHeader().hide();self.grade_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.grade_table.setAlternatingRowColors(True)
        self.grade_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.grade_table)
        root.addWidget(panel,1)
        controls=QHBoxLayout()
        controls.addWidget(label('Выпускной класс:','Muted'))
        self.final_grade_box=QComboBox()
        for v in (11,12):self.final_grade_box.addItem(str(v),v)
        self.final_grade_box.setCurrentIndex(self.final_grade_box.findData(self.db.final_grade()))
        self.final_grade_box.currentIndexChanged.connect(self.update_final_grade)
        controls.addWidget(self.final_grade_box)
        controls.addStretch()
        controls.addWidget(btn('Обновить статистику',self.refresh_dashboard))
        controls.addWidget(btn('Открыть учеников',lambda:self.navigate(1),'Primary'))
        root.addLayout(controls)
        return page

    def make_transfers_page(self):
        page=QWidget();root=QVBoxLayout(page)
        root.setContentsMargins(33,27,33,26);root.setSpacing(16)
        root.addWidget(label('Журнал переводов','PageTitle'))
        root.addWidget(label('Фиксируются школа, направление, дата события (если известна) и время внесения записи.','Muted'))
        table=QTableWidget(0,5);self.transfers_table=table
        table.setHorizontalHeaderLabels(['Ученик','Событие','Школа','Дата события','Записано'])
        table.verticalHeader().hide()
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        root.addWidget(table,1)
        root.addWidget(btn('Обновить журнал',self.refresh_transfers),alignment=Qt.AlignmentFlag.AlignRight)
        return page

    def refresh_transfers(self):
        rows=self.db.con.execute("SELECT e.*,s.surname,s.given_name FROM transfer_events e "
            "JOIN students s ON e.student_id=s.id ORDER BY e.id DESC LIMIT 1000").fetchall()
        t=self.transfers_table;t.setRowCount(len(rows))
        for i,r in enumerate(rows):
            vals=[f"{r['surname']} {r['given_name']}", 'Прибыл из' if r['direction']=='in' else 'Выбыл в',
                  r['school_name'],r['event_date'] or 'Не указана',r['recorded_at'][:16]+' UTC']
            for j,v in enumerate(vals):t.setItem(i,j,QTableWidgetItem(str(v)))

    def refresh_dashboard(self):
        stats=self.db.stats()
        for k,w in self.dashboard_metrics.items():w.setText(str(stats[k]))
        from .storage import school_year
        current=school_year()
        self.year_label.setText(f'Учебный год: {current}/{current+1} · переход 1 сентября · выпускной класс: {self.db.final_grade()}.')
        rows=self.db.grade_distribution();self.grade_table.setRowCount(len(rows))
        for i,r in enumerate(rows):
            for j,v in enumerate((r['grade'],r['section'] or '—',r['total'])):
                self.grade_table.setItem(i,j,QTableWidgetItem(str(v)))

    def update_final_grade(self):
        value=self.final_grade_box.currentData()
        if value is None:return
        if value==self.db.final_grade():return
        if QMessageBox.question(self,'Изменить выпускной класс?',
            f'Выпускным классом станет {value}. Уже сохранённые классы не уменьшаются. Продолжить?',
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:
            self.final_grade_box.blockSignals(True)
            self.final_grade_box.setCurrentIndex(self.final_grade_box.findData(self.db.final_grade()))
            self.final_grade_box.blockSignals(False)
            return
        self.db.set_final_grade(value)
        self.dirty=True
        self.refresh_dashboard()

    def hourly_tasks(self):
        promoted=self.db.promote_for_school_year()
        if promoted:self.dirty=True
        self.refresh()
        self.auto_backup()

    def make_list_page(self,deleted: bool):
        page=QWidget(); root=QVBoxLayout(page)
        root.setContentsMargins(32,27,32,24); root.setSpacing(19)
        top=QHBoxLayout()
        titles=QVBoxLayout(); titles.setSpacing(3)
        titles.addWidget(label("Корзина" if deleted else "Список учеников","PageTitle"))
        titles.addWidget(label("Удалённые карточки можно восстановить." if deleted else "Управляйте учениками и документами Excel / Word.","Muted"))
        top.addLayout(titles); top.addStretch()
        root.addLayout(top)
        if not deleted:
            actions=QHBoxLayout()
            actions.setSpacing(10)
            actions.addStretch()
            template_button = btn("Шаблоны ▾")
            menu = QMenu(template_button)
            choices = [
                ("Excel · полный (без ID)",False,False,'.xlsx'),
                ("Excel · тоҷикӣ (6 колонок)",True,False,'.xlsx'),
                ("Excel · тоҷикӣ (без №)",True,True,'.xlsx'),
                ("Word · полный (без ID)",False,False,'.docx'),
                ("Word · тоҷикӣ (6 колонок)",True,False,'.docx'),
                ("Word · тоҷикӣ (без №)",True,True,'.docx'),
                ("Word DOC · тоҷикӣ (без №)",True,True,'.doc'),
            ]
            for title, tajik, no_number, suffix in choices:
                action = menu.addAction(title)
                action.triggered.connect(lambda checked=False,t=tajik,n=no_number,e=suffix:
                                         self.save_any_template(t,n,e))
            template_button.setMenu(menu)
            actions.addWidget(template_button)
            actions.addWidget(btn("Импорт файла",self.import_students))
            actions.addWidget(btn("Экспорт файла",self.export))
            actions.addWidget(btn("+ Добавить ученика",self.add_student,"Primary"))
            root.addLayout(actions)
        if not deleted:
            tiles=QHBoxLayout(); tiles.setSpacing(13)
            for title,key in [("Всего учеников","active"),("Классов / групп","classes"),("В корзине","trashed")]:
                card=QFrame(); card.setObjectName("Card")
                layout=QVBoxLayout(card);layout.setContentsMargins(19,15,19,15)
                layout.addWidget(label(title,"Muted"))
                num=label("0","Metric"); layout.addWidget(num)
                tiles.addWidget(card)
                if not hasattr(self,"metric_labels"): self.metric_labels={}
                self.metric_labels[key]=num
            root.addLayout(tiles)
        searchbar=QFrame(); searchbar.setObjectName("ToolbarCard")
        sr=QHBoxLayout(searchbar); sr.setContentsMargins(15,14,15,14)
        finder=QLineEdit(); finder.setPlaceholderText("Поиск по ФИО, адресу или родителю...")
        finder.setClearButtonEnabled(True)
        finder.setMinimumWidth(320)
        grade=QComboBox(); grade.addItem("Все классы",None)
        for i in range(1,13):grade.addItem(f"{i} класс",i)
        status_filter=QComboBox()
        status_filter.addItem('Все статусы',None)
        for st in STATUSES:status_filter.addItem(st,st)
        sr.addWidget(finder,1); sr.addWidget(grade);sr.addWidget(status_filter)
        root.addWidget(searchbar)
        card=QFrame();card.setObjectName("TableCard")
        cl=QVBoxLayout(card); cl.setContentsMargins(1,1,1,1)
        table=QTableWidget(0,8)
        table.setHorizontalHeaderLabels(["Фамилия", "Имя", "Класс", "Статус", "Дата рождения", "Родитель / опекун", "Телефон", "Адрес"])
        table.verticalHeader().hide()
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setWordWrap(False)
        table.verticalHeader().setDefaultSectionSize(48)
        hdr=table.horizontalHeader();hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2,QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3,QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4,QHeaderView.ResizeMode.ResizeToContents)
        table.doubleClicked.connect(self.selected_action)
        cl.addWidget(table)
        bottom=QHBoxLayout();bottom.setContentsMargins(14,11,14,11)
        count=label("0 записей","Muted");bottom.addWidget(count)
        bottom.addStretch()
        primary=btn("Восстановить" if deleted else "Редактировать",self.selected_action)
        secondary=btn("Удалить в корзину",self.soft_delete,"Danger") if not deleted else None
        bottom.addWidget(primary)
        if secondary:bottom.addWidget(secondary)
        prev=btn("←",lambda:self.change_page(-1));next_=btn("→",lambda:self.change_page(1))
        page_label=label("1 / 1","Muted")
        bottom.addSpacing(10);bottom.addWidget(prev);bottom.addWidget(page_label);bottom.addWidget(next_)
        cl.addLayout(bottom);root.addWidget(card,1)
        if deleted:
            self.trash_search=finder;self.trash_grade=grade;self.trash_status=status_filter;self.trash_table=table
            self.trash_count=count;self.trash_prev=prev;self.trash_next=next_;self.trash_page_label=page_label
        else:
            self.search=finder;self.grade_filter=grade;self.status_filter=status_filter;self.table=table
            self.count=count;self.prev=prev;self.next=next_;self.page_label=page_label
        finder.textChanged.connect(self.defer_search)
        grade.currentIndexChanged.connect(self.reset_and_refresh)
        status_filter.currentIndexChanged.connect(self.reset_and_refresh)
        return page

    def make_backup_page(self):
        page=QWidget();r=QVBoxLayout(page);r.setContentsMargins(36,29,36,24);r.setSpacing(20)
        r.addWidget(label("Резервные копии","PageTitle"))
        r.addWidget(label("Все резервные копии защищаются паролем входа и шифруются AES-256-GCM.","Muted"))
        c=QFrame();c.setObjectName("Card");v=QVBoxLayout(c);v.setContentsMargins(22,22,22,22);v.setSpacing(12)
        t=label("Сохранить копию");t.setStyleSheet("font-size:19px;font-weight:750")
        v.addWidget(t)
        desc=label("Сохраните файл .srbak на флешку или внешний диск. Рекомендуется делать это регулярно.","Muted")
        desc.setWordWrap(True);v.addWidget(desc)
        v.addWidget(btn("Создать зашифрованную копию",self.manual_backup,"Primary"),alignment=Qt.AlignmentFlag.AlignLeft)
        r.addWidget(c)
        restore=QFrame();restore.setObjectName("Card")
        rr=QVBoxLayout(restore);rr.setContentsMargins(22,22,22,22);rr.setSpacing(12)
        rt=label("Восстановить из копии");rt.setStyleSheet("font-size:19px;font-weight:750")
        rr.addWidget(rt)
        d=label("Текущая база будет заменена. Перед заменой программа автоматически создаст страховочную копию.","Muted")
        d.setWordWrap(True);rr.addWidget(d)
        rr.addWidget(btn("Восстановить из .srbak",self.do_restore),alignment=Qt.AlignmentFlag.AlignLeft)
        r.addWidget(restore)
        r.addWidget(label(f"Папка данных: {data_directory()}","Muted"))
        r.addStretch()
        return page

    def make_about_page(self):
        page=QWidget();r=QVBoxLayout(page);r.setContentsMargins(36,29,36,24);r.setSpacing(17)
        r.addWidget(label("О программе","PageTitle"))
        c=QFrame();c.setObjectName("Card");v=QVBoxLayout(c);v.setContentsMargins(24,24,24,24);v.setSpacing(13)
        for line in [f"{APP_NAME} — локальная система школьного учёта",f"Версия: {__version__}",
                     "Автор: Ашуров Абдулло",
                     "Данные: SQLite (отдельно от папки программы)",
                     "Импорт/экспорт: Excel .xlsx, Word .docx и Word .doc",
                     "Защита входа: PBKDF2-HMAC-SHA256 и задержка при ошибках",
                     "Резервные копии: AES-256-GCM + scrypt",
                     "Интернет: не используется; аккаунт в облаке не нужен.",
                     "Важно: сама база не зашифрована. Защитите учётную запись и диск компьютера.",
                     "Перед удалением папки данных сохраните отдельную резервную копию."]:
            x=label(line);x.setWordWrap(True);v.addWidget(x)
        r.addWidget(c);r.addStretch()
        return page

    def navigate(self,idx):
        self.stack.setCurrentIndex(idx)
        self.deleted=(idx==3)
        self.page=0
        for i,b in enumerate(self.nav_buttons):
            b.setProperty("active","true" if i==idx else "false")
            b.style().unpolish(b);b.style().polish(b)
        self.refresh()

    def defer_search(self,*_):
        if not hasattr(self,"debouncer"):
            self.debouncer=QTimer(self);self.debouncer.setSingleShot(True)
            self.debouncer.timeout.connect(self.reset_and_refresh)
        self.debouncer.start(280)

    def reset_and_refresh(self,*_):
        self.page=0;self.refresh()

    def refresh(self):
        self.refresh_dashboard()
        if self.stack.currentIndex()==2:self.refresh_transfers()
        if self.stack.currentIndex() not in (1,3):return
        is_trash=self.stack.currentIndex()==3
        table=self.trash_table if is_trash else self.table
        text=(self.trash_search if is_trash else self.search).text()
        grade=(self.trash_grade if is_trash else self.grade_filter).currentData()
        status=(self.trash_status if is_trash else self.status_filter).currentData()
        rows,total=self.db.list_students(search=text,grade=grade,status=status,deleted=is_trash,page=self.page,page_size=self.PAGE_SIZE)
        if self.page and not rows:
            self.page=max(0,self.page-1)
            rows,total=self.db.list_students(search=text,grade=grade,status=status,deleted=is_trash,page=self.page,page_size=self.PAGE_SIZE)
        table.setRowCount(len(rows))
        for row_no,item in enumerate(rows):
            values=[item["surname"],item["given_name"],f'{item["grade"]}{item["section"]}',item["status"],
                    item["date_of_birth"],item["guardian_name"],item["guardian_phone"],item["address"]]
            for col,val in enumerate(values):
                cell=QTableWidgetItem(str(val))
                cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                cell.setToolTip(str(val))
                if col==0:cell.setData(Qt.ItemDataRole.UserRole,item["id"])
                table.setItem(row_no,col,cell)
        if rows: table.selectRow(0)
        page_count=max(1,(total+self.PAGE_SIZE-1)//self.PAGE_SIZE)
        count=self.trash_count if is_trash else self.count
        prev=self.trash_prev if is_trash else self.prev
        next_=self.trash_next if is_trash else self.next
        page_label=self.trash_page_label if is_trash else self.page_label
        count.setText(f"Найдено записей: {total}")
        page_label.setText(f"{self.page+1} / {page_count}")
        prev.setEnabled(self.page>0);next_.setEnabled(self.page+1<page_count)
        for key,value in self.db.stats().items():
            if key in self.metric_labels:self.metric_labels[key].setText(str(value))

    def change_page(self,diff):
        self.page=max(0,self.page+diff);self.refresh()

    def selected_id(self):
        table=self.trash_table if self.deleted else self.table
        row=table.currentRow()
        cell=table.item(row,0) if row>=0 else None
        if cell is None:
            message(self,"Выберите запись","Выберите ученика в таблице.")
            return None
        return cell.data(Qt.ItemDataRole.UserRole)

    def selected_action(self,*_):
        record_id=self.selected_id()
        if not record_id:return
        if self.deleted:
            if self.db.restore(record_id):
                self.dirty=True;self.refresh()
        else:self.edit_student(record_id)

    def add_student(self):
        if self.stack.currentIndex()!=1:self.navigate(1)
        dlg=StudentDialog(parent=self)
        if dlg.exec()==QDialog.DialogCode.Accepted:
            try:
                self.db.save_student(dlg.result_record)
                self.dirty=True;self.refresh()
            except Exception as exc:
                message(self,"Ошибка сохранения",str(exc),"error")

    def edit_student(self,record_id):
        existing=self.db.get_student(record_id)
        if not existing:return
        dlg=StudentDialog(existing,self)
        if dlg.exec()==QDialog.DialogCode.Accepted:
            try:
                self.db.save_student(dlg.result_record,record_id)
                self.dirty=True;self.refresh()
            except Exception as exc:
                message(self,"Ошибка сохранения",str(exc),"error")

    def soft_delete(self):
        record_id=self.selected_id()
        if not record_id:return
        if QMessageBox.question(self,"Переместить в корзину?","Карточка исчезнет из списка, но её можно восстановить.",
                                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            if self.db.soft_delete(record_id):
                self.dirty=True;self.refresh()

    def save_any_template(self, tajik: bool = False, no_number: bool = False, suffix: str = '.xlsx'):
        base = ('Шаблон_таджикский_без_номера' if tajik and no_number
                else 'Шаблон_переводов_таджикский' if tajik else 'Шаблон_учеников')
        filter_ = {'xlsx':'Excel (*.xlsx)','docx':'Word (*.docx)','doc':'Word 97-2003 (*.doc)'}[suffix[1:]]
        path, _ = QFileDialog.getSaveFileName(self, 'Сохранить шаблон',
                                              str(Path.home()/(base+suffix)),filter_)
        if not path:return
        try:
            target=Path(path)
            if not target.suffix: target=target.with_suffix(suffix)
            if target.suffix.lower()!=suffix: raise ValueError('Выберите правильное расширение файла')
            save_template(target,tajik=tajik,no_number=no_number)
            message(self,'Шаблон готов','Заполните только данные учеников. Номера и ID не нужны: программа создаст их автоматически.')
        except Exception as exc:message(self,'Ошибка шаблона',str(exc),'error')

    def template(self):
        self.save_any_template()

    def template_tajik(self):
        self.save_any_template(tajik=True)

    def export(self):
        options = ('Excel — полный (*.xlsx);;Excel — тоҷикӣ (*.xlsx);;'
                   'Word — полный (*.docx);;Word — тоҷикӣ (*.docx);;'
                   'Word DOC — полный (*.doc);;Word DOC — тоҷикӣ (*.doc)')
        path,selected=QFileDialog.getSaveFileName(self,'Экспорт учеников',
                         str(Path.home()/'Ученики_экспорт'),options)
        if not path:return
        try:
            target=Path(path)
            tajik='тоҷикӣ' in selected
            ext='.docx' if selected.startswith('Word —') else '.doc' if selected.startswith('Word DOC') else '.xlsx'
            if not target.suffix:target=target.with_suffix(ext)
            if target.suffix.lower()!=ext:
                raise ValueError('Расширение файла не совпадает с выбранным форматом')
            if ext=='.xlsx': export_excel(self.db,target,tajik=tajik)
            else: export_word(self.db,target,tajik=tajik)
            message(self,'Экспорт завершён',
                    'Файл готов. В нём могут быть личные данные детей — храните его в защищённом месте.')
        except Exception as exc:message(self,'Ошибка экспорта',str(exc),'error')

    def import_students(self):
        path,_=QFileDialog.getOpenFileName(self,'Выберите Excel/Word файл','','Excel и Word (*.xlsx *.docx *.doc)')
        if not path:return
        try: preview=parse_file(Path(path))
        except Exception as exc:
            message(self,"Ошибка чтения",str(exc),"error");return
        if preview.errors:
            details="\n".join(preview.errors[:14])
            if len(preview.errors)>14:details+=f"\n... и ещё {len(preview.errors)-14} ошибок."
            message(self,"Импорт отменён: исправьте файл",details,"warning")
            return
        if not preview.records:
            message(self,"Пустой файл","Нет заполненных строк для импорта.")
            return
        answer=QMessageBox(self)
        answer.setWindowTitle("Проверка импорта")
        warning='\n'.join(preview.warnings)
        answer.setText(f"Проверено {len(preview.records)} строк без ошибок.\n" +
                       (warning+'\n' if warning else '') +
                       'Как поступить с уже существующими записями?')
        skip=answer.addButton("Пропустить существующие",QMessageBox.ButtonRole.AcceptRole)
        update=answer.addButton("Обновить существующие",QMessageBox.ButtonRole.ActionRole)
        answer.addButton("Отмена",QMessageBox.ButtonRole.RejectRole)
        answer.exec()
        if answer.clickedButton() not in (skip,update):return
        mode="update" if answer.clickedButton()==update else "skip"
        try:
            added,changed=self.db.apply_import(preview.records,mode)
            self.dirty=added>0 or changed>0 or self.dirty
            self.page=0;self.refresh()
            message(self,"Импорт завершён",f"Добавлено: {added}\nОбновлено: {changed}\nПропущено: {len(preview.records)-added-changed}")
        except Exception as exc:message(self,"Ошибка импорта",f"Изменения отменены.\n{exc}","error")

    def auto_backup(self,force=False):
        try:
            folder=self.db.path.parent/"backups"
            from .storage import _secure_dir
            _secure_dir(folder)
            when=self.db.get_setting("last_autobackup","")
            last=datetime.fromisoformat(when) if when else None
            if force or last is None or datetime.now(timezone.utc)-last>timedelta(hours=24):
                create_backup(self.db,folder/suggested_backup_name(),self.password)
                self.db.set_setting("last_autobackup",datetime.now(timezone.utc).isoformat())
                files=sorted(
                    path for pattern in ("maktabos_*.srbak", "schoolregister_*.srbak")
                    for path in folder.glob(pattern)
                )
                for old in files[:-20]:old.unlink(missing_ok=True)
                self.dirty=False
        except Exception as exc:
            # Important not to claim successful automatic backup in this case.
            message(self,"Резервное копирование не удалось",f"Создайте копию вручную.\n{exc}","warning")

    def manual_backup(self):
        path,_=QFileDialog.getSaveFileName(self,"Сохранить зашифрованную копию",str(Path.home()/suggested_backup_name()),f"{APP_NAME} (*.srbak)")
        if not path:return
        try:
            create_backup(self.db,Path(path),self.password)
            message(self,"Копия создана","Файл защищён паролем входа. Сохраните его отдельно от компьютера.")
        except Exception as exc:message(self,"Ошибка копирования",str(exc),"error")

    def do_restore(self):
        path,_=QFileDialog.getOpenFileName(self,"Открыть резервную копию","",f"{APP_NAME} (*.srbak)")
        if not path:return
        if QMessageBox.question(self,"Подтверждение восстановления",
            "Текущие записи будут заменены данными из копии. Перед заменой создаётся страховочная копия. Продолжить?",
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        from PySide6.QtWidgets import QInputDialog
        pw,ok=QInputDialog.getText(self,"Пароль копии","Пароль, которым была создана копия:",QLineEdit.EchoMode.Password)
        if not ok:return
        try:
            folder=self.db.path.parent/"backups"
            from .storage import _secure_dir
            _secure_dir(folder)
            create_backup(self.db,folder/("before_restore_"+suggested_backup_name()),self.password)
            restore_backup(self.db,Path(path),pw)
            message(self,"Восстановление завершено","База восстановлена. Запустите приложение заново и войдите с паролем резервной копии.")
            self.dirty=False
            self.close()
        except Exception as exc:message(self,"Восстановление не выполнено",str(exc),"error")

    def closeEvent(self,event):
        if self.dirty:
            self.auto_backup(force=True)
        self.password=""
        event.accept()


def main():
    qt=QApplication(sys.argv)
    qt.setApplicationName(APP_NAME)
    qt.setOrganizationName(APP_NAME)
    asset_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    qt.setWindowIcon(QIcon(str(asset_root / "assets" / "maktabos.ico")))
    qt.setStyle("Fusion")
    qt.setStyleSheet(THEME)
    qt.setFont(QFont("Segoe UI",10))
    try:
        db=Database()
    except Exception as exc:
        QMessageBox.critical(None,"Ошибка базы",f"База недоступна:\n{exc}")
        return 1
    login=LoginDialog(db)
    if login.exec()!=QDialog.DialogCode.Accepted:
        db.close();return 0
    window=MainWindow(db,login.password)
    window.show()
    result=qt.exec()
    db.close()
    return result
