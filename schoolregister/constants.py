from __future__ import annotations

APP_NAME = "MaktabOS"
HEADERS = [
    "ID записи", "Фамилия", "Имя", "Отчество", "Дата рождения",
    "Класс", "Литера", "Смена", "Адрес проживания", "ФИО родителя/опекуна",
    "Телефон родителя/опекуна", "Дополнительный телефон", "Дата поступления",
    "Статус", "Примечание", "Прибыл из школы", "Выбыл в школу", "Дата выбытия",
]
FIELDS = [
    "id", "surname", "given_name", "patronymic", "date_of_birth", "grade",
    "section", "shift", "address", "guardian_name", "guardian_phone",
    "extra_phone", "admitted_on", "status", "notes", "from_school", "to_school", "left_on",
]
REQUIRED_FIELDS = ("surname", "given_name", "date_of_birth", "grade", "address")
STATUSES = ("Учится", "Переведён", "Выпущен")
MAX_IMPORT_BYTES = 15 * 1024 * 1024
MAX_IMPORT_ROWS = 30000

TAJIK_HEADERS = ["№", "Ном ва насаб", "Синф", "Рақами телефон", "Рафт (ба кадом мактаб)", "Омад (аз кадом мактаб)"]
