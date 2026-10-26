from __future__ import annotations
import re
import uuid
from datetime import date, datetime
from .constants import FIELDS, STATUSES

MAX_LEN = {"surname": 100, "given_name": 100, "patronymic": 100, "section": 12,
           "shift": 30, "address": 500, "guardian_name": 200,
           "guardian_phone": 50, "extra_phone": 50, "notes": 1000, "from_school": 250, "to_school": 250}
CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
PHONE_PATTERN = re.compile(r"[+0-9()\s.\-]{5,50}")

class StudentValidationError(ValueError):
    pass


def parse_date(value, *, optional=False):
    if value is None or (isinstance(value,str) and not value.strip()):
        if optional:
            return ""
        raise StudentValidationError("Укажите дату")
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        result = value
    elif isinstance(value, str):
        raw = value.strip()
        try:
            result = datetime.strptime(raw, "%Y-%m-%d").date() if re.fullmatch(r"\d{4}-\d{2}-\d{2}",raw) else datetime.strptime(raw,"%d.%m.%Y").date()
        except ValueError as exc:
            raise StudentValidationError("Дата: используйте ГГГГ-ММ-ДД или ДД.ММ.ГГГГ") from exc
    else:
        raise StudentValidationError("Дата должна быть текстом или датой Excel")
    if result.year < 1950 or result > date.today():
        raise StudentValidationError("Дата находится вне допустимого диапазона")
    return result.isoformat()


def clean_record(raw: dict, *, allow_incomplete: bool = False) -> dict:
    result: dict = {}
    if raw.get("id"):
        try:
            result["id"] = str(uuid.UUID(str(raw["id"]).strip()))
        except (ValueError,AttributeError) as exc:
            raise StudentValidationError("Неверный ID записи") from exc
    else:
        result["id"] = ""
    for field in FIELDS:
        if field in ("id","grade","date_of_birth","admitted_on","left_on","status"):
            continue
        value = raw.get(field, "")
        if value is None: value = ""
        if isinstance(value, (int,float)):
            value = str(value)
        if not isinstance(value,str):
            raise StudentValidationError(f"Поле «{field}»: требуется текст")
        value = value.strip()
        if CONTROL.search(value):
            raise StudentValidationError(f"Поле «{field}» содержит управляющие символы")
        if len(value) > MAX_LEN.get(field, 500):
            raise StudentValidationError(f"Слишком длинное поле «{field}»")
        result[field] = value
    for field in (("surname", "given_name") if allow_incomplete else ("surname", "given_name", "address")):
        if not result[field]:
            raise StudentValidationError(f"Обязательное поле «{field}» не заполнено")
    grade = raw.get("grade")
    if isinstance(grade, int):
        grade_num = grade
    elif isinstance(grade, float) and grade.is_integer():
        grade_num = int(grade)
    elif isinstance(grade, str) and re.fullmatch(r"(?:[1-9]|1[0-2])", grade.strip()):
        grade_num = int(grade)
    else:
        raise StudentValidationError("Класс должен быть целым числом от 1 до 12")
    if not 1 <= grade_num <= 12:
        raise StudentValidationError("Класс должен быть от 1 до 12")
    result["grade"] = grade_num
    result["date_of_birth"] = parse_date(raw.get("date_of_birth"), optional=allow_incomplete)
    result["admitted_on"] = parse_date(raw.get("admitted_on"), optional=True)
    result["left_on"] = parse_date(raw.get("left_on"), optional=True)
    status = str(raw.get("status") or "Учится").strip()
    if status not in STATUSES:
        raise StudentValidationError("Статус: Учится, Переведён или Выпущен")
    # Destination implies transfer for a currently studying record. Dates are never invented.
    if result.get("to_school") and status == "Учится":
        status = "Переведён"
    result["status"] = status
    for field in ("guardian_phone", "extra_phone"):
        phone = result[field]
        if phone and not PHONE_PATTERN.fullmatch(phone):
            raise StudentValidationError("Неверный формат телефонного номера")
    return result
