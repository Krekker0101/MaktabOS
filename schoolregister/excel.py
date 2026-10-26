"""Offline Excel and Word interchange, with optional numbering and auto-generated UUIDs."""
from __future__ import annotations
import hashlib
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .constants import FIELDS, HEADERS, MAX_IMPORT_BYTES, MAX_IMPORT_ROWS, TAJIK_HEADERS
from .validation import PHONE_PATTERN, StudentValidationError, clean_record

WIDTHS = [39,21,19,20,20,12,12,12,38,28,26,24,20,18,36,30,30,21]


def asset_template() -> Path:
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parents[1]))
    return root / 'assets' / 'Шаблон_учеников.xlsx'


def asset_tajik_template() -> Path:
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parents[1]))
    return root / 'assets' / 'Шаблон_переводов_таджикский.xlsx'


def asset_tajik_no_number_template() -> Path:
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parents[1]))
    return root / 'assets' / 'Шаблон_таджикский_без_номера.xlsx'


def save_template(target: Path, tajik: bool = False, *, no_number: bool = False) -> None:
    target = Path(target)
    if target.suffix.lower() in ('.docx', '.doc'):
        from .word import save_word_template
        save_word_template(target, tajik=tajik, no_number=no_number)
    elif target.suffix.lower() == '.xlsx':
        source = asset_tajik_no_number_template() if tajik and no_number else asset_tajik_template() if tajik else asset_template()
        shutil.copy2(source, target)
    else:
        raise ValueError('Формат шаблона: .xlsx, .docx или .doc')


def _prepare_sheet(ws, legacy: bool = False):
    headers = TAJIK_HEADERS if legacy else HEADERS
    widths = [8,39,13,24,33,33] if legacy else WIDTHS
    ws.freeze_panes = 'B2' if legacy else 'C2'
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 44
    for col, (name, width) in enumerate(zip(headers, widths),1):
        c = ws.cell(1,col,name)
        c.font=Font(name='Aptos',size=11,bold=True,color='FFFFFF')
        c.fill=PatternFill('solid',fgColor='173D77')
        c.alignment=Alignment(vertical='center',horizontal='center' if legacy else 'left',wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width=width
    ws.auto_filter.ref=f'A1:{get_column_letter(len(headers))}1'
    if legacy:
        for r in range(2,202):
            ws.cell(r,1,r-1)
            for col in range(1,7):
                c=ws.cell(r,col)
                c.fill=PatternFill('solid',fgColor='F3F7FC' if r%2==0 else 'FFFFFF')
                c.border=Border(bottom=Side(style='hair',color='DDE4EE'))
            ws.cell(r,4).number_format='@'
    else:
        for r in range(2,202):
            for c in range(1,len(headers)+1):
                ws.cell(r,c).fill=PatternFill('solid',fgColor='F3F7FC' if r%2==0 else 'FFFFFF')
        for col in ('K','L'):
            for r in range(2,1502):ws[f'{col}{r}'].number_format='@'
    grade=DataValidation(type='whole',operator='between',formula1='1',formula2='12',allow_blank=False)
    grade.showErrorMessage=True;grade.error='Введите класс от 1 до 12'
    ws.add_data_validation(grade)
    grade.add('C2:C30000' if legacy else 'F2:F30000')
    if not legacy:
        status=DataValidation(type='list',formula1='"Учится,Переведён,Выпущен"')
        ws.add_data_validation(status);status.add('N2:N30000')


def _instructions(wb):
    ws=wb.create_sheet('Инструкция');ws.sheet_view.showGridLines=False
    lines=[
        'SCHOOLREGISTER · ИНСТРУКЦИЯ',
        'Рабочий лист «Ученики». Старый 15-колоночный формат также поддерживается.',
        'Обязательные поля: фамилия, имя и класс. Дату рождения и адрес можно добавить позже.',
        '«Прибыл из школы» и «Выбыл в школу» записываются в историю перемещений.',
        '«Выбыл в школу» автоматически задаёт статус «Переведён».',
        'Дата поступления/выбытия заполняется при наличии документально подтверждённой даты.',
        'Учебный год начинается 1 сентября. Программа переводит только статус «Учится».',
        'После изменения имени, класса или статуса проверьте карточку ученика.',
        'Не публикуйте файлы Excel с персональными данными школьников.',
    ]
    ws.column_dimensions['A'].width=118
    for i,line in enumerate(lines,1):
        c=ws.cell(i,1,line);c.font=Font(name='Aptos',size=15 if i==1 else 11,bold=(i==1),color='173D77')
        c.alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[i].height=34 if i==1 else 28


def _text_cell(ws, r: int, c: int, value):
    cell=ws.cell(r,c,value)
    # Force potentially untrusted strings to literal text, never Excel formulas.
    if isinstance(value,str):cell.data_type='s'
    cell.font=Font(name='Aptos',size=10,color='27364A')
    cell.alignment=Alignment(vertical='center')
    if r%2==0:cell.fill=PatternFill('solid',fgColor='F3F7FC')
    return cell


def export_excel(db, target: Path, *, tajik: bool = False):
    wb=Workbook();ws=wb.active;ws.title='Рӯйхат' if tajik else 'Ученики'
    _prepare_sheet(ws,tajik)
    sql='SELECT * FROM students WHERE deleted_at IS NULL ORDER BY surname COLLATE NOCASE,given_name COLLATE NOCASE'
    for index,row in enumerate(db.con.execute(sql),2):
        if tajik:
            name=' '.join(part for part in [row['surname'],row['given_name'],row['patronymic']] if part)
            vals=[index-1,name,str(row['grade'])+(row['section'] or ''),row['guardian_phone'],row['to_school'],row['from_school']]
        else:
            vals=[row[key] for key in FIELDS]
        for col,value in enumerate(vals,1): _text_cell(ws,index,col,value)
    ws.auto_filter.ref=f'A1:{get_column_letter(6 if tajik else len(HEADERS))}{max(2,ws.max_row)}'
    if not tajik:_instructions(wb)
    Path(target).parent.mkdir(parents=True,exist_ok=True)
    wb.save(target)


@dataclass
class ImportPreview:
    records: list[dict]
    errors: list[str]
    blank_rows: int=0
    warnings: list[str]=field(default_factory=list)
    template: str=''


def _normalized(value) -> str:
    s=unicodedata.normalize('NFKC',str(value or '')).casefold()
    return re.sub(r'[^\w]+',' ',s).strip().replace('ё','е')


@dataclass(frozen=True)
class HeaderSpec:
    kind: str
    columns: tuple[int, ...]  # 0-based indices of business data fields; no sequence number
    fields: tuple[str, ...]
    has_number: bool = False


def _header_spec(row) -> HeaderSpec | None:
    names = [_normalized(v) for v in row]
    # Only the exact, documented column orders are accepted: no guessing field meanings.
    forms = [
        ('modern', HEADERS, FIELDS),
        ('original', HEADERS[:15], FIELDS[:15]),
        ('modern', HEADERS[1:], FIELDS[1:]),
        ('original', HEADERS[1:15], FIELDS[1:15]),
    ]
    for kind, headers, fields in forms:
        if names[:len(headers)] == [_normalized(h) for h in headers]:
            return HeaderSpec(kind, tuple(range(len(headers))), tuple(fields))
    # Tajik/Russian school tables can add address columns or change column order.
    tajik_fields = []
    for index, name in enumerate(names):
        field_name = _tajik_field(name)
        if field_name and field_name not in (field for _, field in tajik_fields):
            tajik_fields.append((index, field_name))
    field_names = {field for _, field in tajik_fields}
    has_name = 'full_name' in field_names or {'surname', 'given_name'} <= field_names
    has_supporting_field = bool(field_names & {
        'address', 'phone', 'guardian_phone', 'from_school', 'to_school',
        'guardian_name', 'date_of_birth',
    })
    if has_name and 'grade' in field_names and has_supporting_field:
        return HeaderSpec(
            'tajik',
            tuple(index for index, _ in tajik_fields),
            tuple(field for _, field in tajik_fields),
        )
    return None


def _tajik_field(name: str) -> str | None:
    """Recognize common Russian and Tajik labels without relying on column order."""
    if not name:
        return None
    if (name in ('ном ва насаб', 'ному насаб', 'ном насаб', 'фио', 'ф и о',
                 'номи пурра', 'ному насаби хонанда', 'full name')
            or ('ном' in name and 'насаб' in name)
            or ('имя' in name and 'фамилия' in name)):
        return 'full_name'
    if name in ('насаб', 'фамилия', 'surname'):
        return 'surname'
    if name in ('ном', 'исм', 'имя', 'given name', 'first name'):
        return 'given_name'
    if name in ('номи падар', 'отчество', 'patronymic'):
        return 'patronymic'
    if name.startswith(('синф', 'класс', 'grade')):
        return 'grade'
    if any(term in name for term in ('суроға', 'адрес', 'истиқомат', 'истикомат',
                                      'ҷои зист', 'ҷойи зист', 'ҷои истиқомат',
                                      'ҷойи истиқомат', 'чои истикомат', 'чойи истикомат',
                                      'макони зист', 'место проживания', 'место жительства',
                                      'адрес жительства')):
        return 'address'
    if 'телефон' in name or 'phone' in name:
        return 'phone'
    if name.startswith(('рафт', 'выбыл', 'выбыла')) or 'выбыл в школу' in name:
        return 'to_school'
    if name.startswith(('омад', 'прибыл', 'прибыла')) or 'прибыл из школы' in name:
        return 'from_school'
    if name in ('санаи таваллуд', 'дата рождения', 'date of birth'):
        return 'date_of_birth'
    if name in ('номи падару модар', 'номи падару модар ё парастор',
                'фио родителя', 'родитель', 'опекун', 'guardian'):
        return 'guardian_name'
    return None


def _header_kind(row) -> str | None:
    spec = _header_spec(row)
    return spec.kind if spec else None


def _split_name(value: str):
    parts=str(value or '').strip().split()
    if len(parts)<2:raise StudentValidationError('В графе «Ном ва насаб» нужны фамилия и имя (два слова минимум)')
    suffix=('ов','ова','ев','ева','зода','заде','пур','духт')
    if len(parts)>1 and parts[1].casefold().endswith(suffix) and not parts[0].casefold().endswith(suffix):
        return parts[1],parts[0],' '.join(parts[2:])
    return parts[0],parts[1],' '.join(parts[2:])


def _parse_class(value):
    raw=str(value or '').strip().replace('«','').replace('»','').replace('"','')
    match=re.fullmatch(r'(1[0-2]|[1-9])\s*[-– ]?\s*([A-Za-z\u0400-\u04FF]{0,3})',raw)
    if match:return int(match[1]),match[2].upper()
    raise StudentValidationError('Класс: укажите 1–12, можно с литерой (например 5А)')


def _legacy_record(values, fields) -> dict:
    """Convert a Tajik/Russian school table row using its recognized headings."""
    source = dict(zip(fields, values))
    if source.get('full_name'):
        surname, given, patronymic = _split_name(source['full_name'])
    else:
        surname = str(source.get('surname') or '').strip()
        given = str(source.get('given_name') or '').strip()
        patronymic = str(source.get('patronymic') or '').strip()
    grade,section=_parse_class(source.get('grade'))
    phone=source.get('phone', source.get('guardian_phone', ''))
    if isinstance(phone,float) and phone.is_integer():phone=int(phone)
    return clean_record({
        'surname':surname,'given_name':given,'patronymic':patronymic,
        'grade':grade,'section':section,'guardian_phone':str(phone or '').strip(),
        'guardian_name':source.get('guardian_name') or '',
        'to_school':str(source.get('to_school') or '').strip(),
        'from_school':str(source.get('from_school') or '').strip(),
        'date_of_birth':source.get('date_of_birth') or '',
        'address':source.get('address') or '',
        'status':'Переведён' if str(source.get('to_school') or '').strip() else 'Учится',
    }, allow_incomplete=True)


def _rows_import(rows, *, formula_cells=None) -> ImportPreview:
    records=[];errors=[];warnings=[];blanks=0;seen_ids=set()
    header_idx=None;spec=None
    for index, row in enumerate(rows):
        if index>29:break
        spec=_header_spec(row)
        if spec:
            header_idx=index;break
    if header_idx is None or spec is None:
        raise ValueError('Не найдены заголовки: ожидается полный шаблон или таджикские 5/6 колонок')
    if spec.kind == 'tajik':
        if 'address' not in spec.fields:
            warnings.append('В файле нет столбца места проживания; адрес можно заполнить позднее.')
        if 'date_of_birth' not in spec.fields:
            warnings.append('В файле нет даты рождения; её можно заполнить позднее.')
        if 'full_name' in spec.fields:
            warnings.append('Порядок имени/фамилии определяется приблизительно: проверьте карточки.')
    if 'id' not in spec.fields:
        warnings.append('ID отсутствует: приложение создаст уникальный UUID для каждого нового ученика.')

    for index, row in enumerate(rows[header_idx+1:],header_idx+2):
        if index>MAX_IMPORT_ROWS+31:raise ValueError('Больше 30 000 строк')
        cells=list(row)
        values=[cells[i] if i<len(cells) else None for i in spec.columns]
        if not any(v is not None and str(v).strip() for v in values):
            blanks+=1;continue
        if formula_cells and any((index,col+1) in formula_cells for col in spec.columns):
            errors.append(f'Строка {index}: формулы запрещены, нужны только значения');continue
        try:
            if spec.kind=='tajik':
                source_values=list(values)
                values=list(source_values)
                phone_field = next(
                    (name for name in ('phone', 'guardian_phone') if name in spec.fields),
                    None,
                )
                invalid_phone = False
                numeric_phone = False
                if phone_field:
                    phone_index = spec.fields.index(phone_field)
                    phone_value = values[phone_index]
                    numeric_phone = isinstance(phone_value, (int, float))
                    if isinstance(phone_value, float) and phone_value.is_integer():
                        phone_value = int(phone_value)
                    phone_text = str(phone_value or '').strip()
                    if phone_text and not PHONE_PATTERN.fullmatch(phone_text):
                        values[phone_index] = ''
                        invalid_phone = True
                rec=_legacy_record(values, spec.fields)
                if invalid_phone:
                    warnings.append(
                        f'Строка {index}: номер телефона пропущен из-за недопустимого символа; '
                        'остальные данные ученика импортируются.'
                    )
                # Same fingerprint with or without № and for .xlsx, .docx and .doc.
                source='\x1f'.join(_normalized(v) for v in source_values)
                rec['_fingerprint']=hashlib.sha256(('TJK-v3\x1f'+source).encode()).hexdigest()
                if numeric_phone and 'Телефон записан числом: ведущие нули могли пропасть в Excel.' not in warnings:
                    warnings.append('Телефон записан числом: ведущие нули могли пропасть в Excel.')
            else:
                rec=clean_record(dict(zip(spec.fields,values)),allow_incomplete=True)
                rec['_provided_fields'] = spec.fields
                if not rec['id']:
                    # Identical import reruns remain idempotent even after annual promotion.
                    source='\x1f'.join(_normalized(v) for v in values)
                    rec['_fingerprint']=hashlib.sha256(('FULL-v3\x1f'+source).encode()).hexdigest()
                if rec['id']:
                    if rec['id'] in seen_ids:raise StudentValidationError('Повторяющийся ID')
                    seen_ids.add(rec['id'])
            records.append(rec)
        except StudentValidationError as exc:
            errors.append(f'Строка {index}: {exc}')
        if len(errors)>=100:
            errors.append('Слишком много ошибок (первые 100)');break
    return ImportPreview(records,errors,blanks,warnings,spec.kind)


def parse_excel(path: Path) -> ImportPreview:
    """Compatibility API: reads Excel .xlsx, Word .docx, and legacy Word .doc."""
    path=Path(path)
    if path.suffix.lower() not in ('.xlsx','.docx','.doc'):
        raise ValueError('Поддерживаются Excel .xlsx и Word .docx/.doc')
    if not path.is_file():
        raise FileNotFoundError(f'Файл не найден: {path}')
    if path.stat().st_size>MAX_IMPORT_BYTES:
        raise ValueError('Размер файла превышает 15 МБ')
    if path.suffix.lower() in ('.docx','.doc'):
        from .word import parse_word
        return parse_word(path)
    wb=load_workbook(path,read_only=True,data_only=False,keep_links=False)
    try:
        result=None
        for ws in wb.worksheets:
            if ws.max_row and ws.max_row>MAX_IMPORT_ROWS+31:continue
            allrows=[];formula_rows=set()
            for rnum,cells in enumerate(ws.iter_rows(max_col=max(len(HEADERS),6)),1):
                if rnum>MAX_IMPORT_ROWS+31:break
                for cnum,c in enumerate(cells,1):
                    if c.data_type=='f':formula_rows.add((rnum,cnum))
                allrows.append([c.value for c in cells])
            try:
                result=_rows_import(allrows,formula_cells=formula_rows)
                break
            except ValueError:
                continue
        if result is None:raise ValueError('Не найден поддерживаемый лист с заголовками')
        return result
    finally:
        wb.close()


parse_file = parse_excel
