# MaktabOS 3.0 — школьный реестр, полностью локально

Обновлённое настольное приложение для учёта школьников на одном компьютере. Исходный код: Python 3.10+ / PySide6 (Qt) / SQLite. Для работы реестра интернет **не требуется** (для первой установки Python-библиотек может потребоваться интернет). Выбранная учётная запись ОС должна иметь права на запись в папку данных.

**Автор: Ашуров Абдулло.**

**English:** MaktabOS is an offline school student management application and student information system (SIS) for maintaining a local student register, importing Tajik school lists from Excel and Word, and storing student records in SQLite. English documentation is included below.

## Запуск

Чтобы собрать и установить Windows-версию, запустите `scripts/build_windows.bat` на Windows с установленными Python и Inno Setup 6. Готовый установщик `dist/MaktabOS-Setup-3.0.0.exe` по умолчанию устанавливает приложение в `Program Files\MaktabOS`; путь можно изменить в мастере установки. Установщик создаёт ярлык в меню «Пуск», а ярлык на рабочем столе можно выбрать дополнительно. Готовые бинарные файлы в исходный Git-репозиторий не включены.

Для запуска из исходников установите Python 3.11–3.14 и распакуйте проект в обычную папку. В Windows запустите `scripts/run_windows.bat`. На macOS/Linux запустите `bash scripts/run_linux_macos.sh`. При ручной установке: `python -m pip install -r requirements.txt`, `python -m schoolregister`.

Для сборки Windows-дистрибутива из исходников используйте `scripts/build_windows.bat` на Windows; требуются Inno Setup 6 и Python. Он создаст `dist/MaktabOS-Setup-3.0.0.exe`.

Для macOS/Linux — `bash scripts/run_linux_macos.sh`, при наличии Qt и Python. Для ручной установки: `python -m pip install -r requirements.txt`, `python -m schoolregister`.

## Главное — новые функции

- **Дашборд:** сколько реально учится, сколько выбыло, сколько выпустилось, сколько прибыло, суммарно карточек; распределение **только учащихся** по классам и литерам; список выпускных классов, требующих проверки.
- **Поддержка таджикских таблиц Excel и Word:** стандартная шапка из шести колонок с необязательным «№» поддерживается как раньше. Также распознаются русские и таджикские названия столбцов (включая `Чои истикомат` без диакритик), перестановка столбцов и адрес проживания. Импорт Word обрабатывает все подходящие таблицы документа, а не только первую.
- **Импорт Excel `.xlsx` и таблиц Word `.docx` / `.doc`** с такой же шапкой. Поддерживаются старые формы SchoolRegister на 15/18 полей с ID и новые формы на 14/17 полей без ID. Импорт `.jpg`, `.png`, PDF или старого бинарного `.xls` не предусмотрен: эти документы нужно предварительно перенести в `.xlsx` или `.docx`.
- В стандартном шаблоне переводов нет даты рождения и адреса; если в исходной таблице есть столбец места проживания, адрес будет сохранён в карточке ученика и SQLite. Если адреса или даты рождения нет, приложение оставит поле пустым и покажет предупреждение. Фамилия/имя из одного поля определяются по порядку слов и распространённым окончаниям; **обязательно проверьте** карточки после импорта. Если имя содержит только одно слово, импорт покажет ошибку.
- При заполненном `Рафт (ба кадом мактаб)` ученик получает статус **«Переведён»** и перестаёт учитываться в списке текущих учащихся. Названия школ прибытия и выбытия пишутся в отдельную историю с фактическим временем **внесения** записи. Дата настоящего перевода остаётся пустой, если школа её не указала.
- Карточка включает `Прибыл из школы`, `Выбыл в школу`, необязательные `Дата поступления` и `Дата выбытия`, а также кнопку **«История переводов и классов»**. При редактировании значения история не стирается. Чтобы вернуть переведённого ученика в статус «Учится», удалите из текущего поля «Выбыл в школу» старую школу: её название останется в журнале.
- **Переход в следующий класс с 1 сентября**: для учеников в статусе «Учится» при открытии после входа и затем ежечасно. Пропущено несколько лет — повышается на соответствующее количество классов. Каждая обработка идемпотентна: повторный запуск не повышает класс второй раз. Ученики в статусах «Переведён» и «Выпущен» не повышаются.
- **Выпускной класс** по умолчанию **11** (можно изменить на 12 на дашборде). Ученики выпускного класса автоматически не становятся выпускниками и не переводятся в следующий класс: их нужно проверить и вручную оформить выпуск.
- Экспорт **полной таблицы** (18 полей, включая постоянный ID для обновлений) либо **таджикской** (6 полей) в Excel или Word. Шаблоны для нового ввода не требуют ID и даже колонки «№».
- Поиск, фильтр класса, корзина с восстановлением, смена и восстановление пароля по ключу, журнал действий, автоматические зашифрованные `.srbak` резервные копии и ручное восстановление.


## Новое в 3.0: Word, необязательная нумерация, автоматические ID

- `.xlsx` — полноценный импорт/экспорт Excel и сохранение двух вариантов таджикского шаблона (с пустым `№` или **без** `№`).
- `.docx` — полноценный импорт/экспорт таблиц Word без Word/LibreOffice; в `assets/` есть 3 готовых шаблона.
- `.doc` — **реальный бинарный Word 97–2003**, не просто файл `.docx` с изменённым расширением. Для преобразования/экспорта требуется **Microsoft Word + pywin32** или установленный **LibreOffice**. Для надёжного импорта таблиц `.doc` используются **Antiword** (если установлен) либо Microsoft Word; в некоторых случаях конвертер LibreOffice может уплощать таблицы, после чего приложение покажет ошибку, а не импортирует повреждённые колонки. В таком случае сохраните файл в Word как `.docx` и импортируйте его. `.docx` рекомендуется как основной формат.
- Формат с фотографии: `№ | Ном ва насаб | Синф | Рақами телефон | Рафт (ба кадом мактаб) | Омад (аз кадом мактаб)`; допустима также версия **без `№`**. Порядковый номер никогда не является ID в базе.
- Полные формы: 18 столбцов с существующим UUID для изменения карточек, 17 столбцов без UUID для создания новых карточек; аналогично совместимая старая 15-колоночная форма с UUID и 14-колоночная без UUID.
- При импорте **каждому новому ученику генерируется UUID**, сохраняемый на диске; при последующем полном экспорте ID можно увидеть. Повторная загрузка неизменённой таблицы из поддерживаемых шаблонов не дублирует записи. Для правки уже сохранённого ученика используйте карточку или полный экспорт с ID.
- Если в источнике нет даты рождения, приложение **не делает вывод о личности по одному общему телефону родителя**: два разных ученика могут иметь один номер.

**Готовые файлы в `assets/`:**

- `Шаблон_учеников.xlsx` — 17 столбцов без ID.
- `Шаблон_переводов_таджикский.xlsx` — 6 столбцов с необязательной пустой графой №.
- `Шаблон_таджикский_без_номера.xlsx` — 5 столбцов, графы № вообще нет.
- Такие же три шаблона `.docx`; дополнительный образец `.doc` без №.

Для Windows со старым `.doc`: Word при наличии `pywin32` обеспечит преобразование. Другой вариант — пользоваться `.docx`, который работает автономно без офисных программ. На Windows можно установить LibreOffice для экспорта `.doc`, но импорт старого документа лучше выполнять через Word/Antiword, чтобы сохранить структуру таблиц.

## Как импортировать таблицу с фотографии

1. Откройте `assets/Шаблон_переводов_таджикский.xlsx` либо `assets/Шаблон_таджикский_без_номера.xlsx` (в интерфейсе **Шаблоны ▾**). Обе версии поддерживаются и для Word.
2. Заполните только обязательные данные: ФИО и класс. **Столбец «№» можно не заполнять или удалить полностью. ID в шаблоне не требуется.** Для пустого ухода/прибытия оставьте ячейку пустой; класс допускает форму `5`, `5А`, `10-Б`.
3. Телефон записывайте **как текст** (так сохраняются плюс и ведущие нули).
4. В программе: **Ученики → Импорт файла** → выберите `.xlsx`, `.docx` или `.doc` с реальной таблицей. Сканированные изображения и документы без таблицы не подходят.
5. Просмотрите предупреждения, ошибки и количество строк. При ошибках импорт **целиком отменяется**.
6. Повторный импорт того же таджикского файла пропускает уже обработанные строки. Экспорт полной таблицы с уже созданным ID нужен для корректировок существующих карточек; изменение исходной таджикской строки не следует использовать для обновления старых данных — исправляйте карточку или используйте полный экспорт.

Учебный год класса из импортированного листа считается **текущим** на момент импорта; дата исходного листа не определяется автоматически. Не используйте прошлогодние списки, пока не проверите классы.

## Как хранятся данные

| Система | Основной файл |
| --- | --- |
| Windows | `%LOCALAPPDATA%\MaktabOS\school.db` |
| macOS | `~/Library/Application Support/MaktabOS/school.db` |
| Linux | `~/.local/share/MaktabOS/school.db` |

При первом запуске MaktabOS проверяет прежнюю папку `%LOCALAPPDATA%\SchoolRegister` (и соответствующие каталоги macOS/Linux). Если найдена старая база, она копируется средствами SQLite в папку MaktabOS; исходная база остаётся нетронутой. Старые резервные копии `.srbak` по-прежнему можно восстановить. При обычной переустановке приложения с прежней учётной записью ОС сохранённая база автоматически используется снова. **Удаление папки пользовательских данных или отказ диска уничтожат сведения**, если отсутствуют отдельные копии. Делайте `.srbak` и переносите на защищённый внешний носитель. База SQLite сама не зашифрована: используйте шифрование диска и ограниченный доступ к компьютеру. Excel и Word файлы тоже не зашифрованы.

В приложении используется один локальный аккаунт администратора, отдельные аккаунты сотрудников и синхронизация между устройствами не поддерживаются. При первой настройке задайте пароль и ключ восстановления (не менее 10 символов), введя его дважды. Ключ не сохраняется в открытом виде. Кнопка «Изменить пароль» в главном окне требует текущий пароль; «Забыли пароль?» на экране входа позволяет создать новый пароль по ключу восстановления. Старые базы без такого ключа можно настроить после входа с существующим паролем. Для восстановления на другом устройстве задайте там тот же ключ отдельно: ключи и базы между устройствами автоматически не передаются. Потерянный ключ восстановления невозможно посмотреть или восстановить.

Если забыты и пароль, и ключ восстановления, запустите отдельный файл `MaktabOS-Recovery.exe` из любого каталога. Он автоматически выбирает стандартную базу MaktabOS текущего пользователя и позволяет указать файл вручную, если база находится в другом месте. Закройте приложение перед сбросом. Утилита создаёт резервную копию, задаёт новый пароль и ключ, сохраняя записи учеников; восстановить можно только базу, к которой у вас есть разрешённый доступ. Пароль зашифрованных `.srbak`-копий этой утилитой не меняется. Чтобы собрать EXE из исходников в Windows, запустите `scripts/build_recovery_windows.bat`; готовый файл будет в `dist/MaktabOS-Recovery.exe`.

При обновлении старой базы v1 программа добавляет новые столбцы и таблицы, **не удаляя старые карточки и пароль**. Перед обновлением рекомендуется сделать резервную копию из предыдущей версии. Старые события переводов, не записывавшиеся в v1, восстановить невозможно.

## Проверка и ограничения

В папке проекта: `python -m unittest discover -s tests -v` и `python -m compileall -q schoolregister`.

- Запросы и данные остаются на локальном устройстве. Никакого облака, синхронизации между компьютерами или разделения прав сотрудников нет.
- Не пересчитывает автоматически факт выпуска: это кадровое/школьное решение.
- История события содержит время внесения (UTC); реальная дата перевода заполняется только при наличии достоверных данных.
- Нет OCR фото и сканов. Таблицы должны быть редактируемыми, не фотографиями.
- **Windows `.exe` и интерактивный GUI должны быть проверены непосредственно на Windows перед использованием в школе.** Автоматические тесты проверяют ядро и форматы файлов, но не дают гарантии отсутствия проблем на конкретном устройстве.
- Перед эксплуатацией реального реестра школьников получите разрешение на обработку персональных данных, организуйте доступ сотрудников и политику хранения.

---

## English — MaktabOS offline school student management system

MaktabOS is a desktop **school management application**, **student information system (SIS)**, and **student database** for schools in Tajikistan and other schools that maintain Tajik- or Russian-language student registers. This offline school management system supports student record management, Excel import/export, and Word table import/export. Student data is stored locally in SQLite; the application does not require cloud accounts or an internet connection for normal use.

**Author: Abdullo Ashurov.**

### Download and install on Windows

Prebuilt Windows installer binaries are not included in this source repository. To create the installer, run `scripts/build_windows.bat` on Windows with Python and Inno Setup 6 installed. The generated file is `dist/MaktabOS-Setup-3.0.0.exe`. Its default installation folder is `Program Files\MaktabOS`; you can choose another location in the setup wizard. The installer can add Start menu and optional desktop shortcuts. Python is not required to run the installed Windows application.

### Run from source

The application is written in Python with PySide6 (Qt) and SQLite. For source-based development, install Python 3.11–3.14 and the project dependencies:

```powershell
python -m pip install -r requirements.txt
python -m schoolregister
```

On Windows, `scripts/run_windows.bat` starts the application. On macOS or Linux, run `bash scripts/run_linux_macos.sh` after installing the dependencies.

To build the Windows application and setup installer from source, run `scripts/build_windows.bat` on Windows. Python and Inno Setup 6 are required. The installer is written to `dist/MaktabOS-Setup-3.0.0.exe`.

### Student register and school management features

- Maintain student profiles with names, date of birth, grade and class section, shift, residential address, parent or guardian details, phone numbers, enrollment and leaving dates, status, notes, and previous or next school.
- Import editable school lists from Excel `.xlsx` and Word `.docx` tables. Legacy Word `.doc` is supported with compatible local conversion tools. The importer can read Tajik and Russian headings, different column orders, and Tajik address headings such as `Чои истикомат`, including spelling without diacritics.
- Import a Tajik school register with or without the optional row-number column. The row number is not used as a database ID. Tajik full-name and grade formats such as `5А` or `10-Б` are supported.
- Import all matching tables from a Word document. When a Tajik import contains an invalid optional phone number, the phone can be omitted with a warning while the student's other data, including residential address, is retained.
- Export the full 18-column student register or a compact six-column Tajik school list to Excel or Word. Full exports include stable record IDs for updating existing students; blank-template imports generate unique IDs for new records.
- Track incoming and outgoing school transfers in an append-only history. A leaving school sets the student's status to transferred. Transfer dates are not invented when they are absent from the source.
- View dashboard counts, student status, and grade distribution. Studying students are promoted for missed school years starting September 1; transferred and graduated students are not automatically promoted. The configured final grade is 11 or 12, and graduation remains a manual school decision.
- Search student names and addresses, filter by grade, restore deleted records from the recycle bin, review the audit log, and create or restore encrypted `.srbak` backups.

### Tajik Excel and Word templates

Ready-to-use templates are included in `assets/`:

- `Шаблон_учеников.xlsx` — full student register without an ID column.
- `Шаблон_переводов_таджикский.xlsx` — Tajik transfer list with an optional row-number column.
- `Шаблон_таджикский_без_номера.xlsx` — Tajik list without a row-number column.
- Matching `.docx` templates are also provided. A `.doc` sample is available for compatible Word workflows.

For reliable phone imports, format phone cells as text in the source spreadsheet so that plus signs and leading zeros are preserved. Scanned PDFs, photos, and image-only documents are not supported; convert or retype the data into an editable `.xlsx` or Word table first. The legacy binary `.xls` format is not supported.

### How to import a Tajik school list

1. Start with one of the Tajik Excel templates in `assets/`, or use an editable Word table with equivalent Tajik or Russian column headings. The optional `№` column may be left blank or omitted; it is never treated as a student ID.
2. Enter each student's full name and grade. Grade values can include a class section, for example `5А` or `10-Б`. Add a residential address if it is present in your source list; headings such as `Суроғаи истиқомат`, `Чои истикомат`, and `Адрес проживания` are recognized.
3. In MaktabOS, open **Students → Import file** and select an `.xlsx`, `.docx`, or supported `.doc` file containing a real editable table. A Word document can contain multiple school tables; all matching tables are read.
4. Review the import preview, warnings, and errors before confirming. A malformed optional phone number is omitted with a warning so other student information can still be imported. Other invalid rows may require correcting the source file before import.
5. Re-importing an unchanged Tajik list skips rows already imported. To update an existing student reliably, edit the student profile or use a full export that contains the student's ID; do not use a modified ID-free Tajik row as an update identifier.

Imported grades are assigned to the current academic year at import time; the program does not infer the year of the source list. Review grade values before importing an older school register.

### Local database, privacy, and recovery

The SQLite database is stored in the current user's data folder:

| Operating system | Database location |
| --- | --- |
| Windows | `%LOCALAPPDATA%\MaktabOS\school.db` |
| macOS | `~/Library/Application Support/MaktabOS/school.db` |
| Linux | `~/.local/share/MaktabOS/school.db` |

The application is designed for local, single-computer use. It does not provide cloud storage, automatic multi-device synchronization, or separate staff accounts. The SQLite database and exported Excel/Word files are not encrypted by the application; protect the device with operating-system disk encryption and access controls. Store encrypted backups on a separate, protected drive and keep the recovery key safe. A forgotten recovery key cannot be retrieved.

MaktabOS uses a local administrator account. During first-time setup, create a password and a recovery key of at least 10 characters and enter them twice. The recovery key is not stored in plain text. Change a password from the main window using the current password, or use **Forgot password?** on the sign-in screen with the recovery key. Older databases without a recovery key can be configured after signing in with the existing password. Passwords, recovery keys, and database files are not synchronized between computers.

On first launch, the application checks for an older `SchoolRegister` data folder and copies a legacy database with SQLite while leaving the original untouched. Normal reinstallation under the same operating-system account reuses the existing MaktabOS database. When upgrading an older database, MaktabOS adds the newer schema without deleting student records or the existing password. Create a backup before upgrading.

The separate `MaktabOS-Recovery.exe` utility can reset the local application password when the user has access to the database. Close MaktabOS before running recovery. The utility does not recover passwords for encrypted `.srbak` backup files.

### Tests and limitations

Run the project's tests and Python syntax check from the project folder:

```powershell
python -m unittest discover -s tests -v
python -m compileall -q schoolregister
```

MaktabOS does not perform OCR on photos or scans, infer student identity from a shared parent phone number, or automatically decide which students have graduated. Review imported student names, grades, and addresses before relying on the register. Obtain the required authorization for processing student personal data and define appropriate staff access and data-retention procedures.
