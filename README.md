# Dreamkas Receipt Tool

**Dreamkas Receipt Tool** is a Python utility for preparing, fiscalizing, storing, refunding, and exporting receipts through the Dreamkas API.

**Dreamkas Receipt Tool** — это Python-утилита для подготовки, фискализации, хранения, возврата и экспорта кассовых чеков через Dreamkas API.

---

**Current version / Актуальная версия:** v6.9

---

## Languages / Языки

- [Русская версия](#русская-версия)
- [English version](#english-version)

---

# Русская версия

## Назначение

**Dreamkas Receipt Tool** предназначен для работы с кассами Dreamkas через API Кабинета Дримкас.

Утилита позволяет заполнить чек в Excel-шаблоне, выбрать магазин и кассу, сформировать предчек, отправить чек на фискализацию, дождаться результата, сохранить данные в SQLite, сформировать TXT/PDF-чек и QR-код, а также выполнить полный или частичный возврат ранее пробитого чека.

Программа подходит для небольших сервисных компаний, мастерских, внутренних операторов, выездных инженеров и сценариев, где не нужна полноценная POS-система, но требуется контролируемая фискализация чеков.

---

## Основные возможности

- Выбор магазина из Dreamkas API.
- Выбор кассы из Dreamkas API.
- Заполнение чека через Excel-шаблон.
- Автоматическое открытие Excel-шаблона перед созданием чека.
- Автоматическое восстановление или создание Excel-шаблона, если файл отсутствует.
- Встраивание Excel-шаблона внутрь EXE при сборке.
- Предпросмотр чека в терминале перед отправкой.
- Отправка задания на фискализацию в Dreamkas.
- Ожидание результата операции.
- Сохранение полного предчека и результата в SQLite.
- Защита от повторной отправки одинакового чека.
- Полный возврат чека.
- Частичный возврат отдельных позиций.
- Сохранение TXT-чека.
- Сохранение QR-кода.
- Генерация PDF-чека.
- SMTP-отправка чека покупателю.
- Зашифрованное хранение SMTP-пароля в `settings.txt`.
- Логирование запросов, ответов и ошибок.
- Проверка и автоматическая установка зависимостей при запуске.
- Автоматическое восстановление `pip` через `ensurepip` или `get-pip.py`.
- Диагностика установленного Python через `diagnose_python.bat`.
- Простая GUI-оболочка для запуска утилиты и открытия рабочих папок.
- Сборка в EXE через PyInstaller.

---

## Что добавлено в актуальной версии

- Excel-шаблон встраивается внутрь EXE при сборке через PyInstaller.
- При запуске EXE программа проверяет наличие `dreamkas_receipt_template.xlsx` рядом с программой.
- Если шаблона нет, программа восстанавливает его из встроенного ресурса EXE.
- Если встроенный ресурс недоступен, программа создаёт новый Excel-шаблон программно.
- GUI-оболочка также умеет восстанавливать или создавать Excel-шаблон.
- `build_exe.bat` добавляет Excel-шаблон в EXE через параметр `--add-data`.
- `start_dreamkas.bat`, `start_gui.bat`, `install_dependencies.bat` и `build_exe.bat` проверяют наличие `pip`.
- Если `pip` отсутствует, BAT-файлы пробуют восстановить его через `ensurepip`.
- Если `ensurepip` недоступен или не сработал, BAT-файлы скачивают официальный `get-pip.py` и устанавливают `pip`.
- Добавлен `diagnose_python.bat` для диагностики Python, `pip`, `ensurepip` и импортов зависимостей.
- EXE можно запускать в новой папке без ручного копирования Excel-шаблона.

---

## Как работает фискализация

Общий сценарий:

```text
Excel-шаблон
    ↓
Предчек в терминале
    ↓
Подтверждение пользователя
    ↓
POST /api/receipts в Dreamkas
    ↓
Ожидание операции
    ↓
Получение результата
    ↓
SQLite + TXT + QR + PDF
```

Перед отправкой в Dreamkas программа показывает пользователю предчек и требует подтверждение. Без подтверждения чек не отправляется.

---

## Как работает возврат

В режиме возврата программа читает SQLite-базу и показывает список успешно фискализированных чеков.

Доступные варианты:

```text
1. Вернуть весь чек
2. Вернуть часть позиций
0. Отмена
```

Если в исходном чеке только одна позиция, частичный возврат не предлагается. В этом случае доступен только возврат всего чека или отмена.

Для частичного возврата позиции выбираются через запятую:

```text
1,3,5
```

После выбора программа повторно показывает чек и отмечает позиции, выбранные к возврату.

---

## Защита от дублей

Перед отправкой чек сохраняется в SQLite как полный слепок предчека.

Сохраняются:

- тип операции;
- тип оплаты;
- email покупателя;
- телефон покупателя;
- система налогообложения;
- имя кассира;
- выбранный магазин;
- выбранная касса;
- список позиций;
- суммы;
- JSON-запрос;
- hash предчека.

Если программа обнаруживает похожий чек, она предупреждает пользователя перед повторной отправкой.

Это помогает избежать случайной повторной фискализации одного и того же чека.

---

## Очистка Excel после успешного чека

После успешной фискализации Excel-шаблон автоматически очищается.

Очищаются:

- тип оплаты;
- email покупателя;
- телефон покупателя;
- система налогообложения;
- имя кассира;
- товарные позиции.

Шапка и подсказки остаются.

Настройка:

```text
CLEAR_EXCEL_AFTER_SUCCESS = 1
```

Чтобы отключить очистку:

```text
CLEAR_EXCEL_AFTER_SUCCESS = 0
```

---

## Встроенный Excel-шаблон для EXE

При обычном запуске из исходников рядом с программой используется файл:

```text
dreamkas_receipt_template.xlsx
```

При сборке в EXE шаблон Excel дополнительно встраивается внутрь исполняемого файла как ресурс PyInstaller.

При запуске программа проверяет наличие файла:

```text
dreamkas_receipt_template.xlsx
```

рядом с программой или EXE.

Если файл найден — используется он.

Если файл не найден, программа автоматически:

```text
1. пробует восстановить dreamkas_receipt_template.xlsx из встроенного ресурса EXE;
2. если встроенный ресурс недоступен — создаёт новый Excel-шаблон программно.
```

Таким образом, EXE можно перенести в новую папку и запустить даже без заранее скопированного Excel-шаблона. Шаблон будет создан автоматически при первом запуске.

Это работает как для основной консольной программы, так и для GUI-оболочки.

---

## Имя кассира

Имя кассира указывается в Excel и сохраняется локально:

- в предчеке;
- в SQLite;
- в TXT-чеке;
- в PDF-чеке.

Имя кассира не отправляется в `POST /api/receipts`, потому что этот endpoint Dreamkas отклоняет поле `cashier`.

Если в фискальном чеке на кассе отображается кассир `Администратор`, его нужно менять в настройках кассы, профиля кассира или Кабинета Дримкас.

---

## Структура файлов проекта

```text
dreamkas_receipt.py              основной консольный скрипт
dreamkas_gui.py                  простая GUI-оболочка
dreamkas_receipt_template.xlsx   Excel-шаблон чека
settings.example.txt             пример файла настроек
requirements.txt                 зависимости Python
start_dreamkas.bat               запуск основной программы
start_gui.bat                    запуск GUI-оболочки
install_dependencies.bat         установка зависимостей
diagnose_python.bat              диагностика Python и pip
test_smtp.py                     проверка SMTP
test_smtp.bat                    запуск SMTP-теста
build_exe.bat                    сборка EXE через PyInstaller
SMTP_SETUP.txt                   инструкция по SMTP
README.md                        описание проекта
CHANGELOG.md                     список изменений
```

Во время работы программа создаёт папки:

```text
db/
logs/
receipts_txt/
receipts_qr/
receipts_pdf/
```

---

## Требования

- Windows 10/11.
- Python 3.10 или новее.
- Доступ к интернету.
- Токен Dreamkas API.
- Подключенная и зарегистрированная касса Dreamkas.
- Excel или другой редактор `.xlsx`.

---

## Зависимости

Зависимости указаны в `requirements.txt`:

```text
requests
openpyxl
qrcode[pil]
Pillow
reportlab
pyinstaller
```

При запуске `start_dreamkas.bat` программа проверяет зависимости и автоматически устанавливает отсутствующие пакеты.

Для ручной установки:

```bat
install_dependencies.bat
```

Или:

```bat
python -m pip install -r requirements.txt
```

---

## Восстановление pip

Если Python установлен без `pip`, BAT-файлы сначала пробуют:

```bat
python -m ensurepip --upgrade
```

Если `ensurepip` недоступен или завершился ошибкой, BAT автоматически скачивает и запускает официальный bootstrap-скрипт:

```text
https://bootstrap.pypa.io/get-pip.py
```

Также добавлен файл:

```text
diagnose_python.bat
```

Он показывает активный путь Python, доступные версии Python, состояние `ensurepip`, состояние `pip` и проверку импортов зависимостей.

---

## Первый запуск

1. Распакуйте проект в отдельную папку.
2. Запустите:

```bat
start_dreamkas.bat
```

3. Если `DREAMKAS_TOKEN` отсутствует, программа попросит ввести токен.
4. Программа загрузит список магазинов из Dreamkas API.
5. Выберите магазин.
6. Программа загрузит список касс.
7. Выберите кассу.
8. Выберите режим работы: фискализация, возврат, история или настройки.

---

## Формат Excel-шаблона

Excel-шаблон имеет следующую структуру:

```text
1 строка: Тип оплаты
2 строка: Email покупателя
3 строка: Телефон покупателя
4 строка: Система налогообложения
5 строка: Имя кассира
6 строка: Шапка таблицы
7 строка и ниже: Позиции чека
```

Формат позиций:

```text
A: Наименование
B: Тип — Услуга = 1, Товар = 0
C: Количество
D: Цена в рублях
E: Ставка НДС
```

Пример:

```text
Ремонт кофемашины | 1 | 1 | 2500.00 | 0
```

---

## Тип оплаты

Поддерживаемые значения:

```text
Наличные
Безнал
```

Они преобразуются в значения Dreamkas API:

```text
CASH
CASHLESS
```

---

## Система налогообложения

Примеры значений:

```text
DEFAULT
SIMPLE
SIMPLE_WO
AGRICULT
PATENT
```

Значение должно соответствовать системе налогообложения, зарегистрированной на выбранной кассе.

---

## НДС

Для работы без НДС можно указать:

```text
0
Без НДС
```

Это преобразуется в:

```text
NDS_NO_TAX
```

---

## Файл настроек

Рабочий файл настроек называется:

```text
settings.txt
```

Пример:

```text
DREAMKAS_TOKEN = your_dreamkas_token
SHOP_ID = 212962
SHOP_NAME = Мастерская
DEVICE_ID = 160501
DEVICE_NAME = Касса

API_BASE_URL = https://kabinet.dreamkas.ru/api
OPERATION_TYPE = SALE
TIMEOUT_MINUTES = 5
POLL_INTERVAL_SECONDS = 20
REQUEST_TIMEOUT_SECONDS = 30

DUPLICATE_WINDOW_MINUTES = 10
DUPLICATE_LOOKBACK_DAYS = 30

CLEAR_EXCEL_AFTER_SUCCESS = 1
PDF_ENABLED = 1

EMAIL_ENABLED_AFTER_SUCCESS = 0
SMTP_HOST =
SMTP_PORT = 465
SMTP_USER =
SMTP_PASSWORD_ENC =
SMTP_FROM =
SMTP_SSL = 1
SMTP_STARTTLS = 0
SMTP_VERIFY_CERT = 1
```

---

## SMTP

Утилита может отправлять чек покупателю по email после успешной фискализации.

Для включения:

```text
EMAIL_ENABLED_AFTER_SUCCESS = 1
```

Пример SSL на порту 465:

```text
SMTP_HOST = mail.example.com
SMTP_PORT = 465
SMTP_USER = receipt@example.com
SMTP_FROM = receipt@example.com
SMTP_SSL = 1
SMTP_STARTTLS = 0
SMTP_VERIFY_CERT = 1
```

Пример STARTTLS на порту 587:

```text
SMTP_HOST = mail.example.com
SMTP_PORT = 587
SMTP_USER = receipt@example.com
SMTP_FROM = receipt@example.com
SMTP_SSL = 0
SMTP_STARTTLS = 1
SMTP_VERIFY_CERT = 1
```

Если SMTP-сервер использует некорректный сертификат, для теста можно отключить проверку:

```text
SMTP_VERIFY_CERT = 0
```

Для боевой эксплуатации рекомендуется настроить корректный SSL-сертификат.

Проверка SMTP без пробития чека:

```bat
test_smtp.bat
```

Или:

```bat
python test_smtp.py --to client@example.com
```

---

## Зашифрованное хранение SMTP-пароля

SMTP-пароль хранится в `settings.txt` в зашифрованном виде:

```text
SMTP_PASSWORD_ENC = enc-v1:...
```

Ключ шифрования строится на основе:

```text
DREAMKAS_TOKEN
```

Это позволяет переносить папку проекта между компьютерами вместе с настройками.

Важно: это не защищает от полной кражи `settings.txt`, потому что в этом же файле хранится `DREAMKAS_TOKEN`. Такая схема защищает от случайного просмотра пароля, но не заменяет полноценное хранилище секретов.

---

## PDF, TXT и QR

После успешной фискализации создаются файлы:

```text
receipts_txt/Receipt(YYYY-MM-DD-HH-mm-ss).txt
receipts_qr/Receipt(YYYY-MM-DD-HH-mm-ss).png
receipts_pdf/Receipt(YYYY-MM-DD-HH-mm-ss).pdf
```

Если PDF не нужен:

```text
PDF_ENABLED = 0
```

---

## Логи

Логи сохраняются в папке:

```text
logs/
```

В логах сохраняются:

- запросы к Dreamkas API;
- ответы Dreamkas API;
- ошибки валидации;
- ошибки SMTP;
- ошибки PDF;
- технические исключения.

---

## GUI

Простая GUI-оболочка запускается через:

```bat
start_gui.bat
```

GUI позволяет:

- выбрать Excel-файл;
- открыть Excel;
- запустить консольную программу;
- открыть папку с PDF-чеками;
- открыть папку с логами;
- открыть папку с базой.

Основная логика фискализации остаётся в консольном сценарии.

---

## Сборка в EXE

Для сборки используется PyInstaller:

```bat
build_exe.bat
```

`build_exe.bat` собирает консольную и GUI-версии, а также встраивает Excel-шаблон внутрь EXE через PyInstaller:

```bat
--add-data "dreamkas_receipt_template.xlsx;."
```

После сборки EXE-файлы будут в папке:

```text
dist/
```

Обычно создаются:

```text
DreamkasReceipt.exe
DreamkasReceiptGUI.exe
```

Рядом с EXE желательно хранить:

```text
settings.txt
```

Excel-шаблон `dreamkas_receipt_template.xlsx` можно положить рядом с EXE вручную, но это не обязательно. Если шаблона рядом нет, программа восстановит его из встроенного ресурса EXE или создаст новый шаблон автоматически.

---

## Безопасность и GitHub

Не рекомендуется публиковать в GitHub:

```text
settings.txt
db/
logs/
receipts_txt/
receipts_qr/
receipts_pdf/
```

Пример `.gitignore`:

```gitignore
settings.txt
db/
logs/
receipts_txt/
receipts_qr/
receipts_pdf/
__pycache__/
*.pyc
dist/
build/
*.spec
```

---

## Важное предупреждение

Эта утилита не является официальным продуктом Dreamkas. Перед использованием в боевой работе необходимо протестировать сценарии на небольших тестовых чеках и проверить соответствие требованиям вашей бухгалтерии, ОФД и законодательства.

---


## Автоматическая установка Python

Если на компьютере нет нормального Python 3.10+ или команда `python` указывает на сломанный Windows alias, BAT-файлы пробуют установить Python автоматически через `winget`.

Порядок действий:

```text
1. Проверить наличие Python 3.10+
2. Если Python не найден — проверить наличие winget
3. Если winget доступен — установить Python 3.13
4. Если Python 3.13 не установился — попробовать Python 3.12
5. После установки снова найти Python
6. Проверить pip и зависимости
```

Отдельный файл для установки/проверки Python:

```text
install_python.bat
```

Если `winget` недоступен, установите Python вручную с сайта:

```text
https://www.python.org/downloads/windows/
```

При установке обязательно включите:

```text
pip
Add python.exe to PATH
Python Launcher
```

---

# English version

## Purpose

**Dreamkas Receipt Tool** is a Python utility for working with Dreamkas cash registers through the Dreamkas Cabinet API.

The tool allows an operator to fill out a receipt in an Excel template, choose a shop and device, preview the receipt, submit it for fiscalization, wait for the result, store the data in SQLite, generate TXT/PDF receipt files and a QR code, and perform full or partial refunds.

It is intended for small service companies, repair workshops, internal operators, field engineers, and cases where a full POS system is not required.

---

## Main features

- Shop selection through the Dreamkas API.
- Device selection through the Dreamkas API.
- Excel-based receipt template.
- Automatic Excel template opening before receipt creation.
- Automatic Excel template restore or creation if the file is missing.
- Excel template embedding into the EXE build.
- Receipt preview in the terminal.
- Fiscalization task submission to Dreamkas.
- Operation status polling.
- Full precheck and result storage in SQLite.
- Duplicate receipt warning.
- Full receipt refund.
- Partial item refund.
- TXT receipt export.
- QR code export.
- PDF receipt generation.
- Optional email sending via SMTP.
- Encrypted SMTP password storage in `settings.txt`.
- Request, response, and error logging.
- Dependency checking and installation on startup.
- Automatic `pip` recovery via `ensurepip` or `get-pip.py`.
- Python diagnostics through `diagnose_python.bat`.
- Simple GUI launcher.
- EXE build support via PyInstaller.

---

## Current version additions

- The Excel template is embedded into the EXE during the PyInstaller build.
- On startup, the EXE checks whether `dreamkas_receipt_template.xlsx` exists next to the program.
- If the template is missing, the program restores it from the embedded EXE resource.
- If the embedded resource is unavailable, the program creates a new Excel template programmatically.
- The GUI launcher can also restore or create the Excel template.
- `build_exe.bat` includes the Excel template in the EXE using `--add-data`.
- `start_dreamkas.bat`, `start_gui.bat`, `install_dependencies.bat`, and `build_exe.bat` check whether `pip` is available.
- If `pip` is missing, the BAT files try to restore it using `ensurepip`.
- If `ensurepip` is unavailable or fails, the BAT files download the official `get-pip.py` bootstrap script and install `pip`.
- `diagnose_python.bat` was added for Python, `pip`, `ensurepip`, and dependency import diagnostics.
- The EXE can be launched in a new folder without manually copying the Excel template.

---

## Fiscalization workflow

```text
Excel template
    ↓
Terminal precheck
    ↓
User confirmation
    ↓
POST /api/receipts to Dreamkas
    ↓
Operation polling
    ↓
Receipt result
    ↓
SQLite + TXT + QR + PDF
```

The tool never submits a receipt without showing a preview and asking the user for confirmation.

---

## Refund workflow

In refund mode, the tool reads successful receipts from the SQLite database.

Available actions:

```text
1. Full refund
2. Partial item refund
0. Cancel
```

If the original receipt contains only one item, partial refund is not offered.

For a partial refund, item numbers are entered as a comma-separated list:

```text
1,3,5
```

The tool then shows the receipt again and marks the selected items before sending the refund.

---

## Duplicate protection

Before submission, the full precheck snapshot is stored in SQLite and a hash is calculated.

The snapshot includes:

- operation type;
- payment type;
- buyer email;
- buyer phone;
- tax mode;
- cashier name;
- selected shop;
- selected device;
- line items;
- totals;
- request JSON;
- precheck hash.

If a similar receipt is found, the tool warns the user before submitting it again.

---

## Excel cleanup after success

After successful fiscalization, the Excel template is cleaned automatically.

The tool clears:

- payment type;
- buyer email;
- buyer phone;
- tax mode;
- cashier name;
- receipt items.

Header rows and hints remain unchanged.

Setting:

```text
CLEAR_EXCEL_AFTER_SUCCESS = 1
```

To disable cleanup:

```text
CLEAR_EXCEL_AFTER_SUCCESS = 0
```

---

## Embedded Excel template for EXE

When running from source, the program normally uses this file next to the script:

```text
dreamkas_receipt_template.xlsx
```

When building an EXE, the Excel template is also embedded into the executable as a PyInstaller resource.

On startup, the program checks whether this file exists:

```text
dreamkas_receipt_template.xlsx
```

next to the program or EXE.

If the file exists, it is used.

If the file is missing, the program automatically:

```text
1. tries to restore dreamkas_receipt_template.xlsx from the embedded EXE resource;
2. if the embedded resource is unavailable, creates a new Excel template programmatically.
```

As a result, the EXE can be moved to a new folder and launched even without manually copying the Excel template. The template will be created automatically on first startup.

This works both for the console tool and for the GUI launcher.

---

## Cashier name

The cashier name is read from Excel and saved locally:

- in the precheck;
- in SQLite;
- in TXT files;
- in PDF files.

The cashier name is not sent to `POST /api/receipts`, because the Dreamkas endpoint rejects the `cashier` field.

If the fiscal receipt prints `Administrator`, configure the cashier on the cash register, in Dreamkas Cabinet, or in the device settings.

---

## Project structure

```text
dreamkas_receipt.py              main console script
dreamkas_gui.py                  simple GUI launcher
dreamkas_receipt_template.xlsx   Excel receipt template
settings.example.txt             example settings file
requirements.txt                 Python dependencies
start_dreamkas.bat               console launcher
start_gui.bat                    GUI launcher
install_dependencies.bat         dependency installer
install_python.bat               Python installer/checker
diagnose_python.bat              Python and pip diagnostics
test_smtp.py                     SMTP test script
test_smtp.bat                    SMTP test launcher
build_exe.bat                    PyInstaller EXE builder
SMTP_SETUP.txt                   SMTP setup guide
README.md                        project description
CHANGELOG.md                     change log
```

Runtime folders:

```text
db/
logs/
receipts_txt/
receipts_qr/
receipts_pdf/
```

---

## Requirements

- Windows 10/11.
- Python 3.10 or newer.
- Internet access.
- Dreamkas API token.
- Registered and connected Dreamkas cash register.
- Excel or another `.xlsx` editor.

---

## Dependencies

The dependencies are listed in `requirements.txt`:

```text
requests
openpyxl
qrcode[pil]
Pillow
reportlab
pyinstaller
```

`start_dreamkas.bat` checks the dependencies and installs missing packages automatically.

Manual installation:

```bat
install_dependencies.bat
```

Or:

```bat
python -m pip install -r requirements.txt
```

---

## Python / pip bootstrap

If Python is installed without `pip`, the launcher first tries:

```bat
python -m ensurepip --upgrade
```

If `ensurepip` is unavailable or fails, the launcher downloads and runs the official `get-pip.py` bootstrap script from:

```text
https://bootstrap.pypa.io/get-pip.py
```

The package also includes:

```text
diagnose_python.bat
```

It prints the active Python path, Python versions, `ensurepip` status, `pip` status, and dependency import status.

---

## First run

1. Extract the project into a separate folder.
2. Run:

```bat
start_dreamkas.bat
```

3. If `DREAMKAS_TOKEN` is missing, the tool asks for it.
4. The tool loads shops from the Dreamkas API.
5. Select a shop.
6. The tool loads cash registers.
7. Select a cash register.
8. Select the operation mode: fiscalization, refund, history, or settings.

---

## Excel template format

```text
Row 1: Payment type
Row 2: Buyer email
Row 3: Buyer phone
Row 4: Tax mode
Row 5: Cashier name
Row 6: Table header
Row 7 and below: Receipt items
```

Item columns:

```text
A: Name
B: Type — Service = 1, Product = 0
C: Quantity
D: Price in rubles
E: VAT rate
```

Example:

```text
Coffee machine repair | 1 | 1 | 2500.00 | 0
```

---

## Payment type

Supported values:

```text
Cash
Cashless
```

Russian template values:

```text
Наличные
Безнал
```

Converted Dreamkas API values:

```text
CASH
CASHLESS
```

---

## Tax mode

Examples:

```text
DEFAULT
SIMPLE
SIMPLE_WO
AGRICULT
PATENT
```

The value must match the tax mode registered on the selected cash register.

---

## VAT

For no VAT:

```text
0
No VAT
Без НДС
```

Converted value:

```text
NDS_NO_TAX
```

---

## Settings file

The working settings file is:

```text
settings.txt
```

Example:

```text
DREAMKAS_TOKEN = your_dreamkas_token
SHOP_ID = 212962
SHOP_NAME = Workshop
DEVICE_ID = 160501
DEVICE_NAME = Cash register

API_BASE_URL = https://kabinet.dreamkas.ru/api
OPERATION_TYPE = SALE
TIMEOUT_MINUTES = 5
POLL_INTERVAL_SECONDS = 20
REQUEST_TIMEOUT_SECONDS = 30

DUPLICATE_WINDOW_MINUTES = 10
DUPLICATE_LOOKBACK_DAYS = 30

CLEAR_EXCEL_AFTER_SUCCESS = 1
PDF_ENABLED = 1

EMAIL_ENABLED_AFTER_SUCCESS = 0
SMTP_HOST =
SMTP_PORT = 465
SMTP_USER =
SMTP_PASSWORD_ENC =
SMTP_FROM =
SMTP_SSL = 1
SMTP_STARTTLS = 0
SMTP_VERIFY_CERT = 1
```

---

## SMTP

The tool can send an email to the buyer after successful fiscalization.

Enable:

```text
EMAIL_ENABLED_AFTER_SUCCESS = 1
```

SSL on port 465:

```text
SMTP_HOST = mail.example.com
SMTP_PORT = 465
SMTP_USER = receipt@example.com
SMTP_FROM = receipt@example.com
SMTP_SSL = 1
SMTP_STARTTLS = 0
SMTP_VERIFY_CERT = 1
```

STARTTLS on port 587:

```text
SMTP_HOST = mail.example.com
SMTP_PORT = 587
SMTP_USER = receipt@example.com
SMTP_FROM = receipt@example.com
SMTP_SSL = 0
SMTP_STARTTLS = 1
SMTP_VERIFY_CERT = 1
```

If the SMTP server uses an invalid certificate, verification can be temporarily disabled:

```text
SMTP_VERIFY_CERT = 0
```

For production use, a valid SMTP certificate is recommended.

SMTP test:

```bat
test_smtp.bat
```

Or:

```bat
python test_smtp.py --to client@example.com
```

---

## Encrypted SMTP password

The SMTP password is stored in `settings.txt` in encrypted form:

```text
SMTP_PASSWORD_ENC = enc-v1:...
```

The encryption key is derived from:

```text
DREAMKAS_TOKEN
```

This allows the project folder to be moved together with its settings.

Important: this does not protect against full `settings.txt` compromise, because the same file contains `DREAMKAS_TOKEN`. It protects against casual password viewing, but it is not a replacement for a dedicated secret storage system.

---

## PDF, TXT, and QR

After successful fiscalization:

```text
receipts_txt/Receipt(YYYY-MM-DD-HH-mm-ss).txt
receipts_qr/Receipt(YYYY-MM-DD-HH-mm-ss).png
receipts_pdf/Receipt(YYYY-MM-DD-HH-mm-ss).pdf
```

Disable PDF:

```text
PDF_ENABLED = 0
```

---

## Logs

Logs are stored in:

```text
logs/
```

Logs include:

- Dreamkas API requests;
- Dreamkas API responses;
- validation errors;
- SMTP errors;
- PDF errors;
- technical exceptions.

---

## GUI

The simple GUI launcher is started with:

```bat
start_gui.bat
```

The GUI can:

- select an Excel file;
- open Excel;
- start the console tool;
- open the PDF receipt folder;
- open the log folder;
- open the database folder.

The main fiscalization workflow remains in the console tool.

---

## EXE build

The project can be built with PyInstaller:

```bat
build_exe.bat
```

`build_exe.bat` builds both the console and GUI versions and embeds the Excel template into the EXE using PyInstaller:

```bat
--add-data "dreamkas_receipt_template.xlsx;."
```

Output folder:

```text
dist/
```

Typical output files:

```text
DreamkasReceipt.exe
DreamkasReceiptGUI.exe
```

It is recommended to keep this file near the EXE:

```text
settings.txt
```

The Excel template `dreamkas_receipt_template.xlsx` may also be placed next to the EXE manually, but it is not required. If the template is missing, the program will restore it from the embedded EXE resource or create a new template automatically.

---

## Security and GitHub

Do not commit these files and folders to GitHub:

```text
settings.txt
db/
logs/
receipts_txt/
receipts_qr/
receipts_pdf/
```

Example `.gitignore`:

```gitignore
settings.txt
db/
logs/
receipts_txt/
receipts_qr/
receipts_pdf/
__pycache__/
*.pyc
dist/
build/
*.spec
```

---

## Important notice

This utility is not an official Dreamkas product. Before production use, test the workflow with small test receipts and verify compliance with your accounting, fiscal data operator, and legal requirements.

---

## License

Add your preferred license here, for example:

```text
MIT
Proprietary
Internal use only
```

---

## Author

Project owner: `@darkatrium`


## Automatic Python installation

If the computer does not have a valid Python 3.10+ installation or the `python` command points to a broken Windows alias, the BAT launchers try to install Python automatically using `winget`.

Workflow:

```text
1. Check for Python 3.10+
2. If Python is missing, check for winget
3. If winget is available, install Python 3.13
4. If Python 3.13 fails, try Python 3.12
5. Detect Python again after installation
6. Check pip and dependencies
```

Dedicated Python installer/checker:

```text
install_python.bat
```

If `winget` is unavailable, install Python manually from:

```text
https://www.python.org/downloads/windows/
```

During installation, enable:

```text
pip
Add python.exe to PATH
Python Launcher
```

---
