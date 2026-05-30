Dreamkas Receipt Tool v6.1

Что добавлено:
1. Папки:
   db/
   logs/
   receipts_txt/
   receipts_qr/
   receipts_pdf/

2. SQLite теперь хранит не только результат Dreamkas, но и полный предчек:
   precheck_json
   precheck_hash
   excel_file_path

3. Защита от дублей:
   Новый чек сравнивается с сохраненными предчеками в SQLite.
   Если найден такой же предчек, программа предупреждает и спрашивает,
   пробивать чек или отменить.

4. Очистка Excel:
   После успешной фискализации SALE программа очищает Excel-шаблон:
   поля покупателя/оплаты/СНО/кассира и список позиций.
   Управляется настройкой:
   CLEAR_EXCEL_AFTER_SUCCESS = 1

5. PDF:
   PDF-чек сохраняется в receipts_pdf, если:
   PDF_ENABLED = 1

6. Возвраты:
   Полный и частичный возврат по успешным чекам из SQLite.

7. История и настройки:
   В меню программы есть история чеков и настройки.

8. GUI:
   start_gui.bat запускает простую графическую оболочку.
   Основная безопасная логика по-прежнему работает в консольном окне.

9. Сборка в EXE:
   build_exe.bat соберет:
   dist/DreamkasReceipt.exe
   dist/DreamkasReceiptGUI.exe

Запуск:
   start_dreamkas.bat

GUI:
   start_gui.bat

Сборка EXE:
   build_exe.bat

Важно:
- settings.txt можно скопировать из старой рабочей папки.
- Если Excel открыт в момент очистки, Windows может заблокировать сохранение.
  Тогда программа покажет предупреждение, а чек и данные в SQLite всё равно сохранятся.


SMTP
====

Для настройки отправки чеков на email см. файл:

SMTP_SETUP.txt

Для проверки SMTP без пробития чека используйте:

test_smtp.bat

или команду:

C:\Python314\python.exe test_smtp.py --to client@example.ru


v6.3 — SMTP password encryption
===============================

SMTP password is no longer stored as plain SMTP_PASSWORD.
The program stores it as:

SMTP_PASSWORD_ENC = enc-v1:...

The encryption key is derived from DREAMKAS_TOKEN, so the whole folder can be
moved to another computer together with settings.txt.

Important: this protects against casual password viewing, not against theft of
the whole settings.txt, because settings.txt also contains DREAMKAS_TOKEN.

If the Dreamkas token is rejected by API with 401/403 during startup, the program
clears all SMTP_* settings and disables email sending:

EMAIL_ENABLED_AFTER_SUCCESS = 0

If you change DREAMKAS_TOKEN from the settings menu, SMTP settings are also
cleared because the decryption key changes.

To set SMTP password again, run:

test_smtp.bat

or configure SMTP in the program menu. The password will be requested once and
saved encrypted as SMTP_PASSWORD_ENC.


v6.5 — исправление отправки кассира
===================================

Dreamkas API /api/receipts вернул ошибку:

E_VALIDATION_OBJECT_ALLOW_UNKNOWN
"cashier" is not allowed

Поэтому имя кассира из Excel больше НЕ отправляется в JSON Dreamkas.
Оно по-прежнему сохраняется локально:
- в предчеке;
- в SQLite;
- в TXT/PDF-файлах чека.

Если на самой кассе в фискальном чеке отображается "Администратор",
это нужно менять в настройках кассы/кабинета Dreamkas, а не через поле
"cashier" в POST /api/receipts — этот endpoint его не принимает.

Также защита от дублей теперь не учитывает простые записи draft, которые могли
остаться после ошибки валидации. Дублями считаются только PENDING,
IN_PROGRESS и SUCCESS.


v6.5 — проверка и установка зависимостей при запуске
====================================================

start_dreamkas.bat теперь перед запуском программы проверяет:
- requests
- openpyxl
- qrcode
- Pillow / PIL
- reportlab

Если чего-то нет, BAT автоматически запускает:

python -m pip install -r requirements.txt

Также добавлен файл:

install_dependencies.bat

Его можно запустить вручную, если нужно принудительно установить или обновить все зависимости.


v6.6 — встроенный Excel-шаблон для EXE
======================================

Если программа собрана в EXE, Excel-шаблон добавляется внутрь EXE через PyInstaller:

--add-data "dreamkas_receipt_template.xlsx;."

При запуске программа проверяет наличие файла:

dreamkas_receipt_template.xlsx

рядом с EXE. Если файла нет, программа:
1. восстанавливает его из встроенного ресурса EXE;
2. если ресурс недоступен — создает новый шаблон программно.

Теперь EXE можно запускать даже если рядом нет Excel-шаблона: файл будет создан автоматически.


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


## Улучшенный поиск Python

Установщик Python через `winget` часто ставит интерпретатор в пользовательскую папку:

```text
%LOCALAPPDATA%\Programs\Python\Python313\python.exe
```

Поэтому BAT-файлы проверяют не только команды `python` и `py`, но и типовые пути установки:

```text
C:\Python313\python.exe
%LOCALAPPDATA%\Programs\Python\Python313\python.exe
%ProgramFiles%\Python313\python.exe
```

Также добавлен резервный поиск `python.exe` через PowerShell в папках установки Python.

---


## Improved Python detection

Python installed through `winget` is often placed into the per-user folder:

```text
%LOCALAPPDATA%\Programs\Python\Python313\python.exe
```

Therefore BAT launchers now check not only `python` and `py` commands, but also common installation paths:

```text
C:\Python313\python.exe
%LOCALAPPDATA%\Programs\Python\Python313\python.exe
%ProgramFiles%\Python313\python.exe
```

A PowerShell fallback search for `python.exe` in common Python installation folders is also included.

---


## Ручной путь к Python

Если автоматический поиск Python не смог найти установленный `python.exe`, BAT-файл попросит вручную указать полный путь к интерпретатору.

Пример:

```text
C:\Users\YOUR_USER\AppData\Local\Programs\Python\Python313\python.exe
```

После успешной проверки путь сохраняется в файл:

```text
python_path.txt
```

При следующих запусках программа сначала проверяет путь из `python_path.txt`.

---


## Manual Python path

If automatic Python detection cannot find the installed `python.exe`, the BAT launcher asks for the full path to the interpreter.

Example:

```text
C:\Users\YOUR_USER\AppData\Local\Programs\Python\Python313\python.exe
```

After successful validation, the path is saved to:

```text
python_path.txt
```

On the next runs, the launcher checks `python_path.txt` first.

---


## PDF в формате кассовой ленты

PDF-чек формируется не как обычный отчёт A4, а как узкая кассовая лента шириной 80 мм.

В PDF выводятся основные реквизиты кассового чека:

```text
КАССОВЫЙ ЧЕК
признак расчёта
продавец / ИНН
адрес или место расчёта
дата и время
смена и номер документа, если они есть в ответе Dreamkas
система налогообложения
кассир
email/телефон покупателя
позиции
количество, цена, сумма
НДС
итог
форма оплаты
ФН / ФД / ФП
РН ККТ / ЗН ККТ, если они есть в ответе Dreamkas
QR-код проверки
ссылка проверки чека
```

QR-код в PDF выводится размером 25×25 мм, чтобы быть не меньше минимального размера 20×20 мм для бумажного чека.

PDF является удобной копией для архива или отправки клиенту. Фискальным документом остаётся чек, сформированный кассой и ОФД.

---


## Receipt-tape PDF layout

The PDF receipt is generated as a narrow 80 mm receipt tape, not as a regular A4 report.

The PDF includes the main receipt details:

```text
cash receipt title
operation type
seller / taxpayer ID
payment address or place
date and time
shift and document number if available in Dreamkas response
tax mode
cashier
buyer email/phone
items
quantity, price, line total
VAT
grand total
payment method
FN / FD / FP fiscal fields
cash register registration/serial numbers if available
verification QR code
receipt verification URL
```

The QR code is rendered as 25×25 mm, which is above the 20×20 mm minimum used for printed receipts.

The PDF is an archive/customer-friendly copy. The fiscal document remains the receipt produced by the cash register and fiscal data operator.

---


## Автозаполнение юрлица / ИП по ИНН

Для чека на юрлицо или ИП можно заполнить только ИНН покупателя в Excel.

Если строка `Наименование юрлица / ИП` пустая, программа попробует получить название автоматически через DaData API.

Настройки:

```text
LEGAL_ENTITY_LOOKUP_ENABLED = 1
LEGAL_ENTITY_LOOKUP_PROVIDER = DADATA
DADATA_TOKEN = your_dadata_token
DADATA_BRANCH_TYPE = MAIN
DADATA_UPDATE_EXCEL_NAME = 1
```

Что происходит при чтении Excel:

```text
1. Пользователь указывает тип покупателя Юрлицо / ИП
2. Пользователь вводит ИНН
3. Наименование можно оставить пустым
4. Программа делает запрос к DaData
5. Найденное название попадает в предчек, SQLite, TXT, PDF и ФФД-теги
6. При DADATA_UPDATE_EXCEL_NAME = 1 название также записывается обратно в Excel
```

Если `DADATA_TOKEN` не указан или сервис не нашёл организацию, программа попросит заполнить название вручную.

Контур.Фокус также можно подключить как отдельный провайдер, но для него нужен коммерческий API-доступ и отдельная схема авторизации.

---


## Legal entity / sole proprietor lookup by INN

For B2B receipts, the user can fill only the buyer INN in Excel.

If the `Legal entity / sole proprietor name` row is empty, the tool can fetch the name automatically through the DaData API.

Settings:

```text
LEGAL_ENTITY_LOOKUP_ENABLED = 1
LEGAL_ENTITY_LOOKUP_PROVIDER = DADATA
DADATA_TOKEN = your_dadata_token
DADATA_BRANCH_TYPE = MAIN
DADATA_UPDATE_EXCEL_NAME = 1
```

Workflow:

```text
1. User selects legal entity / sole proprietor buyer type
2. User enters buyer INN
3. Buyer name can be left empty
4. Tool sends a request to DaData
5. Returned name is used in precheck, SQLite, TXT, PDF, and FFD tags
6. If DADATA_UPDATE_EXCEL_NAME = 1, the name is also written back to Excel
```

If `DADATA_TOKEN` is missing or the company is not found, the tool asks the user to fill the buyer name manually.

Kontur.Focus can also be added later as a separate provider, but it requires commercial API access and a separate authorization scheme.

---


## Инструкция по DaData

Для автозаполнения наименования юрлица / ИП по ИНН используйте DaData.

Кратко:

```text
1. Зарегистрируйтесь на https://dadata.ru
2. Подтвердите email
3. Откройте личный кабинет
4. Скопируйте API-токен
5. Вставьте его в settings.txt как DADATA_TOKEN
```

Настройки:

```text
LEGAL_ENTITY_LOOKUP_ENABLED = 1
LEGAL_ENTITY_LOOKUP_PROVIDER = DADATA
DADATA_TOKEN = ваш_api_токен
DADATA_BRANCH_TYPE = MAIN
DADATA_UPDATE_EXCEL_NAME = 1
```

Подробная инструкция находится в файле:

```text
DADATA_SETUP.txt
```

---


## DaData setup guide

DaData is used to auto-fill legal entity / sole proprietor name by INN.

Quick setup:

```text
1. Register at https://dadata.ru
2. Confirm your email
3. Open your DaData account
4. Copy API token
5. Add it to settings.txt as DADATA_TOKEN
```

Settings:

```text
LEGAL_ENTITY_LOOKUP_ENABLED = 1
LEGAL_ENTITY_LOOKUP_PROVIDER = DADATA
DADATA_TOKEN = your_api_token
DADATA_BRANCH_TYPE = MAIN
DADATA_UPDATE_EXCEL_NAME = 1
```

Detailed guide:

```text
DADATA_SETUP.txt
```

---


## Оплата по СБП через T-Банк

Программа поддерживает оплату по СБП через T-Банк.

Сценарий:

```text
1. Оператор выбирает тип оплаты: СБП через T-Банк
2. Программа создает платеж в T-Банке
3. Программа получает QR-код СБП
4. QR отображается прямо в терминале
5. Покупатель оплачивает через банковское приложение
6. Программа проверяет статус оплаты
7. После успешной оплаты чек отправляется в Dreamkas
```

В Dreamkas такой платеж передается как безналичный.

Настройки:

```text
TBANK_SBP_ENABLED = 1
TBANK_TERMINAL_KEY = your_terminal_key
TBANK_PASSWORD_ENC =
TBANK_PAY_TYPE = O
TBANK_PAYMENT_TIMEOUT_MINUTES = 10
TBANK_POLL_INTERVAL_SECONDS = 5
TBANK_QR_DATA_TYPE = PAYLOAD
```

Пароль терминала T-Банка вводится через меню настроек и хранится в зашифрованном виде.

Подробная инструкция:

```text
TBANK_SBP_SETUP.txt
```

---


## SBP payments through T-Bank

The tool supports SBP payments through T-Bank acquiring.

Workflow:

```text
1. Operator selects payment type: SBP through T-Bank
2. The tool creates a T-Bank payment
3. The tool receives SBP QR payload
4. QR code is printed directly in the terminal
5. Buyer pays through a banking app
6. The tool checks payment status
7. After successful payment, the receipt is sent to Dreamkas
```

Dreamkas receives this payment as cashless.

Settings:

```text
TBANK_SBP_ENABLED = 1
TBANK_TERMINAL_KEY = your_terminal_key
TBANK_PASSWORD_ENC =
TBANK_PAY_TYPE = O
TBANK_PAYMENT_TIMEOUT_MINUTES = 10
TBANK_POLL_INTERVAL_SECONDS = 5
TBANK_QR_DATA_TYPE = PAYLOAD
```

Detailed guide:

```text
TBANK_SBP_SETUP.txt
```

---


## Способ безналичной оплаты

Безналичная оплата может приниматься двумя способами:

```text
1. Внешний банковский терминал
2. T-Банк СБП
```

Выбор выполняется в меню:

```text
Настройки → Выбрать способ безналичной оплаты
```

Настройка в `settings.txt`:

```text
CASHLESS_PAYMENT_PROVIDER = EXTERNAL_TERMINAL
```

или:

```text
CASHLESS_PAYMENT_PROVIDER = TBANK_SBP
```

### Внешний банковский терминал

Если выбран внешний терминал, программа не создаёт платеж автоматически.

Перед отправкой чека в Dreamkas она покажет сумму и спросит оператора:

```text
Оплата на внешнем терминале успешно принята? Введите ДА:
```

Если оператор подтверждает оплату, чек отправляется в Dreamkas как безналичный.

Если оператор не подтверждает оплату, чек в Dreamkas не отправляется.

### T-Банк СБП

Если выбран `TBANK_SBP`, программа создаёт платеж в T-Банке, показывает QR-код СБП и ждёт успешный статус оплаты.

Только после успешной оплаты чек отправляется в Dreamkas как безналичный.

---


## Cashless payment provider

Cashless payment can be accepted in two ways:

```text
1. External bank terminal
2. T-Bank SBP
```

The provider is selected in:

```text
Settings → Select cashless payment provider
```

`settings.txt`:

```text
CASHLESS_PAYMENT_PROVIDER = EXTERNAL_TERMINAL
```

or:

```text
CASHLESS_PAYMENT_PROVIDER = TBANK_SBP
```

### External bank terminal

If an external terminal is selected, the tool does not create a payment automatically.

Before sending the receipt to Dreamkas, it shows the amount and asks the operator:

```text
Was the payment accepted on the external terminal? Type YES:
```

If the operator confirms, the receipt is sent to Dreamkas as cashless.

If the operator does not confirm, the Dreamkas receipt is not sent.

### T-Bank SBP

If `TBANK_SBP` is selected, the tool creates a T-Bank payment, displays the SBP QR code, and waits for successful payment status.

Only after successful payment the receipt is sent to Dreamkas as cashless.

---


## Важно при обновлении

Если в окне запуска отображается старая версия, например:

```text
Dreamkas Receipt Tool v6.12
```

значит запускается старый `start_dreamkas.bat` из старой папки.

Рекомендуемый порядок обновления:

```text
1. Распаковать новую версию в отдельную чистую папку
2. Скопировать туда settings.txt
3. При необходимости скопировать папку db
4. Запускать start_dreamkas.bat именно из новой папки
```

В v6.30 исправлены BAT-файлы: PowerShell-команды поиска Python теперь выполняются одной строкой без `^`.

---


## Локальный Python в папке утилиты

Если на компьютере нет нормального Python или системный Python не определяется, утилита может установить Python локально прямо в папку проекта:

```text
_python\python.exe
```

Для этого добавлен файл:

```text
install_local_python.bat
```

Он скачивает официальный установщик Python и устанавливает его в подпапку `_python`, не меняя системный `PATH`.

После этого все BAT-файлы сначала проверяют:

```text
_python\python.exe
```

и используют его, если он есть.

Это делает запуск стабильным даже на компьютерах, где:

```text
python
py
winget
PATH
```

работают некорректно или не видят установленный Python.

---


## Установка Python в LocalAppData

В v6.30 локальная установка Python перенесена из папки проекта в стабильную папку пользователя:

```text
%LOCALAPPDATA%\DreamkasReceiptTool\Python313\python.exe
```

Это сделано потому, что установщик Python может не создавать файлы в папках OneDrive, на Рабочем столе или в путях с кириллицей.

Если обычный установщик Python не создаст `python.exe`, BAT попробует резервный вариант:

```text
%LOCALAPPDATA%\DreamkasReceiptTool\Python313_embed\python.exe
```

То есть программа больше не зависит от системного `python`, `py`, `winget` или `PATH`.

---


## Жёсткий путь Python

В v6.30 установка Python упрощена.

Теперь программа использует фиксированный путь:

```text
C:\Python314\python.exe
```

Если Python не найден, BAT скачивает официальный установщик Python и устанавливает его в:

```text
C:\Python314
```

Если Windows не разрешит установку в корень диска `C:\`, запустите BAT-файл от имени администратора.

Рекомендуемый порядок:

```text
1. install_python.bat
2. install_dependencies.bat
3. start_dreamkas.bat
```

---


## Установка Python в C:\Python314

В v6.30 установщик Python всегда скачивается заново, чтобы исключить повреждённый или старый файл в `%TEMP%`.

Изменения:

```text
1. Старый установщик в TEMP удаляется
2. Python installer скачивается заново
3. Установка идёт в C:\Python314
4. Используется InstallAllUsers=1
5. При ошибке показываются последние строки dreamkas-python-install.log
```

Для установки в `C:\Python314` лучше запускать:

```text
install_python.bat
```

от имени администратора.

---


## Ремонт установки Python в C:\Python314

В v6.30 исправлен случай, когда `C:\Python314\python.exe` создаётся, но Python не находит стандартную библиотеку `encodings`.

Причина обычно в неполной установке: в логе видно `Include_exe = 0` или отсутствует стандартная библиотека.

Что делает v6.30:

```text
1. Проверяет C:\Python314\python.exe через import encodings
2. Если Python битый — удаляет C:\Python314
3. Скачивает установщик Python заново
4. Запускает установку с явными параметрами:
   Include_exe=1
   Include_lib=1
   Include_pip=1
5. Если обычный установщик снова не дал рабочий Python,
   распаковывает embedded Python ZIP в C:\Python314
```

Если удаление `C:\Python314` или запись в корень диска не проходит, запустите `install_python.bat` от имени администратора.

---


## Обязательный контакт покупателя

В v6.30 удалён вариант:

```text
0. Не указывать
```

Теперь при создании электронного чека оператор должен выбрать один из вариантов:

```text
1. Email
2. Телефон
3. Email + телефон
```

Если выбран email — нужно ввести email.  
Если выбран телефон — нужно ввести телефон.  
Если выбран `Email + телефон` — нужно заполнить оба поля.

---


## Выбор системы налогообложения

В v6.30 система налогообложения выбирается из меню на русском языке.

Программа показывает список:

```text
0. По умолчанию на кассе / в Dreamkas
1. Общая система налогообложения — ОСН
2. Упрощённая система — доходы — УСН доходы
3. Упрощённая система — доходы минус расходы — УСН доходы-расходы
4. Единый сельскохозяйственный налог — ЕСХН
5. Патентная система налогообложения — ПСН / патент
6. Единый налог на вменённый доход — ЕНВД, архивный режим
```

Оператор вводит только цифру.

В Dreamkas передаётся соответствующий внутренний код:

```text
0 → DEFAULT
1 → OSN
2 → SIMPLE
3 → SIMPLE_WO
4 → AGRICULT
5 → PATENT
6 → ENVD
```

При нажатии Enter используется последнее выбранное значение из `settings.txt`.

---

## Статус ожидания оплаты СБП

Во время ожидания оплаты по СБП строка статуса теперь не печатается каждый раз новой строкой.

Вместо этого программа перезаписывает последнюю строку:

```text
Статус оплаты: FORM_SHOWED. Следующая проверка через 5 сек.
```

Это сделано, чтобы QR-код СБП в терминале не сдвигался вниз при каждой проверке статуса.

---


## Выпадающие списки в Excel

В Excel-шаблон возвращены выпадающие списки:

```text
Колонка B — Тип позиции:
1 = услуга
0 = товар

Колонка E — Ставка НДС:
0
5
7
10
22
5/105
7/107
10/110
22/122
```

Эти списки есть как в готовом `dreamkas_receipt_template.xlsx`, так и в шаблоне, который программа создаёт автоматически, если файл Excel отсутствует.

---

## Возврат денег по СБП

Если исходный чек был оплачен через T-Банк СБП и в SQLite сохранён `sbp_payment_id`, то после успешной фискализации чека возврата программа предложит вернуть деньги через T-Банк.

Сценарий:

```text
1. Оператор делает возврат чека в Dreamkas
2. Чек возврата успешно фискализирован
3. Программа видит, что исходный чек был оплачен через СБП
4. Программа предлагает отправить возврат денег через Т-Банк
5. При подтверждении отправляется запрос /v2/Cancel
```

Для частичного возврата в T-Банк передаётся сумма возврата в копейках.

---


## СБП для юрлиц / ИП: обычная СБП или B2B/I2I

Если покупатель выбран как `Юрлицо / ИП` и способ оплаты — `T-Банк СБП`, программа теперь спрашивает формат оплаты:

```text
1. Обычная СБП-оплата QR через интернет-эквайринг T-Банка
2. СБП B2B/I2I — ссылка/QR для оплаты юрлицом или ИП через T-API
```

Обычная СБП работает через EACQ:

```text
/v2/Init
/v2/GetQr
/v2/GetState
```

B2B/I2I работает через T-API:

```text
POST https://business.tbank.ru/openapi/api/v1/b2b/qr/onetime
GET  https://business.tbank.ru/openapi/api/v1/b2b/qr/{qrId}/info
```

Для B2B/I2I нужны отдельные настройки:

```text
TBANK_B2B_SBP_ENABLED = 1
TBANK_B2B_API_TOKEN_ENC = ...
TBANK_B2B_ACCOUNT_NUMBER = расчетный счет 20 или 22 цифры
TBANK_B2B_TTL_DAYS = 1
TBANK_B2B_VAT = 0
TBANK_B2B_REDIRECT_URL =
TBANK_B2B_PURPOSE_PREFIX = Оплата по счету
```

Важно: `TBANK_B2B_API_TOKEN` — это не `TerminalKey` и не пароль интернет-эквайринга. Это отдельный Bearer token T-API с правами на выставление B2B-ссылок/QR через СБП.

---


## Отмена ожидания при фискализации и оплате

В v6.30 во время длительных операций можно нажать:

```text
ESC
```

Это работает во время:

```text
- ожидания оплаты СБП;
- ожидания оплаты СБП B2B/I2I;
- ожидания ответа Dreamkas по заданию фискализации.
```

После нажатия `ESC` программа возвращается в главное меню.

Важно: если задание уже было отправлено в Dreamkas, локальная отмена не гарантирует отмену операции на кассе. Касса может продолжить выполнение задания. Поэтому запись помечается в SQLite как:

```text
CANCELLED_BY_OPERATOR
```

Перед повторной пробивкой такого же чека нужно проверить статус операции, чтобы не создать дубль.

---
