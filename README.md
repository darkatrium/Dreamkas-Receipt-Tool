# Dreamkas Receipt Tool

**Dreamkas Receipt Tool** — консольная утилита для Windows, которая помогает формировать кассовые чеки по Excel-шаблону, отправлять задания на фискализацию в Dreamkas, сохранять результаты локально и работать с возвратами.

Актуальная версия: **v6.30**

---

## Возможности

- Фискализация чеков через Dreamkas API.
- Возврат полного чека или отдельных позиций.
- Excel-шаблон только для товарных позиций.
- Ввод реквизитов чека через терминал.
- Работа с физлицами, юрлицами и ИП.
- Автозаполнение наименования юрлица / ИП по ИНН через DaData.
- Сохранение истории чеков в SQLite.
- Проверка похожих чеков перед повторной фискализацией.
- Генерация TXT-чека.
- Генерация PDF-чека в виде кассовой ленты.
- Генерация QR-кода проверки чека.
- Очистка Excel-шаблона после успешной фискализации.
- Отправка копий чеков по SMTP.
- Поддержка наличной оплаты.
- Поддержка безналичной оплаты через внешний терминал.
- Поддержка оплаты через T-Банк СБП.
- Поддержка СБП B2B/I2I для юрлиц и ИП через T-API.
- Возврат денег через T-Банк при возврате чека, если исходная оплата была через СБП.
- Возможность отменить ожидание оплаты / фискализации клавишей `ESC`.
- Автоматическая установка Python в `C:\Python314`.
- BAT-файлы для запуска, диагностики, установки зависимостей и сборки EXE.
- Документация на русском и английском языках.

---

## Общая логика работы

```text
1. Оператор запускает start_dreamkas.bat
2. Программа проверяет Python и зависимости
3. Оператор выбирает магазин и кассу Dreamkas
4. Оператор вводит имя кассира
5. В главном меню выбирает:
   - фискализация нового чека
   - возврат чека
   - настройки
   - выход
6. При фискализации нового чека:
   - реквизиты вводятся в терминале
   - позиции вводятся в Excel
   - программа показывает предчек
   - оператор подтверждает отправку
   - при необходимости выполняется оплата
   - чек отправляется в Dreamkas
   - результат сохраняется локально
```

---

## Требования

- Windows 10 / Windows 11.
- Интернет-доступ.
- Аккаунт Dreamkas.
- Токен Dreamkas API.
- ID магазина и кассы Dreamkas.
- Excel или совместимый табличный редактор.
- Python 3.10+.

Если Python отсутствует, утилита может установить его автоматически в:

```text
C:\Python314
```

Для установки в `C:\Python314` рекомендуется запускать `install_python.bat` от имени администратора.

---

## Быстрый старт

### 1. Распаковать архив

Распакуйте проект в отдельную папку, например:

```text
C:\DreamkasReceiptTool
```

Не рекомендуется запускать из ZIP-архива.

---

### 2. Установить Python

Запустите:

```text
install_python.bat
```

Если установка в `C:\Python314` не проходит, запустите BAT-файл от имени администратора.

---

### 3. Установить зависимости

Запустите:

```text
install_dependencies.bat
```

---

### 4. Запустить программу

Запустите:

```text
start_dreamkas.bat
```

---

## Файлы проекта

```text
dreamkas_receipt.py              основная консольная программа
dreamkas_gui.py                  экспериментальный GUI
dreamkas_receipt_template.xlsx   Excel-шаблон товарных позиций
settings.example.txt             пример файла настроек
requirements.txt                 Python-зависимости

start_dreamkas.bat               запуск консольной версии
start_gui.bat                    запуск GUI
install_python.bat               установка/проверка Python
install_local_python.bat         установка Python в C:\Python314
install_dependencies.bat         установка зависимостей
diagnose_python.bat              диагностика Python
build_exe.bat                    сборка EXE через PyInstaller
test_smtp.bat                    проверка SMTP
test_smtp.py                     SMTP-тест

README.md                        документация
README.txt                       текстовая версия документации
CHANGELOG.md                     история изменений
DADATA_SETUP.txt                 инструкция по DaData
SMTP_SETUP.txt                   инструкция по SMTP
TBANK_SBP_SETUP.txt              инструкция по T-Банк СБП
GITHUB_COMMIT_AND_RELEASE_TEXT.md подсказки для GitHub-коммитов и релизов
```

---

## Папки, создаваемые программой

```text
db/                SQLite-база чеков
logs/              логи и ошибки API
receipts_txt/      TXT-копии чеков
receipts_pdf/      PDF-копии чеков
receipts_qr/       QR-коды чеков и QR-коды СБП
```

---

## Что не нужно публиковать в GitHub

Файл `settings.txt` содержит реальные токены и настройки, поэтому его нельзя выкладывать в репозиторий.

Рекомендуемый `.gitignore`:

```gitignore
settings.txt
python_path.txt
db/
logs/
receipts_txt/
receipts_pdf/
receipts_qr/
__pycache__/
*.pyc
dist/
build/
*.spec
```

---

## Настройки

Настройки хранятся в файле:

```text
settings.txt
```

Формат:

```text
TOKEN_KEY = VALUE
```

Пример:

```text
DREAMKAS_TOKEN = your_dreamkas_token
STORE_ID =
DEVICE_ID =
DEVICE_NAME =
DEFAULT_CASHIER_NAME =
DEFAULT_PAYMENT_TYPE = CASHLESS
CASHLESS_PAYMENT_PROVIDER = EXTERNAL_TERMINAL
DEFAULT_TAX_MODE = SIMPLE_WO
```

При первом запуске программа запросит токен Dreamkas, затем получит список магазинов и касс.

---

## Dreamkas

Программа использует Dreamkas API для:

```text
- получения списка магазинов;
- получения списка касс;
- отправки задания на фискализацию;
- проверки статуса операции;
- получения результата фискализации.
```

При запуске программа каждый раз предлагает выбрать магазин и кассу. Выбранные значения сохраняются в `settings.txt`.

---

## Excel-шаблон

Начиная с актуальных версий, Excel используется только для товарных позиций.

Все реквизиты чека вводятся в терминале.

### Формат Excel

```text
1 строка: шапка таблицы
2 строка и ниже: позиции чека
```

Колонки:

```text
A: Наименование
B: Тип — Услуга = 1 / Товар = 0
C: Количество
D: Цена
E: Ставка НДС
```

---

## Выпадающие списки в Excel

В Excel-шаблоне есть выпадающие списки.

### Колонка B — тип позиции

```text
1 = услуга
0 = товар
```

### Колонка E — ставка НДС

```text
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

Если Excel-шаблон отсутствует, программа создаст новый шаблон автоматически и добавит в него эти же выпадающие списки.

---

## Ввод реквизитов чека

Перед открытием Excel программа спрашивает в терминале:

```text
- тип оплаты;
- контакт покупателя;
- тип покупателя;
- ИНН и название юрлица / ИП;
- систему налогообложения;
- имя кассира.
```

Имя кассира спрашивается один раз при запуске программы и действует до закрытия программы.

---

## Контакт покупателя

Для электронного чека контакт покупателя обязателен.

Меню:

```text
Контакт покупателя для электронного чека:
  1. Email
  2. Телефон
  3. Email + телефон
Выберите вариант [1/2/3]:
```

Если выбран email — нужно указать email.

Если выбран телефон — нужно указать телефон.

Если выбран `Email + телефон` — нужно указать оба значения.

Телефон можно вводить так:

```text
+79991234567
89991234567
79991234567
```

Российский номер через `8` автоматически приводится к формату `+7`.

---

## Тип покупателя

Поддерживаются:

```text
1. Физлицо
2. Юрлицо / ИП
```

Если выбран покупатель `Юрлицо / ИП`, программа запросит ИНН.

Название юрлица можно:

```text
- ввести вручную;
- получить автоматически через DaData.
```

---

## Автозаполнение юрлица / ИП по ИНН через DaData

Для автозаполнения используется DaData.

Настройки:

```text
LEGAL_ENTITY_LOOKUP_ENABLED = 1
LEGAL_ENTITY_LOOKUP_PROVIDER = DADATA
DADATA_TOKEN = your_dadata_token
DADATA_BRANCH_TYPE = MAIN
DADATA_UPDATE_EXCEL_NAME = 1
SEND_LEGAL_ENTITY_TAGS_TO_DREAMKAS = 1
```

Если покупатель — юрлицо / ИП и введён только ИНН, программа делает запрос в DaData и подставляет найденное название в:

```text
- предчек;
- SQLite;
- TXT-чек;
- PDF-чек;
- ФФД-теги Dreamkas.
```

Подробная инструкция находится в:

```text
DADATA_SETUP.txt
```

---

## Система налогообложения

Система налогообложения выбирается из меню на русском языке.

```text
0. По умолчанию на кассе / в Dreamkas
1. Общая система налогообложения — ОСН
2. Упрощённая система — доходы — УСН доходы
3. Упрощённая система — доходы минус расходы — УСН доходы-расходы
4. Единый сельскохозяйственный налог — ЕСХН
5. Патентная система налогообложения — ПСН / патент
6. Единый налог на вменённый доход — ЕНВД, архивный режим
```

Соответствие внутренним кодам:

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

## Типы оплаты

Поддерживаются:

```text
- наличные;
- безнал через внешний терминал;
- безнал через T-Банк СБП;
- СБП B2B/I2I для юрлиц и ИП.
```

---

## Настройка способа безналичной оплаты

Настройка:

```text
CASHLESS_PAYMENT_PROVIDER = EXTERNAL_TERMINAL
```

или:

```text
CASHLESS_PAYMENT_PROVIDER = TBANK_SBP
```

Выбор выполняется в меню:

```text
Настройки → Выбрать способ безналичной оплаты
```

---

## Внешний банковский терминал

Если выбран внешний терминал, программа не создаёт платеж автоматически.

Перед отправкой чека в Dreamkas она показывает сумму и спрашивает:

```text
Оплата на внешнем терминале успешно принята? Введите ДА:
```

Если оператор вводит `ДА`, чек отправляется в Dreamkas как безналичный.

Если оператор не подтверждает оплату, чек в Dreamkas не отправляется.

---

## T-Банк СБП

Если выбран T-Банк СБП, программа:

```text
1. Создаёт платеж в T-Банке.
2. Получает QR-код СБП.
3. Показывает QR-код в терминале.
4. Сохраняет QR-код в receipts_qr/.
5. Проверяет статус платежа.
6. После успешной оплаты отправляет чек в Dreamkas.
```

Используются методы интернет-эквайринга T-Банка:

```text
/v2/Init
/v2/GetQr
/v2/GetState
```

Настройки:

```text
TBANK_SBP_ENABLED = 1
TBANK_TERMINAL_KEY =
TBANK_PASSWORD_ENC =
TBANK_PAY_TYPE = O
TBANK_PAYMENT_TIMEOUT_MINUTES = 10
TBANK_POLL_INTERVAL_SECONDS = 5
TBANK_REQUEST_TIMEOUT_SECONDS = 30
TBANK_QR_DATA_TYPE = PAYLOAD
TBANK_DESCRIPTION_PREFIX = Оплата заказа
```

Пароль терминала хранится в `settings.txt` в зашифрованном виде.

---

## СБП для юрлиц / ИП: обычная СБП или B2B/I2I

Если покупатель выбран как `Юрлицо / ИП` и способ оплаты — `T-Банк СБП`, программа спрашивает формат оплаты:

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

Настройки B2B/I2I:

```text
TBANK_B2B_SBP_ENABLED = 1
TBANK_B2B_API_TOKEN_ENC =
TBANK_B2B_ACCOUNT_NUMBER =
TBANK_B2B_TTL_DAYS = 1
TBANK_B2B_VAT = 0
TBANK_B2B_REDIRECT_URL =
TBANK_B2B_PURPOSE_PREFIX = Оплата по счету
```

Важно: `TBANK_B2B_API_TOKEN` — это не `TerminalKey` и не пароль интернет-эквайринга. Это отдельный Bearer token T-API.

---

## Ожидание оплаты СБП

Во время ожидания СБП строка статуса обновляется в одной строке:

```text
Статус оплаты: FORM_SHOWED. Следующая проверка через 5 сек.
```

Это сделано, чтобы QR-код в терминале не сдвигался вниз.

---

## Отмена ожидания

Во время длительных операций можно нажать:

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

Важно: если задание уже было отправлено в Dreamkas, локальная отмена не гарантирует отмену операции на кассе. Касса может продолжить выполнение задания.

Такая запись помечается в SQLite как:

```text
CANCELLED_BY_OPERATOR
```

Перед повторной пробивкой такого же чека нужно проверить статус операции, чтобы не создать дубль.

---

## Предчек

Перед отправкой программа показывает предчек:

```text
- тип чека;
- тип оплаты;
- покупатель;
- юрлицо / ИП;
- СНО;
- кассир;
- список позиций;
- количество;
- цена;
- сумма;
- НДС;
- итог.
```

Оператор должен подтвердить отправку.

---

## Проверка похожих чеков

Перед фискализацией программа сохраняет данные предчека и проверяет, не был ли похожий чек уже пробит.

Если найден похожий чек, программа предупреждает оператора.

Это помогает избежать случайной повторной пробивки одного и того же чека.

---

## Фискализация

После подтверждения программа отправляет задание в Dreamkas.

Затем она опрашивает статус операции до результата:

```text
SUCCESS
ERROR
```

При успехе сохраняются:

```text
- SQLite-запись;
- TXT-чек;
- PDF-чек;
- QR-код;
- JSON-ответ Dreamkas;
- пути к созданным файлам.
```

---

## PDF-чек

PDF-чек формируется в виде узкой кассовой ленты.

В PDF выводятся:

```text
- тип чека;
- продавец;
- кассир;
- покупатель;
- СНО;
- позиции;
- количество;
- цена;
- НДС;
- итог;
- форма оплаты;
- ФН / ФД / ФП, если есть в ответе;
- ссылка проверки;
- QR-код.
```

PDF-файлы сохраняются в:

```text
receipts_pdf/
```

---

## TXT-чек

TXT-файлы сохраняются в:

```text
receipts_txt/
```

---

## QR-коды

QR-коды сохраняются в:

```text
receipts_qr/
```

Туда попадают:

```text
- QR-коды проверки фискальных чеков;
- QR-коды оплаты СБП;
- QR-коды / ссылки B2B/I2I.
```

---

## Возврат чека

В режиме возврата программа показывает список чеков из SQLite.

Оператор выбирает чек и вид возврата:

```text
- полный возврат;
- частичный возврат;
- отмена.
```

Если в исходном чеке одна позиция, доступен только полный возврат или отмена.

Для частичного возврата оператор выбирает позиции через запятую:

```text
1,5
```

После выбора программа показывает чек с отмеченными позициями и просит подтвердить возврат.

---

## Возврат денег через T-Банк СБП

Если исходный чек был оплачен через T-Банк СБП и в SQLite сохранён `sbp_payment_id`, то после успешной фискализации чека возврата программа предложит вернуть деньги через T-Банк.

Сценарий:

```text
1. Оператор делает возврат чека в Dreamkas.
2. Чек возврата успешно фискализирован.
3. Программа видит, что исходный чек был оплачен через СБП.
4. Программа предлагает отправить возврат денег через Т-Банк.
5. При подтверждении отправляется запрос /v2/Cancel.
```

Для частичного возврата в T-Банк передаётся сумма возврата в копейках.

Если оператор пропускает возврат денег, это сохраняется в SQLite.

Если возврат денег через T-Банк завершился ошибкой, фискальный чек возврата остаётся созданным, а ошибка сохраняется в SQLite.

---

## SQLite

История чеков хранится в:

```text
db/receipts.sqlite3
```

В базе сохраняются:

```text
- внешний ID операции;
- ID операции Dreamkas;
- ID фискального чека;
- тип чека;
- статус;
- дата создания;
- дата завершения;
- кассир;
- покупатель;
- юрлицо / ИП;
- ИНН;
- СНО;
- позиции;
- суммы;
- ответы API;
- пути к TXT/PDF/QR;
- данные СБП;
- данные возвратов;
- ошибки.
```

---

## SMTP

Программа умеет отправлять копии чеков на email.

Настройки:

```text
SMTP_ENABLED = 0
SMTP_HOST =
SMTP_PORT = 465
SMTP_SSL = 1
SMTP_STARTTLS = 0
SMTP_USERNAME =
SMTP_PASSWORD_ENC =
SMTP_FROM =
SMTP_SEND_TO_BUYER = 1
SMTP_SEND_TO_OWNER = 0
SMTP_OWNER_EMAIL =
SMTP_VERIFY_CERT = 1
```

SMTP-пароль хранится в зашифрованном виде.

Подробная инструкция находится в:

```text
SMTP_SETUP.txt
```

---

## Шифрование паролей

В `settings.txt` могут храниться зашифрованные значения:

```text
SMTP_PASSWORD_ENC
TBANK_PASSWORD_ENC
TBANK_B2B_API_TOKEN_ENC
```

Ключом шифрования используется Dreamkas token.

Если Dreamkas token меняется, зашифрованные SMTP/T-Банк значения очищаются и их нужно ввести заново.

---

## Установка Python

Актуальная логика установки Python:

```text
1. Программа ищет C:\Python314\python.exe.
2. Если Python не найден или битый, запускается установка.
3. Python ставится в C:\Python314.
4. Проверяется import encodings.
5. Если обычный установщик не дал рабочий Python,
   используется embedded Python ZIP.
```

Для установки в `C:\Python314` может потребоваться запуск от имени администратора.

---

## Диагностика

Для диагностики Python:

```text
diagnose_python.bat
```

Для проверки SMTP:

```text
test_smtp.bat
```

Логи API и ошибок находятся в:

```text
logs/
```

---

## Сборка EXE

Для сборки EXE:

```text
build_exe.bat
```

Результат будет в папке:

```text
dist/
```

---

## Безопасность

Не публикуйте в GitHub:

```text
settings.txt
db/
logs/
receipts_txt/
receipts_pdf/
receipts_qr/
```

Файл `settings.example.txt` можно публиковать.

---

## Важное юридическое замечание

ЧЕРЕЗ ЭТУ УТИЛИТУ НЕЛЬЗЯ ФИСКАЛИЗИРОВАТЬ ПРОСЛЕЖИВАЕМЫЕ ТОВАРЫ, утилита не заменяет кассу, ОФД или фискальный накопитель. Фискальным документом является чек, сформированный кассой и зарегистрированный через Dreamkas/ОФД.

PDF/TXT/QR-файлы, которые создаёт программа, являются локальными копиями для архива, проверки и отправки покупателю.

---

# English version

## Dreamkas Receipt Tool

**Dreamkas Receipt Tool** is a Windows console utility for preparing receipt line items in Excel, sending fiscalization tasks to Dreamkas, storing results locally, and handling refunds.

Current version: **v6.30**

---

## Features

- Dreamkas API receipt fiscalization.
- Full and partial receipt refunds.
- Excel template for receipt line items only.
- Receipt metadata entered through terminal.
- Individual, legal entity, and sole proprietor buyers.
- Legal entity / sole proprietor lookup by INN through DaData.
- Local SQLite receipt history.
- Duplicate-like receipt warning.
- TXT receipt copy.
- Receipt-style PDF copy.
- QR code generation.
- Excel cleanup after successful fiscalization.
- SMTP receipt sending.
- Cash payments.
- Cashless payments through external terminal.
- T-Bank SBP payments.
- T-Bank SBP B2B/I2I for legal entities and sole proprietors.
- T-Bank money refund after fiscal refund when original payment was made through SBP.
- ESC cancellation during payment/fiscalization waiting.
- Python auto-installation into `C:\Python314`.
- BAT launchers for startup, diagnostics, dependencies, and EXE build.

---

## Quick start

```text
1. Extract the archive into a separate folder.
2. Run install_python.bat.
3. Run install_dependencies.bat.
4. Run start_dreamkas.bat.
```

If installation into `C:\Python314` fails, run `install_python.bat` as Administrator.

---

## Excel template

Excel contains only line items.

Columns:

```text
A: Name
B: Type — Service = 1 / Product = 0
C: Quantity
D: Price
E: VAT rate
```

Dropdowns are available for item type and VAT.

---

## Buyer contact

Buyer contact is required for electronic receipts.

Options:

```text
1. Email
2. Phone
3. Email + phone
```

---

## Tax mode

The operator chooses tax mode from a numbered Russian-language menu.

Internal mapping:

```text
0 → DEFAULT
1 → OSN
2 → SIMPLE
3 → SIMPLE_WO
4 → AGRICULT
5 → PATENT
6 → ENVD
```

---

## Cashless payment providers

```text
CASHLESS_PAYMENT_PROVIDER = EXTERNAL_TERMINAL
```

or:

```text
CASHLESS_PAYMENT_PROVIDER = TBANK_SBP
```

External terminal requires manual operator confirmation before fiscalization.

T-Bank SBP creates a payment, displays a QR code, waits for successful status, and only then sends the receipt to Dreamkas.

---

## T-Bank SBP B2B/I2I

For legal entity / sole proprietor buyers, if T-Bank SBP is selected, the program asks:

```text
1. Regular SBP QR through T-Bank acquiring
2. B2B/I2I SBP link/QR through T-API
```

B2B/I2I settings:

```text
TBANK_B2B_SBP_ENABLED = 1
TBANK_B2B_API_TOKEN_ENC =
TBANK_B2B_ACCOUNT_NUMBER =
TBANK_B2B_TTL_DAYS = 1
TBANK_B2B_VAT = 0
TBANK_B2B_REDIRECT_URL =
TBANK_B2B_PURPOSE_PREFIX = Payment
```

---

## Refunds

The tool supports full and partial fiscal refunds.

If the original sale was paid through T-Bank SBP, the tool offers to send money refund through T-Bank after successful fiscal refund.

---

## ESC cancellation

During long waits, the operator can press:

```text
ESC
```

This works during:

```text
- SBP payment polling;
- B2B/I2I payment polling;
- Dreamkas fiscalization polling.
```

If the Dreamkas task was already sent, ESC only cancels local waiting. The cash register may still complete the task.

The local receipt record is marked as:

```text
CANCELLED_BY_OPERATOR
```

---

## Security

Do not commit:

```text
settings.txt
db/
logs/
receipts_txt/
receipts_pdf/
receipts_qr/
```

---

## License

Add your project license here.
