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
