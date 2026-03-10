# Инструкция: сканирование и сохранение (Pantum BM2300W)

## 1. Разовая настройка ноутбука

1. Подключите МФУ Pantum BM2300W по USB или сети.
2. Установите официальный драйвер Pantum (печать + сканер).
3. Запустите проверку:

   powershell -ExecutionPolicy Bypass -File scripts/setup_pantum_bm2300w.ps1

4. Проверьте, что устройство видно в проекте:

   .\.venv\Scripts\python.exe scripts\scanner_control.py diagnose
   .\.venv\Scripts\python.exe scripts\scanner_control.py list

Ожидаемо: в списке есть Pantum BM2300W series.

## 2. Куда сохранять сканы

Рабочая входящая папка этого объекта:

- docflow/objects/x5-ufa-e2_logistics_park/10_scan_inbox/

Если сканируете через утилиту Pantum, установите эту папку как Scan to Folder destination.

## 3. Рекомендуемые параметры сканирования

- Разрешение: 300 dpi
- Цветность: grayscale (для текстов), color (если важны печати/цветные отметки)
- Формат:
  - JPG/PNG для одиночных листов
  - PDF для многостраничных актов
- Ориентация: авто или портрет

## 4. Обязательный формат имени файла

YYYYMMDD__DOC_TYPE__SUBJECT__[EMPLOYEE_ID].ext

DOC_TYPE:

- AWR: акт скрытых работ
- ORDER: приказ
- PASSPORT: паспорт сотрудника (EMPLOYEE_ID обязателен)

Примеры:

- 20260310__AWR__armirovanie_fundamenta_os_1_7.pdf
- 20260310__ORDER__appoint_hse_responsible.jpg
- 20260310__PASSPORT__melnikov_av__002.jpg

## 5. Автоматическая раскладка и OCR

Запуск обработки входящих сканов:

.\.venv\Scripts\python.exe scripts\scan_archive_docs.py ingest --object-root docflow/objects/x5-ufa-e2_logistics_park --ocr-lang rus+eng

Что произойдет:

- AWR -> 05_execution_docs/hidden_work_acts/
- ORDER -> 01_orders_and_appointments/
- PASSPORT -> 02_personnel/employees/<id>_*/01_identity_and_contract/
- OCR-текст сохранится рядом как *.ocr.txt (если движок OCR доступен)
- Нераспознанные файлы уйдут в 10_scan_inbox/manual_review/

## 6. Как вносить удостоверения в приказы

1. Скан удостоверения кладете в inbox по формату имени (обычно тип ORDER или PASSPORT).
2. Запускаете ingest.
3. Открываете нужный приказ в:
   - docflow/objects/x5-ufa-e2_logistics_park/01_orders_and_appointments/
4. В таблице приложения заполняете:
   - номер удостоверения
   - дату выдачи
   - срок действия
   - путь к скану

## 7. Архивный пакет на сдачу

Сборка ZIP за период:

.\.venv\Scripts\python.exe scripts\scan_archive_docs.py bundle --object-root docflow/objects/x5-ufa-e2_logistics_park --from-date 2026-03-01 --to-date 2026-03-31

Результат:

- docflow/objects/x5-ufa-e2_logistics_park/09_archive/scan_bundles/
