# Быстрый старт: сканирование и архивирование

## 0) Настройка нового ноутбука (Pantum BM2300W)

1. Подключить МФУ и установить официальный драйвер Pantum BM2300W (Printer + Scanner).
2. Проверить, что служба WIA (`stisvc`) в состоянии Running.
3. Установить OCR-движок Tesseract (рекомендуется).
4. Выполнить проверку:

```bash
python scripts/scanner_control.py diagnose
python scripts/scanner_control.py list
```

Автоматизированный helper:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_pantum_bm2300w.ps1
```

## 1) Куда класть свежие сканы

Папка входящих сканов объекта:

- `docflow/objects/<object>/10_scan_inbox/`

Если хотите сканировать прямо с МФУ из командной строки, используйте:

```bash
python scripts/scanner_control.py list
python scripts/scanner_control.py scan-to-inbox \
  --object-root docflow/objects/x5-ufa-e2_logistics_park \
  --doc-type AWR \
  --subject foundation_axis_1_7 \
  --device-index 1 \
  --format jpg \
  --dpi 300
```

## 2) Как называть файлы

Формат имени:

`YYYYMMDD__DOC_TYPE__SUBJECT__[EMPLOYEE_ID].ext`

DOC_TYPE:

- `AWR` - акт скрытых работ
- `PASSPORT` - паспорт сотрудника (обязательно указывать EMPLOYEE_ID)
- `ORDER` - приказ

Примеры:

- `20260310__AWR__foundation_axis_1_7.pdf`
- `20260310__PASSPORT__ivanov_ii__001.pdf`
- `20260310__ORDER__appoint_hse_responsible.pdf`

## 3) Запуск автоматической раскладки

```bash
python scripts/scan_archive_docs.py ingest --object-root docflow/objects/x5-ufa-e2_logistics_park
```

OCR включен по умолчанию. Дополнительные параметры:

```bash
python scripts/scan_archive_docs.py ingest \
  --object-root docflow/objects/x5-ufa-e2_logistics_park \
  --ocr-lang rus+eng \
  --max-pdf-pages 4
```

Что делает команда:

- переносит акты скрытых работ в `05_execution_docs/hidden_work_acts/`
- переносит паспорта в досье сотрудника `02_personnel/employees/<id>_*/01_identity_and_contract/`
- переносит приказы в `01_orders_and_appointments/`
- нерспознанные или ошибочные файлы переносит в `10_scan_inbox/manual_review/`
- пишет причины в `10_scan_inbox/manual_review/review_log.csv`
- при успешном OCR создает рядом sidecar-файл `*.ocr.txt`

## 4) Сбор архивного пакета по периоду

```bash
python scripts/scan_archive_docs.py bundle \
  --object-root docflow/objects/x5-ufa-e2_logistics_park \
  --from-date 2026-03-01 \
  --to-date 2026-03-31
```

Результат: ZIP в `09_archive/scan_bundles/`.
