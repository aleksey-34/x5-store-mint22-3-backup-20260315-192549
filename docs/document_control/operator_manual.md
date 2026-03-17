# Пояснительная инструкция для ведения документооборота

Документ для роли: зам. прораба / ответственный за документооборот на объекте.

## 1. Роль системы

X5_Storage в текущем виде - это локальное АРМ (автоматизированное рабочее место) документооборота стройплощадки.

Система состоит из:

- API-приложения для учета записей
- локальной базы данных
- стандартной структуры папок
- пакета шаблонов
- обработки сообщений Telegram по правилам

Сейчас это рабочий контур из текущего репозитория. Позже его можно развернуть как продукт на сервере с интерфейсом.

## 2. Что вы получаете на вход

1. Реквизиты организации.
2. Список сотрудников, должности, допуски, удостоверения.
3. Ежедневные сообщения прорабов о выполнении работ.

## 3. Подготовка перед стартом объекта

1. Заполнить карточку реквизитов:

   [docs/templates/organization/company_requisites_card.md](../templates/organization/company_requisites_card.md)

2. Заполнить список сотрудников:

   [docs/templates/personnel/employees_sample.csv](../templates/personnel/employees_sample.csv)

3. Сгенерировать дерево папок и досье сотрудников:

   ```bash
   python scripts/bootstrap_object_docs.py \
     --object-code X5-UFA-E2 \
     --object-name logistics_park \
     --employees-csv docs/templates/personnel/employees_sample.csv
   ```

4. Проверить структуру по стандарту:

   [docs/document_control/folder_structure.md](folder_structure.md)

## 4. Какие документы вести

Перечень оставлен только по пунктам, отмеченным знаком "+".

База для контроля:

- [docs/document_control/standard_registers.md](standard_registers.md)
- [docs/templates/checklists/startup_checklist_form2.md](../templates/checklists/startup_checklist_form2.md)
- [docs/templates/journals/journal_registry_matrix.md](../templates/journals/journal_registry_matrix.md)

## 5. Ежедневный рабочий цикл

### Утро

1. Проверить незакрытые замечания и сроки удостоверений.
2. Обновить графики/статусы работ.
3. Проверить готовность нарядов-допусков и журналов.

### День

1. Получать сообщения прорабов в Telegram.
2. Пропускать сообщения через обработку в API.
3. Фиксировать важные события в журнале.

### Вечер

1. Сверить факт выполнения за день.
2. Обновить реестр исполнительной документации.
3. Подготовить документы на печать/подпись/сдачу.

## 6. Работа с Telegram

Сценарий через Swagger UI:

1. Открыть http://127.0.0.1:8000/docs.
2. В разделе `/telegram/rules/` создать правила ключевых слов.
3. В разделе `/telegram/messages/process` передать текст сообщения.
4. В разделе `/journal/` проверить, какие записи созданы.

Принцип:

- если ключевое слово найдено, выполняется действие по правилу
- результат и следы обработки фиксируются в журнале

## 7. Как печатать и сохранять документы

## 6.1 Локальный AI-помощник (Ollama)

Система поддерживает локальную модель для ежедневной подготовки текстов без отправки данных во внешний облачный сервис.

Доступные API-маршруты:

- `GET /local-llm/status` - проверить доступность локального runtime и модели
- `POST /local-llm/chat` - отправить запрос в локальную модель

Рекомендуемый сценарий:

1. Перед стартом смены проверить `status`.
2. Передавать в `chat` задачу и контекст (объект, раздел, сроки).
3. Полученный текст использовать как черновик для приказов, чек-листов и служебных записей.
4. Финальную редакцию и ответственность за содержание оставлять за ответственным ИТР.

Пример `POST /local-llm/chat`:

```json
{
   "prompt": "Сформируй 8 пунктов контроля на сегодня",
   "context": "Объект Уфа Х5; приоритет: журналы и наряды",
   "temperature": 0.2,
   "num_predict": 220
}
```

## 7. Как печатать и сохранять документы

Рекомендуемая схема:

1. Готовите документ в редактируемом формате (DOCX/XLSX/MD).
2. Проверяете комплектность и реквизиты.
3. Печатаете экземпляр для подписи.
4. Подписанный документ сканируете в PDF (300 dpi).
5. Сохраняете:
   - редактируемую версию
   - подписанный PDF
6. Вносите запись в реестр документов с путем к файлу.

Шаблон реестра исполнительной документации:

[docs/templates/execution/executive_docs_register.csv](../templates/execution/executive_docs_register.csv)

Рекомендуемое имя файла:

`YYYYMMDD_DOC_TYPE_SUBJECT_vNN.ext`

## 8. Автоматическое сканирование и архивирование

### Что автоматизировано

- акты скрытых работ
- паспорта сотрудников
- приказы

Также доступно управление сканированием с МФУ из CLI через WIA.

### Шаги работы

1. Отсканировать документ в PDF/JPG/PNG.
   - вручную и положить в inbox
   - или сканировать прямо командой `scanner_control.py scan-to-inbox`
2. Переименовать по формату:

   `YYYYMMDD__DOC_TYPE__SUBJECT__[EMPLOYEE_ID].ext`

3. Положить файл во входящую папку объекта:

   `10_scan_inbox/`

4. Запустить ingest:

   ```bash
   python scripts/scan_archive_docs.py ingest --object-root docflow/objects/x5-ufa-e2_logistics_park
   ```

   OCR включен по умолчанию; если движок не установлен, архивирование продолжится без OCR.

5. Проверить результат:
   - разложенные файлы по целевым папкам
   - нераспознанные файлы в `10_scan_inbox/manual_review/`
   - причины в `10_scan_inbox/manual_review/review_log.csv`

### Папки назначения

- AWR -> `05_execution_docs/hidden_work_acts/`
- PASSPORT -> `02_personnel/employees/<id>_*/01_identity_and_contract/`
- ORDER -> `01_orders_and_appointments/`

### Диагностика и управление сканером (Windows/WIA)

```bash
python scripts/scanner_control.py diagnose
python scripts/scanner_control.py list
python scripts/scanner_control.py scan-to-inbox \
   --object-root docflow/objects/x5-ufa-e2_logistics_park \
   --doc-type ORDER \
   --subject appoint_hse_responsible \
   --device-index 1 \
   --format jpg \
   --dpi 300
```

Для нового ноутбука с Pantum BM2300W используйте helper-скрипт:

`scripts/setup_pantum_bm2300w.ps1`

### Архивный пакет на сдачу

```bash
python scripts/scan_archive_docs.py bundle \
  --object-root docflow/objects/x5-ufa-e2_logistics_park \
  --from-date 2026-03-01 \
  --to-date 2026-03-31
```

Результат: ZIP-файл в `09_archive/scan_bundles/`.

Подробный quickstart:

[docs/templates/execution/scan_ingest_quickstart.md](../templates/execution/scan_ingest_quickstart.md)

Инструкция для Pantum BM2300W:

[docs/document_control/scanner_setup_pantum_bm2300w.md](scanner_setup_pantum_bm2300w.md)

## 9. Где смотреть шаблоны вживую

Открыть индекс:

[docs/templates/INDEX.md](../templates/INDEX.md)

Дальше открыть Markdown Preview и переходить по ссылкам в документе.

## 10. Контур развития (учтено на будущее)

Модуль снабжения можно добавить как отдельный блок:

- реестр заявок на материалы
- статусы согласования и закупки
- график поставки vs план работ
- уведомления о дефицитах и влиянии на сроки

Текущая структура проекта это позволяет без переделки базового контура документооборота.
