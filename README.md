# X5_Storage

X5_Storage - это локальное АРМ документооборота стройплощадки для зам. прораба/прораба:

- ведение исполнительной и сопроводительной документации
- ведение графиков и факта выполнения работ
- ведение журналов ОТ/ПБ по согласованному набору
- обработка ежедневных сообщений из Telegram по правилам

## Что это в итоге

Сейчас это рабочее приложение + шаблоны + структура хранения:

- backend API (FastAPI)
- локальная база SQLite
- стандарт папок и шаблоны документов
- скрипт генерации дерева объекта и досье сотрудников

Пока система работает локально из этой рабочей папки, но архитектура уже подходит для оборачивания в полноценный продукт (web-интерфейс, сервер, роли, отчеты).

## На базе чего работает

- Python 3.11+
- FastAPI
- SQLAlchemy
- SQLite
- Telethon (каркас интеграции Telegram)
- Ollama (локальная LLM, опционально)

## Быстрый запуск

1. Создать и активировать виртуальное окружение.
2. Установить зависимости:

   pip install -r requirements.txt

3. Скопировать шаблон переменных:

   copy .env.example .env

4. Запустить API:

   uvicorn app.main:app --reload

5. Открыть Swagger UI:

   http://127.0.0.1:8000/docs

## Переход на Linux Mint 22.3

Для ускорения работы на Linux Mint 22.3 добавлена отдельная подготовка окружения.

1. Выполнить bootstrap-скрипт:

```bash
bash scripts/setup_linux_mint_22_3.sh
```

2. Запуск API через Linux-скрипт:

```bash
bash scripts/run_api_clean.sh --port 8000 --reload
```

3. В VS Code доступна задача:

`Run FastAPI API (Linux Mint)`

Детальный чеклист миграции: [docs/migration/linux_mint_22_3_migration.md](docs/migration/linux_mint_22_3_migration.md)

## Как это ведется в работе

Полная инструкция: [docs/document_control/operator_manual.md](docs/document_control/operator_manual.md)

Короткий цикл:

1. Заполняете реквизиты организации и состав сотрудников.
2. Генерируете структуру папок объекта.
3. Заполняете шаблоны по перечню "+".
4. Каждый день фиксируете факт выполнения и сообщения прорабов.
5. Выгружаете, печатаете, подписываете и архивируете документы.

## Исходные данные от вас

- реквизиты организации
- список сотрудников, должности, удостоверения и сроки
- ежедневные сообщения прорабов в Telegram

## Генерация структуры по сотрудникам

Заполните файл [docs/templates/personnel/employees_sample.csv](docs/templates/personnel/employees_sample.csv), затем выполните:

```bash
python scripts/bootstrap_object_docs.py \
  --object-code X5-UFA-E2 \
  --object-name logistics_park \
  --employees-csv docs/templates/personnel/employees_sample.csv
```

Результат создается в [docflow/objects](docflow/objects).

## Telegram в ежедневной работе

Поддерживаемые API-маршруты:

- /telegram/rules/ - создать и посмотреть правила
- /telegram/messages/process - обработать входящее сообщение
- /journal/ - увидеть результат ручной или автоматической фиксации
- /arm/metrics - сводка по комплектности и метрикам объекта
- /arm/todo/today - список приоритетных задач на день
- /arm/checklist - статус позиций чек-листа Форма 2
- /arm/assist - вопрос к локальной LLM с контекстом объекта
- /arm/dashboard - легкая веб-страница с KPI/TODO/пробелами

Локальная LLM (Ollama):

- /local-llm/status - проверить доступность локальной модели
- /local-llm/runtime - проверить runtime, активные модели и CPU/GPU режим
- /local-llm/profiles - посмотреть быстрые профили quality/speed
- /local-llm/chat - отправить запрос в локальную модель с контекстом объекта
- /local-llm/chat/profile - запрос в модель через профиль с fallback на легкую модель

Пример обработки сообщения:

```json
{
  "message_text": "Участок Б3: сварка завершена, бетон завтра",
  "chat_id": "foreman_chat_1"
}
```

Пример запроса к локальной LLM:

```json
{
   "prompt": "Сформируй TODO на сегодня по документообороту",
   "context": "Объект: Уфа Х5, фокус: журналы ОТ/ПБ и наряды",
   "temperature": 0.2,
   "num_predict": 220
}
```

## Автоматизация сканов и архива

Поддержаны типы сканов:

- акты скрытых работ (`AWR`)
- паспорта сотрудников (`PASSPORT`)
- приказы (`ORDER`)

Формат имени входящего скана:

`YYYYMMDD__DOC_TYPE__SUBJECT__[EMPLOYEE_ID].ext`

Команда автоматической раскладки:

```bash
python scripts/scan_archive_docs.py ingest --object-root docflow/objects/x5-ufa-e2_logistics_park
```

OCR включен по умолчанию (rus+eng). При успешном распознавании создается sidecar-файл `*.ocr.txt` рядом с архивным файлом.

Управление сканером (WIA, Windows):

```bash
python scripts/scanner_control.py diagnose
python scripts/scanner_control.py list
python scripts/scanner_control.py scan-to-inbox \
   --object-root docflow/objects/x5-ufa-e2_logistics_park \
   --doc-type AWR \
   --subject foundation_axis_1_7 \
   --device-index 1 \
   --format jpg \
   --dpi 300
```

Первичная настройка нового ноутбука под Pantum BM2300W:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_pantum_bm2300w.ps1
```

Команда архивации в ZIP за период:

```bash
python scripts/scan_archive_docs.py bundle \
   --object-root docflow/objects/x5-ufa-e2_logistics_park \
   --from-date 2026-03-01 \
   --to-date 2026-03-31
```

MVP-классификация сканов в ручной очереди:

```bash
python scripts/scan_archive_docs.py classify \
   --object-root docflow/objects/x5-ufa-e2_logistics_park \
   --recursive
```

Нераспознанные файлы автоматически уходят в `10_scan_inbox/manual_review/`.

## Как печатать и в каком формате сохранять

Рекомендуемый стандарт сдачи:

1. Рабочая версия: DOCX/XLSX или MD по шаблону.
2. Подписанная версия: PDF (A4, портрет/альбом по форме).
3. Бумажный экземпляр: печать и подписи по требованиям заказчика.
4. Скан подписанного документа: PDF, читаемый, 300 dpi.
5. Хранить две копии: редактируемую и подписанную PDF.

Именование файлов:

`YYYYMMDD_DOC_TYPE_SUBJECT_vNN.ext`

Пример:

`20260309_ORDER_HSE_RESPONSIBLE_v01.pdf`

## Где смотреть шаблоны

Откройте [docs/templates/INDEX.md](docs/templates/INDEX.md) и используйте Markdown Preview для просмотра и перехода по ссылкам.

## Что уже включено в шаблоны

- реквизиты организации
- стартовый чеклист документов
- чеклист ежедневного/еженедельного контроля
- карточка сотрудника и реестр аттестаций
- матрица журналов по пунктам "+"
- реестр исполнительной документации
- seed-файл правил Telegram

## План расширения (на будущее)

В архитектуре можно добавить модуль снабжения без слома текущей логики:

- заявки и потребности по материалам
- сопоставление план/факт поставок
- контроль дефицитов и рисков сроков
- связка с Telegram-уведомлениями и журналом событий
