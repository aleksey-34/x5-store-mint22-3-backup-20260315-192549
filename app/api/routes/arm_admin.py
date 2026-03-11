from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from html import escape
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import HTMLResponse
from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app.core.config import settings
from app.db.session import get_db
from app.models.document import Document
from app.models.journal_entry import JournalEntry
from app.models.work_schedule import WorkSchedule
from app.schemas.arm_admin import (
    ArmAssistRequest,
    ArmAssistResponse,
    ArmChecklistItem,
    ArmDashboardResponse,
    ArmMetrics,
    ArmTodoItem,
    ArmTodoResponse,
)
from app.services.local_llm import (
    LocalLLMConnectionError,
    LocalLLMRequestError,
    check_local_llm_available,
    generate_with_local_llm_profile,
)

router = APIRouter(prefix="/arm", tags=["arm"])


@dataclass(frozen=True)
class ChecklistRule:
    code: str
    title: str
    folder: str
    pattern: str
    required_min: int


CHECKLIST_RULES: tuple[ChecklistRule, ...] = (
    ChecklistRule("1.1.1", "Приказ: уполномоченный представитель", "01_orders_and_appointments", "*ORDER_01_*", 1),
    ChecklistRule("1.1.2", "Приказ: допуск к СМР", "01_orders_and_appointments", "*ORDER_02_*", 1),
    ChecklistRule("1.1.3", "Приказы: ответственные по направлениям", "01_orders_and_appointments", "*ORDER_03_*", 1),
    ChecklistRule("1.1.4", "Приказ: стропальщики", "01_orders_and_appointments", "*ORDER_04_*", 1),
    ChecklistRule("1.1.5", "Приказ: электрохозяйство", "01_orders_and_appointments", "*ORDER_05_*", 1),
    ChecklistRule("1.1.6", "Приказ: стажировка рабочих", "01_orders_and_appointments", "*ORDER_06_*", 1),
    ChecklistRule("1.1.7", "Приказ: допуск по профессиям", "01_orders_and_appointments", "*ORDER_07_*", 1),
    ChecklistRule("1.1.8", "Приказы: распределение ответственности по прорабам", "01_orders_and_appointments", "*ORDER_09_*", 1),
    ChecklistRule("1.1.9", "Приказы: ТБ по прорабам", "01_orders_and_appointments", "*ORDER_10_*", 1),
    ChecklistRule("1.1.10", "Приказы: ТБ по прорабам (2)", "01_orders_and_appointments", "*ORDER_11_*", 1),
    ChecklistRule("1.3.3", "Наряд-допуск на опасные работы", "01_orders_and_appointments", "*PERMIT_12_*", 1),
    ChecklistRule("1.4", "ППР", "05_execution_docs/ppr", "*", 1),
    ChecklistRule("1.5", "ППРв", "05_execution_docs/pprv_work_at_height", "*", 1),
    ChecklistRule("1.6", "Акт-допуск на производство СМР", "05_execution_docs/admission_acts", "*", 1),
    ChecklistRule("3.1", "Журналы производства", "04_journals/production", "*", 6),
    ChecklistRule("3.2", "Журналы ОТ/ПБ", "04_journals/labor_safety", "*", 9),
    ChecklistRule("4", "Сканы удостоверений и протоколов", "02_personnel/employees", "*.pdf", 8),
    ChecklistRule("5", "Нормативная база", "06_normative_base", "*", 1),
)


def resolve_object_root() -> Path:
    return Path(settings.object_root).resolve()


def _count_files(root: Path, folder: str, pattern: str) -> int:
    target = root / folder
    if not target.exists() or not target.is_dir():
        return 0

    if pattern == "*":
        return sum(1 for item in target.rglob("*") if item.is_file())

    return sum(1 for item in target.rglob(pattern) if item.is_file())


def _build_checklist(root: Path) -> list[ArmChecklistItem]:
    items: list[ArmChecklistItem] = []
    for rule in CHECKLIST_RULES:
        found = _count_files(root=root, folder=rule.folder, pattern=rule.pattern)
        items.append(
            ArmChecklistItem(
                code=rule.code,
                title=rule.title,
                location=str((root / rule.folder).as_posix()),
                required_min=rule.required_min,
                found=found,
                ready=found >= rule.required_min,
            )
        )
    return items


def _collect_metrics(db: Session, root: Path) -> ArmMetrics:
    db_documents_total = int(db.execute(select(func.count(Document.id))).scalar_one())
    db_journal_entries_total = int(db.execute(select(func.count(JournalEntry.id))).scalar_one())
    db_schedules_total = int(db.execute(select(func.count(WorkSchedule.id))).scalar_one())

    orders_md_total = _count_files(root=root, folder="01_orders_and_appointments", pattern="*.md")
    orders_pdf_ready_total = _count_files(
        root=root,
        folder="01_orders_and_appointments/print_pdf_ready",
        pattern="*.pdf",
    )
    journals_production_total = _count_files(root=root, folder="04_journals/production", pattern="*")
    journals_labor_safety_total = _count_files(root=root, folder="04_journals/labor_safety", pattern="*")

    scan_root = root / "10_scan_inbox"
    scan_inbox_pending_total = 0
    if scan_root.exists():
        scan_inbox_pending_total = sum(1 for item in scan_root.iterdir() if item.is_file())

    scan_manual_review_total = _count_files(
        root=root,
        folder="10_scan_inbox/manual_review",
        pattern="*",
    )

    return ArmMetrics(
        db_documents_total=db_documents_total,
        db_journal_entries_total=db_journal_entries_total,
        db_schedules_total=db_schedules_total,
        orders_md_total=orders_md_total,
        orders_pdf_ready_total=orders_pdf_ready_total,
        journals_production_total=journals_production_total,
        journals_labor_safety_total=journals_labor_safety_total,
        scan_inbox_pending_total=scan_inbox_pending_total,
        scan_manual_review_total=scan_manual_review_total,
    )


def _build_todos(
    checklist: list[ArmChecklistItem],
    metrics: ArmMetrics,
    local_llm_reachable: bool,
) -> list[ArmTodoItem]:
    todos: list[ArmTodoItem] = []

    for item in checklist:
        if item.ready:
            continue
        todos.append(
            ArmTodoItem(
                priority="high" if item.code in {"3.1", "3.2", "4"} else "medium",
                title=f"Закрыть позицию {item.code}: {item.title}",
                details=f"Найдено {item.found} из минимум {item.required_min}",
            )
        )

    if metrics.scan_manual_review_total > 0:
        todos.append(
            ArmTodoItem(
                priority="high",
                title="Разобрать manual_review после сканирования",
                details=f"В очереди {metrics.scan_manual_review_total} файлов",
            )
        )

    if metrics.orders_pdf_ready_total == 0 and metrics.orders_md_total > 0:
        todos.append(
            ArmTodoItem(
                priority="medium",
                title="Собрать пакет PDF для печати",
                details="Запустить scripts/build_and_open_print_pack.ps1",
            )
        )

    if not local_llm_reachable:
        todos.append(
            ArmTodoItem(
                priority="medium",
                title="Проверить доступность локальной LLM",
                details="Проверить Ollama: /local-llm/status",
            )
        )

    # Keep daily list compact for foreman workflow.
    return todos[:12]


def _build_dashboard_payload(db: Session) -> ArmDashboardResponse:
    root = resolve_object_root()
    checklist = _build_checklist(root=root)
    metrics = _collect_metrics(db=db, root=root)
    local_llm_reachable, local_llm_version = check_local_llm_available()

    checklist_total = len(checklist)
    checklist_ready = sum(1 for item in checklist if item.ready)
    checklist_progress_percent = round(
        (checklist_ready / checklist_total) * 100.0 if checklist_total else 0.0,
        1,
    )
    top_gaps = [item.title for item in checklist if not item.ready][:6]

    return ArmDashboardResponse(
        generated_at=datetime.now(tz=timezone.utc),
        object_root=str(root.as_posix()),
        checklist_total=checklist_total,
        checklist_ready=checklist_ready,
        checklist_progress_percent=checklist_progress_percent,
        local_llm_reachable=local_llm_reachable,
        local_llm_version=local_llm_version,
        top_gaps=top_gaps,
        metrics=metrics,
        checklist=checklist,
    )


def _build_arm_context(payload: ArmDashboardResponse, todos: ArmTodoResponse) -> str:
    todo_lines = "\n".join(f"- [{item.priority}] {item.title}" for item in todos.items)
    gaps_lines = "\n".join(f"- {item}" for item in payload.top_gaps)

    return (
        "Контекст АРМ объекта:\n"
        f"- Путь объекта: {payload.object_root}\n"
        f"- Комплектность по чек-листу: {payload.checklist_ready}/{payload.checklist_total} "
        f"({payload.checklist_progress_percent}%)\n"
        f"- Документов в БД: {payload.metrics.db_documents_total}\n"
        f"- Записей журнала в БД: {payload.metrics.db_journal_entries_total}\n"
        f"- Графиков в БД: {payload.metrics.db_schedules_total}\n"
        f"- PDF ready: {payload.metrics.orders_pdf_ready_total}\n"
        f"- Pending scan inbox: {payload.metrics.scan_inbox_pending_total}\n"
        f"- Manual review: {payload.metrics.scan_manual_review_total}\n"
        "Основные пробелы:\n"
        f"{gaps_lines or '- Нет критичных пробелов'}\n"
        "TODO на сегодня:\n"
        f"{todo_lines or '- TODO пуст'}"
    )


@router.get("/checklist", response_model=list[ArmChecklistItem])
def arm_checklist() -> list[ArmChecklistItem]:
    root = resolve_object_root()
    return _build_checklist(root=root)


@router.get("/metrics", response_model=ArmDashboardResponse)
def arm_metrics(db: Session = Depends(get_db)) -> ArmDashboardResponse:
    return _build_dashboard_payload(db=db)


@router.get("/todo/today", response_model=ArmTodoResponse)
def arm_todo_today(db: Session = Depends(get_db)) -> ArmTodoResponse:
    payload = _build_dashboard_payload(db=db)
    items = _build_todos(
        checklist=payload.checklist,
        metrics=payload.metrics,
        local_llm_reachable=payload.local_llm_reachable,
    )
    return ArmTodoResponse(
        generated_at=payload.generated_at,
        object_root=payload.object_root,
        items=items,
    )


@router.post("/assist", response_model=ArmAssistResponse)
def arm_assist(payload: ArmAssistRequest, db: Session = Depends(get_db)) -> ArmAssistResponse:
    if not settings.local_llm_enabled:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Local LLM integration is disabled",
        )

    dashboard = _build_dashboard_payload(db=db)
    todos = ArmTodoResponse(
        generated_at=dashboard.generated_at,
        object_root=dashboard.object_root,
        items=_build_todos(
            checklist=dashboard.checklist,
            metrics=dashboard.metrics,
            local_llm_reachable=dashboard.local_llm_reachable,
        ),
    )
    context = _build_arm_context(payload=dashboard, todos=todos)

    try:
        result, used_profile, fallback_used = generate_with_local_llm_profile(
            prompt=payload.question,
            context=context,
            profile=payload.profile,
            model=payload.model,
            system_prompt=None,
            temperature=payload.temperature,
            num_predict=payload.num_predict,
            allow_fallback=payload.allow_fallback,
        )
    except LocalLLMConnectionError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc),
        ) from exc
    except LocalLLMRequestError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=str(exc),
        ) from exc

    return ArmAssistResponse(
        model=result.model,
        response=result.response,
        done=result.done,
        used_profile=used_profile,
        fallback_used=fallback_used,
        total_duration_sec=result.total_duration_sec,
        eval_tokens=result.eval_tokens,
        eval_tokens_per_sec=result.eval_tokens_per_sec,
    )


@router.get("/dashboard", response_class=HTMLResponse)
def arm_dashboard_html(db: Session = Depends(get_db)) -> HTMLResponse:
    payload = _build_dashboard_payload(db=db)
    todos = _build_todos(
        checklist=payload.checklist,
        metrics=payload.metrics,
        local_llm_reachable=payload.local_llm_reachable,
    )

    metrics_html = "".join(
        f"<li><b>{escape(key)}:</b> {escape(str(value))}</li>"
        for key, value in payload.metrics.model_dump().items()
    )
    gaps_html = "".join(f"<li>{escape(item)}</li>" for item in payload.top_gaps) or "<li>Пробелов не найдено</li>"
    todo_html = "".join(
        f"<li>[{escape(item.priority)}] {escape(item.title)}"
        + (f" - {escape(item.details)}" if item.details else "")
        + "</li>"
        for item in todos
    ) or "<li>TODO пуст</li>"

    html = f"""
<!doctype html>
<html lang=\"ru\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>АРМ объекта: дашборд</title>
  <style>
    :root {{
      --bg: #f4efe7;
      --card: #fffaf3;
      --ink: #1f2a36;
      --accent: #0f766e;
      --warn: #c2410c;
      --muted: #5b6672;
    }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Tahoma, sans-serif;
      background: linear-gradient(135deg, #f4efe7 0%, #e9f2f4 100%);
      color: var(--ink);
    }}
    .wrap {{
      max-width: 1100px;
      margin: 24px auto;
      padding: 0 16px 24px;
    }}
    .hero {{
      background: var(--card);
      border: 1px solid #d8cfc2;
      border-radius: 16px;
      padding: 18px;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.06);
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
      gap: 16px;
      margin-top: 16px;
    }}
    .card {{
      background: var(--card);
      border: 1px solid #d8cfc2;
      border-radius: 16px;
      padding: 14px;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.05);
    }}
    h1 {{ margin: 0 0 8px; font-size: 24px; }}
    h2 {{ margin: 0 0 8px; font-size: 18px; color: var(--accent); }}
    .meta {{ color: var(--muted); font-size: 14px; }}
    .kpi {{ font-size: 34px; font-weight: 700; color: var(--accent); margin: 8px 0; }}
    .warn {{ color: var(--warn); }}
    ul {{ margin: 8px 0 0; padding-left: 18px; }}
    li {{ margin: 6px 0; }}
  </style>
</head>
<body>
  <div class=\"wrap\">
    <section class=\"hero\">
      <h1>АРМ объекта: X5 UFA E2</h1>
      <div class=\"meta\">Object root: {escape(payload.object_root)}</div>
      <div class=\"meta\">Generated at (UTC): {escape(payload.generated_at.isoformat())}</div>
      <div class=\"kpi\">{payload.checklist_progress_percent}%</div>
      <div>Комплектность: {payload.checklist_ready}/{payload.checklist_total}</div>
      <div class=\"meta\">LLM: {'reachable' if payload.local_llm_reachable else 'not reachable'} {escape(payload.local_llm_version or '')}</div>
    </section>

    <section class=\"grid\">
      <article class=\"card\">
        <h2>Метрики</h2>
        <ul>{metrics_html}</ul>
      </article>
      <article class=\"card\">
        <h2>Критичные пробелы</h2>
        <ul>{gaps_html}</ul>
      </article>
      <article class=\"card\">
        <h2>TODO на сегодня</h2>
        <ul>{todo_html}</ul>
      </article>
    </section>
  </div>
</body>
</html>
"""
    return HTMLResponse(content=html)
