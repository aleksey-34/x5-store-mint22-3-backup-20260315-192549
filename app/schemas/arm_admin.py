from datetime import datetime

from pydantic import BaseModel, Field


class ArmChecklistItem(BaseModel):
    code: str
    title: str
    location: str
    required_min: int = Field(ge=1)
    found: int = Field(ge=0)
    ready: bool


class ArmMetrics(BaseModel):
    db_documents_total: int = Field(ge=0)
    db_journal_entries_total: int = Field(ge=0)
    db_schedules_total: int = Field(ge=0)
    orders_md_total: int = Field(ge=0)
    orders_pdf_ready_total: int = Field(ge=0)
    journals_production_total: int = Field(ge=0)
    journals_labor_safety_total: int = Field(ge=0)
    scan_inbox_pending_total: int = Field(ge=0)
    scan_manual_review_total: int = Field(ge=0)


class ArmDashboardResponse(BaseModel):
    generated_at: datetime
    object_root: str
    checklist_total: int = Field(ge=0)
    checklist_ready: int = Field(ge=0)
    checklist_progress_percent: float = Field(ge=0.0, le=100.0)
    local_llm_reachable: bool
    local_llm_version: str | None = None
    top_gaps: list[str]
    metrics: ArmMetrics
    checklist: list[ArmChecklistItem]


class ArmTodoItem(BaseModel):
    priority: str
    title: str
    details: str | None = None


class ArmTodoResponse(BaseModel):
    generated_at: datetime
    object_root: str
    items: list[ArmTodoItem]


class ArmAssistRequest(BaseModel):
    question: str = Field(min_length=1)
    model: str | None = None
    temperature: float = Field(default=0.2, ge=0.0, le=1.5)
    num_predict: int = Field(default=320, ge=1, le=4096)


class ArmAssistResponse(BaseModel):
    model: str
    response: str
    done: bool
    total_duration_sec: float | None = None
    eval_tokens: int | None = None
    eval_tokens_per_sec: float | None = None
