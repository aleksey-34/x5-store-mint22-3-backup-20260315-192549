from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field


ArmAssistProfile = Literal["fast", "balanced", "quality"]


class ArmChecklistItem(BaseModel):
    code: str
    title: str
    location: str
    required_min: int = Field(ge=1)
    found: int = Field(ge=0)
    ready: bool
    found_files: list[str] = Field(default_factory=list)


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
    action_path: str | None = None


class ArmTodoResponse(BaseModel):
    generated_at: datetime
    object_root: str
    items: list[ArmTodoItem]


class ArmAssistRequest(BaseModel):
    question: str = Field(min_length=1)
    profile: ArmAssistProfile = "balanced"
    allow_fallback: bool = True
    model: str | None = None
    temperature: float | None = Field(default=None, ge=0.0, le=1.5)
    num_predict: int | None = Field(default=None, ge=1, le=4096)


class ArmAssistResponse(BaseModel):
    model: str
    response: str
    done: bool
    used_profile: ArmAssistProfile
    fallback_used: bool
    total_duration_sec: float | None = None
    eval_tokens: int | None = None
    eval_tokens_per_sec: float | None = None


class ArmFsEntry(BaseModel):
    name: str
    rel_path: str
    is_dir: bool
    size: int | None = None
    modified_at: str | None = None


class ArmFsTreeResponse(BaseModel):
    root: str
    rel_path: str
    entries: list[ArmFsEntry]


class ArmFileReadResponse(BaseModel):
    rel_path: str
    content: str
    encoding: str = "utf-8"


class ArmFileWriteRequest(BaseModel):
    rel_path: str = Field(min_length=1)
    content: str


class ArmActionResponse(BaseModel):
    ok: bool
    message: str


class ArmObjectProfileResponse(BaseModel):
    object_name: str = ""
    project_code: str = ""
    organization: str = ""
    work_stage: str = ""
    start_date: str | None = None
    metadata_rel_path: str
    ppr_source_options: list[str] = Field(default_factory=list)
    ppr_context_rel_path: str | None = None


class ArmObjectProfileUpdateRequest(BaseModel):
    object_name: str | None = None
    project_code: str | None = None
    organization: str | None = None
    work_stage: str | None = None
    start_date: str | None = Field(default=None, pattern=r"^$|^\d{2}\.\d{2}\.\d{4}$")


class ArmPprImportRequest(BaseModel):
    rel_path: str = Field(min_length=1)


class ArmSpeechTranscribeResponse(BaseModel):
    ok: bool
    text: str = ""
    provider: str
    message: str


class ArmScannerDevice(BaseModel):
    index: int
    name: str
    device_id: str | None = None


class ArmScannerDevicesResponse(BaseModel):
    devices: list[ArmScannerDevice]


class ArmScanCaptureRequest(BaseModel):
    doc_type: str = Field(min_length=1)
    subject: str = Field(min_length=1)
    employee_id: str | None = None
    device_index: int = Field(default=1, ge=1)
    image_format: str = Field(default="jpg", min_length=1)
    scan_profile: int = Field(default=1, ge=1, le=3)
    dpi: int = Field(default=300, ge=75, le=1200)
    grayscale: bool = False


class ArmScanIngestRequest(BaseModel):
    enable_ocr: bool = True
    ocr_lang: str = "rus+eng"
    tesseract_cmd: str | None = None
    max_pdf_pages: int = Field(default=4, ge=1, le=30)


class ArmScanIngestItem(BaseModel):
    source_name: str
    status: str
    message: str
    destination: str | None = None
    document_id: int | None = None
    ocr_status: str
    ocr_text_path: str | None = None
    suggested_doc_type: str | None = None
    suggested_confidence: float | None = None


class ArmScanIngestResponse(BaseModel):
    archived: int = Field(ge=0)
    manual_review: int = Field(ge=0)
    items: list[ArmScanIngestItem]


EmployeeChecklistActionMode = Literal["selected", "missing", "all"]
ProfessionGroupKey = Literal["default", "electric", "supervisor", "itr", "custom"]


class ArmEmployeeChecklistItem(BaseModel):
    code: str
    title: str
    folder_rel_path: str
    expected_patterns: list[str]
    required_count: int = Field(default=1, ge=1)
    found_count: int = Field(ge=0)
    ready: bool
    found_files: list[str]
    related_count: int = Field(default=0, ge=0)
    related_files: list[str] = Field(default_factory=list)
    guidance: str


class ArmEmployeeChecklistResponse(BaseModel):
    employee_rel_path: str
    employee_id: str | None = None
    employee_name: str | None = None
    profile_position: str | None = None
    profession: str
    total_required: int = Field(ge=0)
    ready_count: int = Field(ge=0)
    missing_count: int = Field(ge=0)
    progress_percent: float = Field(ge=0.0, le=100.0)
    items: list[ArmEmployeeChecklistItem]


class ArmEmployeeChecklistGenerateRequest(BaseModel):
    employee_rel_path: str = Field(min_length=1)
    profession: str | None = None
    order_date: str | None = Field(default=None, pattern=r"^\d{2}\.\d{2}\.\d{4}$")
    mode: EmployeeChecklistActionMode = "missing"
    codes: list[str] = Field(default_factory=list)
    overwrite: bool = False


class ArmEmployeeChecklistGenerateResponse(BaseModel):
    ok: bool
    employee_rel_path: str
    profession: str
    mode: EmployeeChecklistActionMode
    created_files: list[str]
    skipped_files: list[str]
    message: str


class ArmProfessionOption(BaseModel):
    key: ProfessionGroupKey
    label: str


class ArmEmployeeCatalogItem(BaseModel):
    employee_rel_path: str
    employee_id: str | None = None
    employee_name: str
    position: str | None = None
    profession_group: ProfessionGroupKey
    profession_label: str


class ArmEmployeeCatalogResponse(BaseModel):
    total: int = Field(ge=0)
    items: list[ArmEmployeeCatalogItem]
    profession_options: list[ArmProfessionOption]


class ArmEmployeeOverviewAction(BaseModel):
    code: str
    title: str
    guidance: str
    scope: str
    missing_employees: int = Field(ge=0)


class ArmEmployeeOverviewEmployee(BaseModel):
    employee_rel_path: str
    employee_id: str | None = None
    employee_name: str
    position: str | None = None
    profession: str
    total_required: int = Field(ge=0)
    ready_count: int = Field(ge=0)
    missing_count: int = Field(ge=0)
    progress_percent: float = Field(ge=0.0, le=100.0)
    top_missing_codes: list[str]


class ArmEmployeeOverviewGroup(BaseModel):
    profession_group: ProfessionGroupKey
    profession_label: str
    employees_total: int = Field(ge=0)
    ready_employees: int = Field(ge=0)
    average_progress_percent: float = Field(ge=0.0, le=100.0)
    missing_actions: list[ArmEmployeeOverviewAction]
    employees: list[ArmEmployeeOverviewEmployee]


class ArmEmployeeChecklistOverviewResponse(BaseModel):
    generated_at: datetime
    profession_filter: str | None = None
    groups: list[ArmEmployeeOverviewGroup]
