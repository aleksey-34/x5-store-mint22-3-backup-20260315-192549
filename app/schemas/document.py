from datetime import datetime

from pydantic import BaseModel, ConfigDict


class DocumentCreate(BaseModel):
    title: str
    doc_type: str
    status: str = "new"
    file_path: str | None = None
    notes: str | None = None
    fix_comment: str | None = None
    marked_for_deletion: bool = False


class DocumentStatusUpdate(BaseModel):
    status: str
    fix_comment: str | None = None


class DocumentDeletionMarkUpdate(BaseModel):
    marked_for_deletion: bool


class DocumentRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    title: str
    doc_type: str
    status: str
    file_path: str | None
    notes: str | None
    fix_comment: str | None
    marked_for_deletion: bool
    created_at: datetime
    updated_at: datetime
