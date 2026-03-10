from datetime import datetime

from pydantic import BaseModel, ConfigDict


class DocumentCreate(BaseModel):
    title: str
    doc_type: str
    status: str = "draft"
    file_path: str | None = None
    notes: str | None = None


class DocumentRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    title: str
    doc_type: str
    status: str
    file_path: str | None
    notes: str | None
    created_at: datetime
    updated_at: datetime
