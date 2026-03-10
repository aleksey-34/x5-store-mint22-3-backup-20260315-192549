from datetime import datetime

from pydantic import BaseModel, ConfigDict


class JournalEntryCreate(BaseModel):
    category: str
    content: str
    source: str = "manual"


class JournalEntryRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    category: str
    content: str
    source: str
    created_at: datetime
