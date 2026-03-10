from datetime import date

from pydantic import BaseModel, ConfigDict, Field


class WorkScheduleCreate(BaseModel):
    title: str
    planned_start: date
    planned_end: date
    actual_start: date | None = None
    actual_end: date | None = None
    progress_percent: float = Field(default=0.0, ge=0.0, le=100.0)
    status: str = "planned"
    notes: str | None = None


class WorkScheduleRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    title: str
    planned_start: date
    planned_end: date
    actual_start: date | None
    actual_end: date | None
    progress_percent: float
    status: str
    notes: str | None
