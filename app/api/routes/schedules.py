from fastapi import APIRouter, Depends, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.models.work_schedule import WorkSchedule
from app.schemas.work_schedule import WorkScheduleCreate, WorkScheduleRead

router = APIRouter(prefix="/schedules", tags=["schedules"])


@router.post("/", response_model=WorkScheduleRead, status_code=status.HTTP_201_CREATED)
def create_schedule(payload: WorkScheduleCreate, db: Session = Depends(get_db)) -> WorkSchedule:
    schedule = WorkSchedule(
        title=payload.title,
        planned_start=payload.planned_start,
        planned_end=payload.planned_end,
        actual_start=payload.actual_start,
        actual_end=payload.actual_end,
        progress_percent=payload.progress_percent,
        status=payload.status,
        notes=payload.notes,
    )
    db.add(schedule)
    db.commit()
    db.refresh(schedule)
    return schedule


@router.get("/", response_model=list[WorkScheduleRead])
def list_schedules(db: Session = Depends(get_db)) -> list[WorkSchedule]:
    schedules = db.execute(select(WorkSchedule).order_by(WorkSchedule.planned_start.asc())).scalars().all()
    return list(schedules)
