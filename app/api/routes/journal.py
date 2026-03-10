from fastapi import APIRouter, Depends, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.models.journal_entry import JournalEntry
from app.schemas.journal_entry import JournalEntryCreate, JournalEntryRead

router = APIRouter(prefix="/journal", tags=["journal"])


@router.post("/", response_model=JournalEntryRead, status_code=status.HTTP_201_CREATED)
def create_journal_entry(payload: JournalEntryCreate, db: Session = Depends(get_db)) -> JournalEntry:
    entry = JournalEntry(
        category=payload.category,
        content=payload.content,
        source=payload.source,
    )
    db.add(entry)
    db.commit()
    db.refresh(entry)
    return entry


@router.get("/", response_model=list[JournalEntryRead])
def list_journal_entries(db: Session = Depends(get_db)) -> list[JournalEntry]:
    entries = db.execute(select(JournalEntry).order_by(JournalEntry.created_at.desc())).scalars().all()
    return list(entries)
