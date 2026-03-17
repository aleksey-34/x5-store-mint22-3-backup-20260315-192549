from fastapi import APIRouter, Depends, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.models.document import Document
from app.schemas.document import (
    DocumentCreate,
    DocumentDeletionMarkUpdate,
    DocumentRead,
    DocumentStatusUpdate,
)

router = APIRouter(prefix="/documents", tags=["documents"])

DOCUMENT_STATUS_VALUES = {"approved", "new", "fix"}


@router.post("/", response_model=DocumentRead, status_code=status.HTTP_201_CREATED)
def create_document(payload: DocumentCreate, db: Session = Depends(get_db)) -> Document:
    status_value = payload.status.strip().lower() if payload.status else "new"
    if status_value not in DOCUMENT_STATUS_VALUES:
        status_value = "new"

    item = Document(
        title=payload.title,
        doc_type=payload.doc_type,
        status=status_value,
        file_path=payload.file_path,
        notes=payload.notes,
        fix_comment=(payload.fix_comment or "").strip() or None,
        marked_for_deletion=bool(payload.marked_for_deletion),
    )
    db.add(item)
    db.commit()
    db.refresh(item)
    return item


@router.get("/", response_model=list[DocumentRead])
def list_documents(db: Session = Depends(get_db)) -> list[Document]:
    documents = db.execute(select(Document).order_by(Document.created_at.desc())).scalars().all()
    return list(documents)


@router.patch("/{document_id}/status", response_model=DocumentRead)
def update_document_status(document_id: int, payload: DocumentStatusUpdate, db: Session = Depends(get_db)) -> Document:
    item = db.get(Document, document_id)
    if item is None:
        from fastapi import HTTPException

        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Документ не найден")

    status_value = (payload.status or "").strip().lower()
    if status_value not in DOCUMENT_STATUS_VALUES:
        from fastapi import HTTPException

        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Недопустимый статус")

    item.status = status_value
    item.fix_comment = (payload.fix_comment or "").strip() or None
    db.add(item)
    db.commit()
    db.refresh(item)
    return item


@router.patch("/{document_id}/deletion-mark", response_model=DocumentRead)
def update_document_deletion_mark(
    document_id: int,
    payload: DocumentDeletionMarkUpdate,
    db: Session = Depends(get_db),
) -> Document:
    item = db.get(Document, document_id)
    if item is None:
        from fastapi import HTTPException

        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Документ не найден")

    item.marked_for_deletion = bool(payload.marked_for_deletion)
    db.add(item)
    db.commit()
    db.refresh(item)
    return item
