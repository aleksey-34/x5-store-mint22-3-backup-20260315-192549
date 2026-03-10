from fastapi import APIRouter, Depends, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.models.telegram_rule import TelegramRule
from app.schemas.telegram import (
    TelegramMessageIn,
    TelegramMessageProcessResponse,
    TelegramRuleCreate,
    TelegramRuleRead,
)
from app.services.telegram_automation import process_message

router = APIRouter(prefix="/telegram", tags=["telegram"])


@router.post("/rules/", response_model=TelegramRuleRead, status_code=status.HTTP_201_CREATED)
def create_rule(payload: TelegramRuleCreate, db: Session = Depends(get_db)) -> TelegramRule:
    rule = TelegramRule(
        keyword=payload.keyword,
        action=payload.action,
        is_active=payload.is_active,
        description=payload.description,
    )
    db.add(rule)
    db.commit()
    db.refresh(rule)
    return rule


@router.get("/rules/", response_model=list[TelegramRuleRead])
def list_rules(db: Session = Depends(get_db)) -> list[TelegramRule]:
    rules = db.execute(select(TelegramRule).order_by(TelegramRule.created_at.desc())).scalars().all()
    return list(rules)


@router.post("/messages/process", response_model=TelegramMessageProcessResponse)
def process_telegram_message(
    payload: TelegramMessageIn,
    db: Session = Depends(get_db),
) -> TelegramMessageProcessResponse:
    return process_message(db, payload)
