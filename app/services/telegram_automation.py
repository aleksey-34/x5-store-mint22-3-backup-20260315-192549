from sqlalchemy import select
from sqlalchemy.orm import Session

from app.models.journal_entry import JournalEntry
from app.models.telegram_rule import TelegramRule
from app.schemas.telegram import TelegramActionResult, TelegramMessageIn, TelegramMessageProcessResponse


def process_message(db: Session, payload: TelegramMessageIn) -> TelegramMessageProcessResponse:
    message_text = payload.message_text.strip()

    db.add(
        JournalEntry(
            category="telegram",
            content=f"Incoming message: {message_text}",
            source="telegram",
        )
    )

    rules = list(
        db.execute(select(TelegramRule).where(TelegramRule.is_active.is_(True))).scalars().all()
    )

    actions: list[TelegramActionResult] = []
    lowered_message = message_text.lower()

    for rule in rules:
        if rule.keyword.lower() in lowered_message:
            action_status = _execute_action(db, rule, payload)
            actions.append(
                TelegramActionResult(
                    rule_id=rule.id,
                    action=rule.action,
                    status=action_status,
                )
            )

    db.commit()

    return TelegramMessageProcessResponse(
        matched_rules=len(actions),
        actions=actions,
    )


def _execute_action(db: Session, rule: TelegramRule, payload: TelegramMessageIn) -> str:
    if rule.action == "create_journal_note":
        db.add(
            JournalEntry(
                category="telegram_rule",
                content=(
                    f"Rule '{rule.keyword}' matched in chat {payload.chat_id or 'unknown'}: "
                    f"{payload.message_text}"
                ),
                source="automation",
            )
        )
        return "note_created"

    if rule.action == "mark_delay_risk":
        db.add(
            JournalEntry(
                category="risk",
                content=f"Delay risk detected by rule '{rule.keyword}': {payload.message_text}",
                source="automation",
            )
        )
        return "risk_logged"

    return "no_handler"
