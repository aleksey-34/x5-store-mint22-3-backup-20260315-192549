from datetime import datetime

from pydantic import BaseModel, ConfigDict


class TelegramRuleCreate(BaseModel):
    keyword: str
    action: str
    is_active: bool = True
    description: str | None = None


class TelegramRuleRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    keyword: str
    action: str
    is_active: bool
    description: str | None
    created_at: datetime


class TelegramMessageIn(BaseModel):
    message_text: str
    chat_id: str | None = None


class TelegramActionResult(BaseModel):
    rule_id: int
    action: str
    status: str


class TelegramMessageProcessResponse(BaseModel):
    matched_rules: int
    actions: list[TelegramActionResult]
