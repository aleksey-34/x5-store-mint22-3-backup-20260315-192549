from pydantic import BaseModel, Field


class LocalLLMChatRequest(BaseModel):
    prompt: str = Field(min_length=1)
    context: str | None = None
    model: str | None = None
    system_prompt: str | None = None
    temperature: float = Field(default=0.2, ge=0.0, le=1.5)
    num_predict: int = Field(default=256, ge=1, le=4096)


class LocalLLMChatResponse(BaseModel):
    model: str
    response: str
    done: bool
    total_duration_sec: float | None = None
    eval_tokens: int | None = None
    eval_tokens_per_sec: float | None = None


class LocalLLMStatusResponse(BaseModel):
    enabled: bool
    base_url: str
    default_model: str
    is_reachable: bool
    version: str | None = None
