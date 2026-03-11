from typing import Literal

from pydantic import BaseModel, Field


LocalLLMProfileId = Literal["fast", "balanced", "quality"]


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


class LocalLLMRuntimeModelResponse(BaseModel):
    name: str
    digest: str | None = None
    parameter_size: str | None = None
    quantization_level: str | None = None
    size_bytes: int | None = None
    size_vram_bytes: int | None = None
    expires_at: str | None = None


class LocalLLMRuntimeResponse(BaseModel):
    enabled: bool
    base_url: str
    default_model: str
    is_reachable: bool
    version: str | None = None
    acceleration: str
    running_models_count: int
    running_models: list[LocalLLMRuntimeModelResponse]


class LocalLLMProfilePresetResponse(BaseModel):
    profile: LocalLLMProfileId
    title: str
    description: str
    model: str
    temperature: float = Field(ge=0.0, le=1.5)
    num_predict: int = Field(ge=1, le=4096)


class LocalLLMProfilesResponse(BaseModel):
    default_profile: LocalLLMProfileId
    recommended_profile: LocalLLMProfileId
    fallback_model: str
    presets: list[LocalLLMProfilePresetResponse]


class LocalLLMProfileChatRequest(BaseModel):
    prompt: str = Field(min_length=1)
    context: str | None = None
    profile: LocalLLMProfileId = "balanced"
    model: str | None = None
    system_prompt: str | None = None
    temperature: float | None = Field(default=None, ge=0.0, le=1.5)
    num_predict: int | None = Field(default=None, ge=1, le=4096)
    allow_fallback: bool = True


class LocalLLMProfileChatResponse(BaseModel):
    model: str
    response: str
    done: bool
    used_profile: LocalLLMProfileId
    fallback_used: bool
    total_duration_sec: float | None = None
    eval_tokens: int | None = None
    eval_tokens_per_sec: float | None = None
