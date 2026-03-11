from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import httpx

from app.core.config import settings


class LocalLLMConnectionError(RuntimeError):
    """Raised when the local LLM runtime is unavailable."""


class LocalLLMRequestError(RuntimeError):
    """Raised when local LLM request fails with an API-side error."""


@dataclass
class LocalLLMResult:
    model: str
    response: str
    done: bool
    total_duration_sec: float | None
    eval_tokens: int | None
    eval_tokens_per_sec: float | None


@dataclass
class LocalLLMRuntimeModel:
    name: str
    digest: str | None
    parameter_size: str | None
    quantization_level: str | None
    size_bytes: int | None
    size_vram_bytes: int | None
    expires_at: str | None


@dataclass
class LocalLLMRuntimeSnapshot:
    version: str | None
    running_models: list[LocalLLMRuntimeModel]
    acceleration: str


LocalLLMProfileId = Literal["fast", "balanced", "quality"]


@dataclass(frozen=True)
class LocalLLMProfilePreset:
    profile: LocalLLMProfileId
    title: str
    description: str
    model: str
    temperature: float
    num_predict: int


def _build_prompt(prompt: str, context: str | None) -> str:
    base_prompt = prompt.strip()
    if not context:
        return base_prompt

    context_text = context.strip()
    if not context_text:
        return base_prompt

    return f"Контекст объекта:\n{context_text}\n\nЗадача:\n{base_prompt}"


def _to_seconds(value_ns: int | None) -> float | None:
    if value_ns is None:
        return None
    return round(float(value_ns) / 1_000_000_000, 3)


def get_local_llm_profile_presets() -> list[LocalLLMProfilePreset]:
    fallback_model = settings.local_llm_fallback_model.strip() or settings.local_llm_model
    base_model = settings.local_llm_model.strip()

    return [
        LocalLLMProfilePreset(
            profile="fast",
            title="Быстрый черновик",
            description="Минимальная задержка для коротких задач прораба.",
            model=fallback_model,
            temperature=0.15,
            num_predict=160,
        ),
        LocalLLMProfilePreset(
            profile="balanced",
            title="Сбалансированный",
            description="Стандартный режим для ежедневных рабочих ответов.",
            model=base_model,
            temperature=0.2,
            num_predict=320,
        ),
        LocalLLMProfilePreset(
            profile="quality",
            title="Подробный",
            description="Больше текста и контекста, выше задержка.",
            model=base_model,
            temperature=0.25,
            num_predict=560,
        ),
    ]


def _find_profile_preset(profile: LocalLLMProfileId) -> LocalLLMProfilePreset:
    presets = {item.profile: item for item in get_local_llm_profile_presets()}
    return presets[profile]


def check_local_llm_available() -> tuple[bool, str | None]:
    if not settings.local_llm_enabled:
        return False, None

    url = f"{settings.local_llm_base_url.rstrip('/')}/api/version"
    try:
        with httpx.Client(timeout=5.0, trust_env=False) as client:
            response = client.get(url)
            response.raise_for_status()
            payload = response.json()
            version = payload.get("version")
            return True, str(version) if version is not None else None
    except Exception:  # noqa: BLE001
        return False, None


def _infer_acceleration(running_models: list[LocalLLMRuntimeModel]) -> str:
    if not running_models:
        return "idle"

    has_vram = any((model.size_vram_bytes or 0) > 0 for model in running_models)
    return "gpu_or_hybrid" if has_vram else "cpu"


def fetch_local_llm_runtime_snapshot() -> LocalLLMRuntimeSnapshot:
    if not settings.local_llm_enabled:
        raise LocalLLMConnectionError("Local LLM integration is disabled in settings")

    base_url = settings.local_llm_base_url.rstrip("/")
    timeout = max(10, settings.local_llm_timeout_seconds)

    try:
        with httpx.Client(timeout=timeout, trust_env=False) as client:
            version_response = client.get(f"{base_url}/api/version")
            version_response.raise_for_status()
            version_payload = version_response.json()

            ps_response = client.get(f"{base_url}/api/ps")
            ps_response.raise_for_status()
            ps_payload = ps_response.json()
    except Exception as exc:  # noqa: BLE001
        raise LocalLLMConnectionError(
            f"Failed to connect to local LLM at {settings.local_llm_base_url}"
        ) from exc

    models_payload = ps_payload.get("models", [])
    running_models: list[LocalLLMRuntimeModel] = []

    for item in models_payload:
        details = item.get("details") if isinstance(item, dict) else None
        details = details if isinstance(details, dict) else {}

        size = item.get("size") if isinstance(item, dict) else None
        size_vram = item.get("size_vram") if isinstance(item, dict) else None

        running_models.append(
            LocalLLMRuntimeModel(
                name=str(item.get("name", "")),
                digest=str(item.get("digest")) if item.get("digest") is not None else None,
                parameter_size=(
                    str(details.get("parameter_size"))
                    if details.get("parameter_size") is not None
                    else None
                ),
                quantization_level=(
                    str(details.get("quantization_level"))
                    if details.get("quantization_level") is not None
                    else None
                ),
                size_bytes=int(size) if isinstance(size, (int, float)) else None,
                size_vram_bytes=(
                    int(size_vram) if isinstance(size_vram, (int, float)) else None
                ),
                expires_at=(
                    str(item.get("expires_at"))
                    if isinstance(item, dict) and item.get("expires_at") is not None
                    else None
                ),
            )
        )

    return LocalLLMRuntimeSnapshot(
        version=(
            str(version_payload.get("version"))
            if isinstance(version_payload, dict) and version_payload.get("version") is not None
            else None
        ),
        running_models=running_models,
        acceleration=_infer_acceleration(running_models),
    )


def generate_with_local_llm(
    *,
    prompt: str,
    context: str | None,
    model: str | None,
    system_prompt: str | None,
    temperature: float,
    num_predict: int,
) -> LocalLLMResult:
    if not settings.local_llm_enabled:
        raise LocalLLMConnectionError("Local LLM integration is disabled in settings")

    selected_model = model or settings.local_llm_model
    selected_system_prompt = system_prompt or settings.local_llm_system_prompt
    composed_prompt = _build_prompt(prompt=prompt, context=context)

    request_payload: dict[str, object] = {
        "model": selected_model,
        "prompt": composed_prompt,
        "system": selected_system_prompt,
        "stream": False,
        "options": {
            "temperature": temperature,
            "num_predict": num_predict,
        },
    }

    url = f"{settings.local_llm_base_url.rstrip('/')}/api/generate"
    timeout = max(10, settings.local_llm_timeout_seconds)

    try:
        with httpx.Client(timeout=timeout, trust_env=False) as client:
            response = client.post(url, json=request_payload)
    except Exception as exc:  # noqa: BLE001
        raise LocalLLMConnectionError(
            f"Failed to connect to local LLM at {settings.local_llm_base_url}"
        ) from exc

    if response.status_code >= 400:
        raise LocalLLMRequestError(
            f"Local LLM API returned status {response.status_code}: {response.text[:400]}"
        )

    payload = response.json()
    eval_tokens = payload.get("eval_count")
    eval_duration_ns = payload.get("eval_duration")
    eval_tokens_per_sec: float | None = None

    eval_duration_sec = _to_seconds(eval_duration_ns)
    if eval_tokens is not None and eval_duration_sec and eval_duration_sec > 0:
        eval_tokens_per_sec = round(float(eval_tokens) / eval_duration_sec, 2)

    return LocalLLMResult(
        model=str(payload.get("model", selected_model)),
        response=str(payload.get("response", "")).strip(),
        done=bool(payload.get("done", False)),
        total_duration_sec=_to_seconds(payload.get("total_duration")),
        eval_tokens=int(eval_tokens) if eval_tokens is not None else None,
        eval_tokens_per_sec=eval_tokens_per_sec,
    )


def generate_with_local_llm_profile(
    *,
    prompt: str,
    context: str | None,
    profile: LocalLLMProfileId,
    model: str | None,
    system_prompt: str | None,
    temperature: float | None,
    num_predict: int | None,
    allow_fallback: bool,
) -> tuple[LocalLLMResult, LocalLLMProfileId, bool]:
    preset = _find_profile_preset(profile)

    selected_model = model or preset.model
    selected_temperature = temperature if temperature is not None else preset.temperature
    selected_num_predict = num_predict if num_predict is not None else preset.num_predict

    try:
        result = generate_with_local_llm(
            prompt=prompt,
            context=context,
            model=selected_model,
            system_prompt=system_prompt,
            temperature=selected_temperature,
            num_predict=selected_num_predict,
        )
        return result, profile, False
    except LocalLLMRequestError as primary_error:
        fallback_candidates: list[tuple[str, LocalLLMProfileId]] = []

        configured_fallback = settings.local_llm_fallback_model.strip()
        if configured_fallback and configured_fallback != selected_model:
            fallback_candidates.append((configured_fallback, "fast"))

        base_model = settings.local_llm_model.strip()
        if base_model and base_model != selected_model and all(m != base_model for m, _ in fallback_candidates):
            fallback_candidates.append((base_model, "balanced"))

        if not allow_fallback or not fallback_candidates:
            raise

        last_error: Exception = primary_error
        for candidate_model, candidate_profile in fallback_candidates:
            try:
                fallback_result = generate_with_local_llm(
                    prompt=prompt,
                    context=context,
                    model=candidate_model,
                    system_prompt=system_prompt,
                    temperature=min(selected_temperature, 0.2),
                    num_predict=min(selected_num_predict, 220),
                )
                return fallback_result, candidate_profile, True
            except LocalLLMRequestError as fallback_error:
                last_error = fallback_error

        raise last_error
