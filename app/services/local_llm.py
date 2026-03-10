from __future__ import annotations

from dataclasses import dataclass

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


def check_local_llm_available() -> tuple[bool, str | None]:
    if not settings.local_llm_enabled:
        return False, None

    url = f"{settings.local_llm_base_url.rstrip('/')}/api/version"
    try:
        with httpx.Client(timeout=5.0) as client:
            response = client.get(url)
            response.raise_for_status()
            payload = response.json()
            version = payload.get("version")
            return True, str(version) if version is not None else None
    except Exception:  # noqa: BLE001
        return False, None


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
        with httpx.Client(timeout=timeout) as client:
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
