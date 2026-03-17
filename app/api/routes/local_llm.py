from fastapi import APIRouter, HTTPException, status

from app.core.config import settings
from app.schemas.local_llm import (
    LocalLLMChatRequest,
    LocalLLMChatResponse,
    LocalLLMProfileChatRequest,
    LocalLLMProfileChatResponse,
    LocalLLMProfilePresetResponse,
    LocalLLMProfilesResponse,
    LocalLLMRuntimeModelResponse,
    LocalLLMRuntimeResponse,
    LocalLLMStatusResponse,
)
from app.services.local_llm import (
    LocalLLMConnectionError,
    LocalLLMRequestError,
    check_local_llm_available,
    fetch_local_llm_runtime_snapshot,
    generate_with_local_llm,
    generate_with_local_llm_profile,
    get_local_llm_profile_presets,
)

router = APIRouter(prefix="/local-llm", tags=["local-llm"])


@router.get("/status", response_model=LocalLLMStatusResponse)
def local_llm_status() -> LocalLLMStatusResponse:
    is_reachable, version = check_local_llm_available()
    return LocalLLMStatusResponse(
        enabled=settings.local_llm_enabled,
        base_url=settings.local_llm_base_url,
        default_model=settings.local_llm_model,
        is_reachable=is_reachable,
        version=version,
    )


@router.get("/runtime", response_model=LocalLLMRuntimeResponse)
def local_llm_runtime() -> LocalLLMRuntimeResponse:
    if not settings.local_llm_enabled:
        return LocalLLMRuntimeResponse(
            enabled=False,
            base_url=settings.local_llm_base_url,
            default_model=settings.local_llm_model,
            is_reachable=False,
            version=None,
            acceleration="disabled",
            running_models_count=0,
            running_models=[],
        )

    try:
        snapshot = fetch_local_llm_runtime_snapshot()
    except LocalLLMConnectionError:
        return LocalLLMRuntimeResponse(
            enabled=True,
            base_url=settings.local_llm_base_url,
            default_model=settings.local_llm_model,
            is_reachable=False,
            version=None,
            acceleration="unreachable",
            running_models_count=0,
            running_models=[],
        )

    return LocalLLMRuntimeResponse(
        enabled=True,
        base_url=settings.local_llm_base_url,
        default_model=settings.local_llm_model,
        is_reachable=True,
        version=snapshot.version,
        acceleration=snapshot.acceleration,
        running_models_count=len(snapshot.running_models),
        running_models=[
            LocalLLMRuntimeModelResponse(
                name=model.name,
                digest=model.digest,
                parameter_size=model.parameter_size,
                quantization_level=model.quantization_level,
                size_bytes=model.size_bytes,
                size_vram_bytes=model.size_vram_bytes,
                expires_at=model.expires_at,
            )
            for model in snapshot.running_models
        ],
    )


@router.get("/profiles", response_model=LocalLLMProfilesResponse)
def local_llm_profiles() -> LocalLLMProfilesResponse:
    recommended = "balanced"
    if settings.local_llm_enabled:
        try:
            snapshot = fetch_local_llm_runtime_snapshot()
            if snapshot.acceleration in {"cpu", "idle"}:
                recommended = "fast"
        except LocalLLMConnectionError:
            recommended = "fast"

    presets = get_local_llm_profile_presets()
    return LocalLLMProfilesResponse(
        default_profile="balanced",
        recommended_profile=recommended,
        fallback_model=settings.local_llm_fallback_model,
        presets=[
            LocalLLMProfilePresetResponse(
                profile=item.profile,
                title=item.title,
                description=item.description,
                model=item.model,
                temperature=item.temperature,
                num_predict=item.num_predict,
            )
            for item in presets
        ],
    )


@router.post("/chat", response_model=LocalLLMChatResponse)
def local_llm_chat(payload: LocalLLMChatRequest) -> LocalLLMChatResponse:
    if not settings.local_llm_enabled:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Local LLM integration is disabled",
        )

    try:
        result = generate_with_local_llm(
            prompt=payload.prompt,
            context=payload.context,
            model=payload.model,
            system_prompt=payload.system_prompt,
            temperature=payload.temperature,
            num_predict=payload.num_predict,
        )
    except LocalLLMConnectionError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc),
        ) from exc
    except LocalLLMRequestError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=str(exc),
        ) from exc

    return LocalLLMChatResponse(
        model=result.model,
        response=result.response,
        done=result.done,
        total_duration_sec=result.total_duration_sec,
        eval_tokens=result.eval_tokens,
        eval_tokens_per_sec=result.eval_tokens_per_sec,
    )


@router.post("/chat/profile", response_model=LocalLLMProfileChatResponse)
def local_llm_chat_profile(payload: LocalLLMProfileChatRequest) -> LocalLLMProfileChatResponse:
    if not settings.local_llm_enabled:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Local LLM integration is disabled",
        )

    try:
        result, used_profile, fallback_used = generate_with_local_llm_profile(
            prompt=payload.prompt,
            context=payload.context,
            profile=payload.profile,
            model=payload.model,
            system_prompt=payload.system_prompt,
            temperature=payload.temperature,
            num_predict=payload.num_predict,
            allow_fallback=payload.allow_fallback,
        )
    except LocalLLMConnectionError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc),
        ) from exc
    except LocalLLMRequestError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=str(exc),
        ) from exc

    return LocalLLMProfileChatResponse(
        model=result.model,
        response=result.response,
        done=result.done,
        used_profile=used_profile,
        fallback_used=fallback_used,
        total_duration_sec=result.total_duration_sec,
        eval_tokens=result.eval_tokens,
        eval_tokens_per_sec=result.eval_tokens_per_sec,
    )
