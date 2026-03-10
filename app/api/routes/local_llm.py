from fastapi import APIRouter, HTTPException, status

from app.core.config import settings
from app.schemas.local_llm import (
    LocalLLMChatRequest,
    LocalLLMChatResponse,
    LocalLLMStatusResponse,
)
from app.services.local_llm import (
    LocalLLMConnectionError,
    LocalLLMRequestError,
    check_local_llm_available,
    generate_with_local_llm,
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
