from fastapi.testclient import TestClient

from app.main import app
from app.services.local_llm import (
    LocalLLMConnectionError,
    LocalLLMProfilePreset,
    LocalLLMResult,
    LocalLLMRuntimeModel,
    LocalLLMRuntimeSnapshot,
)


def test_local_llm_status_endpoint(monkeypatch) -> None:
    def fake_status() -> tuple[bool, str | None]:
        return True, "0.17.7"

    monkeypatch.setattr("app.api.routes.local_llm.check_local_llm_available", fake_status)

    with TestClient(app) as client:
        response = client.get("/local-llm/status")

    assert response.status_code == 200
    payload = response.json()
    assert payload["is_reachable"] is True
    assert payload["version"] == "0.17.7"


def test_local_llm_chat_endpoint(monkeypatch) -> None:
    def fake_generate_with_local_llm(**_: object) -> LocalLLMResult:
        return LocalLLMResult(
            model="llama3.2:3b",
            response="ok",
            done=True,
            total_duration_sec=1.2,
            eval_tokens=42,
            eval_tokens_per_sec=35.0,
        )

    monkeypatch.setattr("app.api.routes.local_llm.generate_with_local_llm", fake_generate_with_local_llm)

    with TestClient(app) as client:
        response = client.post(
            "/local-llm/chat",
            json={
                "prompt": "Сформируй краткий план на день",
                "context": "Объект Уфа Х5",
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["model"] == "llama3.2:3b"
    assert payload["response"] == "ok"
    assert payload["eval_tokens"] == 42


def test_local_llm_runtime_endpoint(monkeypatch) -> None:
    def fake_snapshot() -> LocalLLMRuntimeSnapshot:
        return LocalLLMRuntimeSnapshot(
            version="0.17.7",
            acceleration="cpu",
            running_models=[
                LocalLLMRuntimeModel(
                    name="llama3.2:3b",
                    digest="abc",
                    parameter_size="3.2B",
                    quantization_level="Q4_K_M",
                    size_bytes=2_000_000_000,
                    size_vram_bytes=0,
                    expires_at="2026-03-11T10:00:00Z",
                )
            ],
        )

    monkeypatch.setattr("app.api.routes.local_llm.fetch_local_llm_runtime_snapshot", fake_snapshot)

    with TestClient(app) as client:
        response = client.get("/local-llm/runtime")

    assert response.status_code == 200
    payload = response.json()
    assert payload["is_reachable"] is True
    assert payload["acceleration"] == "cpu"
    assert payload["running_models_count"] == 1
    assert payload["running_models"][0]["name"] == "llama3.2:3b"


def test_local_llm_runtime_endpoint_unreachable(monkeypatch) -> None:
    def fake_snapshot() -> LocalLLMRuntimeSnapshot:
        raise LocalLLMConnectionError("offline")

    monkeypatch.setattr("app.api.routes.local_llm.fetch_local_llm_runtime_snapshot", fake_snapshot)

    with TestClient(app) as client:
        response = client.get("/local-llm/runtime")

    assert response.status_code == 200
    payload = response.json()
    assert payload["is_reachable"] is False
    assert payload["acceleration"] == "unreachable"


def test_local_llm_profiles_endpoint(monkeypatch) -> None:
    def fake_presets() -> list[LocalLLMProfilePreset]:
        return [
            LocalLLMProfilePreset(
                profile="fast",
                title="Fast",
                description="Quick replies",
                model="qwen2.5:1.5b",
                temperature=0.15,
                num_predict=160,
            ),
            LocalLLMProfilePreset(
                profile="balanced",
                title="Balanced",
                description="Default",
                model="llama3.2:3b",
                temperature=0.2,
                num_predict=320,
            ),
        ]

    def fake_snapshot() -> LocalLLMRuntimeSnapshot:
        return LocalLLMRuntimeSnapshot(version="0.17.7", running_models=[], acceleration="cpu")

    monkeypatch.setattr("app.api.routes.local_llm.get_local_llm_profile_presets", fake_presets)
    monkeypatch.setattr("app.api.routes.local_llm.fetch_local_llm_runtime_snapshot", fake_snapshot)

    with TestClient(app) as client:
        response = client.get("/local-llm/profiles")

    assert response.status_code == 200
    payload = response.json()
    assert payload["default_profile"] == "balanced"
    assert payload["recommended_profile"] == "fast"
    assert len(payload["presets"]) == 2


def test_local_llm_chat_profile_endpoint(monkeypatch) -> None:
    def fake_profile_chat(**_: object) -> tuple[LocalLLMResult, str, bool]:
        return (
            LocalLLMResult(
                model="qwen2.5:1.5b",
                response="ok-fast",
                done=True,
                total_duration_sec=0.9,
                eval_tokens=60,
                eval_tokens_per_sec=55.0,
            ),
            "fast",
            True,
        )

    monkeypatch.setattr("app.api.routes.local_llm.generate_with_local_llm_profile", fake_profile_chat)

    with TestClient(app) as client:
        response = client.post(
            "/local-llm/chat/profile",
            json={
                "prompt": "TODO на сегодня",
                "profile": "balanced",
                "allow_fallback": True,
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["response"] == "ok-fast"
    assert payload["used_profile"] == "fast"
    assert payload["fallback_used"] is True
