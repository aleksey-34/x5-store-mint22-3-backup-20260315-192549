from fastapi.testclient import TestClient

from app.main import app
from app.services.local_llm import LocalLLMResult


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
