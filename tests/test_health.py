from fastapi.testclient import TestClient

from app.main import app


def test_healthcheck() -> None:
    with TestClient(app) as client:
        response = client.get("/health")
        assert response.status_code == 200

        payload = response.json()
        assert payload["status"] == "ok"
