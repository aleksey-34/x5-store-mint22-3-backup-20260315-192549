from pathlib import Path

from fastapi.testclient import TestClient

from app.main import app
from app.services.local_llm import LocalLLMResult


def _prepare_object_root(root: Path) -> None:
    (root / "01_orders_and_appointments").mkdir(parents=True, exist_ok=True)
    (root / "01_orders_and_appointments" / "print_pdf_ready").mkdir(parents=True, exist_ok=True)
    (root / "04_journals" / "production").mkdir(parents=True, exist_ok=True)
    (root / "04_journals" / "labor_safety").mkdir(parents=True, exist_ok=True)
    (root / "05_execution_docs" / "ppr").mkdir(parents=True, exist_ok=True)
    (root / "05_execution_docs" / "pprv_work_at_height").mkdir(parents=True, exist_ok=True)
    (root / "05_execution_docs" / "admission_acts").mkdir(parents=True, exist_ok=True)
    (root / "02_personnel" / "employees").mkdir(parents=True, exist_ok=True)
    (root / "06_normative_base").mkdir(parents=True, exist_ok=True)
    (root / "10_scan_inbox").mkdir(parents=True, exist_ok=True)


def test_arm_metrics_endpoint(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    (root / "01_orders_and_appointments" / "20260311_ORDER_01_test_v01.md").write_text("ok", encoding="utf-8")
    (root / "01_orders_and_appointments" / "20260311_ORDER_02_test_v01.md").write_text("ok", encoding="utf-8")
    (root / "01_orders_and_appointments" / "print_pdf_ready" / "order_01.pdf").write_text("pdf", encoding="utf-8")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.check_local_llm_available", lambda: (True, "0.17.7"))

    with TestClient(app) as client:
        response = client.get("/arm/metrics")

    assert response.status_code == 200
    payload = response.json()
    assert payload["object_root"].endswith("object")
    assert payload["checklist_total"] > 0
    assert payload["metrics"]["orders_md_total"] >= 2
    assert payload["local_llm_reachable"] is True


def test_arm_todo_today_has_items_for_gaps(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.check_local_llm_available", lambda: (False, None))

    with TestClient(app) as client:
        response = client.get("/arm/todo/today")

    assert response.status_code == 200
    payload = response.json()
    assert len(payload["items"]) > 0


def test_arm_assist_uses_local_llm(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)
    (root / "01_orders_and_appointments" / "20260311_ORDER_01_test_v01.md").write_text("ok", encoding="utf-8")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.check_local_llm_available", lambda: (True, "0.17.7"))

    def fake_generate_with_local_llm(**_: object) -> LocalLLMResult:
        return LocalLLMResult(
            model="llama3.2:3b",
            response="Готово",
            done=True,
            total_duration_sec=0.7,
            eval_tokens=20,
            eval_tokens_per_sec=28.5,
        )

    monkeypatch.setattr("app.api.routes.arm_admin.generate_with_local_llm", fake_generate_with_local_llm)

    with TestClient(app) as client:
        response = client.post("/arm/assist", json={"question": "Что сделать сегодня?"})

    assert response.status_code == 200
    payload = response.json()
    assert payload["response"] == "Готово"
    assert payload["model"] == "llama3.2:3b"
