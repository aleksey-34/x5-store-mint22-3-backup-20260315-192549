from fastapi.testclient import TestClient

from app.main import app


def test_telegram_message_processing_triggers_rule() -> None:
    with TestClient(app) as client:
        create_rule_response = client.post(
            "/telegram/rules/",
            json={
                "keyword": "delay",
                "action": "mark_delay_risk",
                "is_active": True,
                "description": "Detect delay words",
            },
        )
        assert create_rule_response.status_code == 201

        process_response = client.post(
            "/telegram/messages/process",
            json={"message_text": "Concrete delivery delay on sector B", "chat_id": "site-chat"},
        )
        assert process_response.status_code == 200

        payload = process_response.json()
        assert payload["matched_rules"] >= 1
        assert payload["actions"][0]["status"] in {"risk_logged", "note_created", "no_handler"}
