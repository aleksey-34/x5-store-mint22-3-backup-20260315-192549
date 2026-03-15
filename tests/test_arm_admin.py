from pathlib import Path

from fastapi.testclient import TestClient
from sqlalchemy import select

from app.main import app
from app.db.session import SessionLocal
from app.models.document_content import DocumentContent
from app.services.local_llm import LocalLLMResult
from app.services.office_export import (
    OfficeDocxExportResult,
    OfficeXlsxExportResult,
    build_office_pack_zip,
)


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


def _prepare_employee(root: Path, employee_folder: str) -> Path:
    employee_root = root / "02_personnel" / "employees" / employee_folder
    employee_root.mkdir(parents=True, exist_ok=True)
    for section in [
        "01_identity_and_contract",
        "02_admission_orders",
        "03_briefings_and_training",
        "04_attestation_and_certificates",
        "05_ppe_issue",
        "06_permits_and_work_admission",
    ]:
        (employee_root / section).mkdir(parents=True, exist_ok=True)
    return employee_root


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

    def fake_generate_with_local_llm_profile(**_: object) -> tuple[LocalLLMResult, str, bool]:
        return (
            LocalLLMResult(
                model="llama3.2:3b",
                response="Готово",
                done=True,
                total_duration_sec=0.7,
                eval_tokens=20,
                eval_tokens_per_sec=28.5,
            ),
            "balanced",
            False,
        )

    monkeypatch.setattr(
        "app.api.routes.arm_admin.generate_with_local_llm_profile",
        fake_generate_with_local_llm_profile,
    )

    with TestClient(app) as client:
        response = client.post("/arm/assist", json={"question": "Что сделать сегодня?"})

    assert response.status_code == 200
    payload = response.json()
    assert payload["response"] == "Готово"
    assert payload["model"] == "llama3.2:3b"
    assert payload["used_profile"] == "balanced"
    assert payload["fallback_used"] is False


def test_arm_assist_scenario_creates_employee_order_and_vehicle_pass(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.settings.local_llm_enabled", False)

    question = (
        "добавь сотрудника Иванов Иван Иванович id 777 должность прораб "
        "и сделай приказ 12 и заявку на пропуск авто А123АА102"
    )

    with TestClient(app) as client:
        response = client.post("/arm/assist", json={"question": question})

    assert response.status_code == 200
    payload = response.json()
    assert payload["model"] == "scenario-engine"
    assert "Сценарий кадрового оформления выполнен" in payload["response"]

    employees_root = root / "02_personnel" / "employees"
    employee_dirs = sorted(employees_root.glob("777_*"))
    assert employee_dirs

    employee_root = employee_dirs[0]
    profile_path = employee_root / "employee_profile.txt"
    assert profile_path.exists()
    profile_text = profile_path.read_text(encoding="utf-8")
    assert "last_name: Иванов" in profile_text
    assert "position: прораб" in profile_text

    order_files = sorted((root / "01_orders_and_appointments" / "drafts_from_assistant").glob("*ORDER_12*"))
    assert order_files

    pass_files = sorted((root / "00_incoming_requests" / "drafts_from_assistant").glob("*LETTER_ADMISSION*"))
    assert pass_files
    pass_text = pass_files[0].read_text(encoding="utf-8")
    assert "А123АА102" in pass_text


def test_arm_assist_scenario_adds_multiple_employees_and_vehicle_pass(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.settings.local_llm_enabled", False)

    question = (
        "добавь сотрудников\n"
        "Лазарев Алексей Егорович 09.06.1987 Инженер ПТО\n"
        "Абдрахманов Вадим Нурмухаметович 29.03.1987 Стропальщик\n"
        "Бикбулатов Артур Римович 05.09.1986 Геодезист\n"
        "и сделай заявку на пропуск автомобиля"
    )

    with TestClient(app) as client:
        response = client.post("/arm/assist", json={"question": question})

    assert response.status_code == 200
    payload = response.json()
    assert payload["model"] == "scenario-engine"
    assert "Обработано сотрудников: 3" in payload["response"]

    employees_root = root / "02_personnel" / "employees"
    names = sorted(p.name for p in employees_root.iterdir() if p.is_dir())
    assert any("лазарев_алексей_егорович" in name for name in names)
    assert any("абдрахманов_вадим_нурмухаметович" in name for name in names)
    assert any("бикбулатов_артур_римович" in name for name in names)

    pass_files = sorted((root / "00_incoming_requests" / "drafts_from_assistant").glob("*LETTER_ADMISSION*"))
    assert pass_files
    pass_text = pass_files[0].read_text(encoding="utf-8")
    assert "Бикбулатов Артур Римович" in pass_text


def test_arm_assist_scenario_handles_comma_transport_batch(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    grech_root = _prepare_employee(root, "018_гречушников")
    yakupov_root = _prepare_employee(root, "013_якупов")
    bikbulatov_root = _prepare_employee(root, "017_бикбулатов_артур_римович")

    (grech_root / "employee_profile.txt").write_text(
        "employee_id: 018\n"
        "last_name: Гречушников\n"
        "first_name: Марсель\n"
        "middle_name: Ильдарович\n"
        "position: ИТР\n",
        encoding="utf-8",
    )
    (yakupov_root / "employee_profile.txt").write_text(
        "employee_id: 013\n"
        "last_name: Якупов\n"
        "first_name: Расим\n"
        "middle_name: Рафаилович\n"
        "position: Прораб\n",
        encoding="utf-8",
    )
    (bikbulatov_root / "employee_profile.txt").write_text(
        "employee_id: 017\n"
        "last_name: Бикбулатов\n"
        "first_name: Артур\n"
        "middle_name: Римович\n"
        "position: Геодезист\n",
        encoding="utf-8",
    )

    old_permit_md = bikbulatov_root / "06_permits_and_work_admission" / "20260312_LETTER_ADMISSION_old_v01.md"
    old_permit_pdf = bikbulatov_root / "06_permits_and_work_admission" / "20260312_LETTER_ADMISSION_old_v01.pdf"
    old_permit_md.write_text("старый допуск", encoding="utf-8")
    old_permit_pdf.write_bytes(b"%PDF-1.4")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.settings.local_llm_enabled", False)

    question = (
        "Привет даю задачу\n\n"
        "отредактировать и/или добавить данные к сотрудникам\n"
        "Гречушников, kia rio Н192МО 174\n"
        "Якупов, kia ceed Т916АК 774\n"
        "Бикбулатов, Skoda Rapid Р487ХМ 102\n\n"
        "добавить водителей и технику\n"
        "Сагитов Фанис Галимьянович, кран зулион К630НМ 702\n"
        "Гафаров Роберт Азаматович, xcmg с/у 25кK5S А674ВО 702\n\n"
        "удалить заявку на пропуск у Бикбулатова и сделать одну общую заявку на пропуск и папку с техникой если есть"
    )

    with TestClient(app) as client:
        response = client.post("/arm/assist", json={"question": question})

    assert response.status_code == 200
    payload = response.json()
    assert payload["model"] == "scenario-engine"
    assert "Обработано сотрудников: 5" in payload["response"]
    assert "Сформирована общая заявка на пропуск" in payload["response"]
    assert "Архивировано заявок по Бикбулатову" in payload["response"]

    assert not old_permit_md.exists()
    assert not old_permit_pdf.exists()
    archived = list((root / "09_archive" / "removed_admission_requests").glob("**/*LETTER_ADMISSION*"))
    assert len(archived) >= 2

    employee_dirs = sorted(p.name for p in (root / "02_personnel" / "employees").iterdir() if p.is_dir())
    assert any("сагитов_фанис_галимьянович" in name for name in employee_dirs)
    assert any("гафаров_роберт_азаматович" in name for name in employee_dirs)

    pass_files = sorted((root / "00_incoming_requests" / "drafts_from_assistant").glob("*LETTER_ADMISSION*"))
    assert pass_files
    pass_text = pass_files[-1].read_text(encoding="utf-8")
    assert "Гречушников" in pass_text
    assert "Якупов" in pass_text
    assert "Бикбулатов" in pass_text
    assert "Сагитов Фанис Галимьянович" in pass_text
    assert "Гафаров Роберт Азаматович" in pass_text
    assert "Н192МО174" in pass_text
    assert "Т916АК774" in pass_text
    assert "Р487ХМ102" in pass_text
    assert "К630НМ702" in pass_text
    assert "А674ВО702" in pass_text

    equipment_registry_files = list((root / "00_incoming_requests" / "equipment_from_assistant").glob("**/equipment_registry.md"))
    assert equipment_registry_files
    equipment_registry_text = equipment_registry_files[0].read_text(encoding="utf-8")
    assert "Сагитов" in equipment_registry_text
    assert "Гафаров" in equipment_registry_text


def test_arm_employee_checklist_endpoint(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)
    employee_root = _prepare_employee(root, "001_ivanov_ivan_ivanovich")

    (employee_root / "employee_profile.txt").write_text(
        "employee_id: 001\n"
        "last_name: Иванов\n"
        "first_name: Иван\n"
        "middle_name: Иванович\n"
        "position: электромонтажник\n",
        encoding="utf-8",
    )
    (employee_root / "02_admission_orders" / "admission_note.md").write_text("ok", encoding="utf-8")
    (root / "01_orders_and_appointments" / "20260311_ORDER_11_tb_roles_v01.md").write_text("ok", encoding="utf-8")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/employees/checklist",
            params={"employee_rel_path": "02_personnel/employees/001_ivanov_ivan_ivanovich"},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["employee_id"] == "001"
    assert payload["employee_name"] == "Иванов Иван Иванович"
    assert payload["total_required"] > 0
    assert any(item["code"] == "E01_PROFILE" and item["ready"] for item in payload["items"])


def test_arm_employee_checklist_generate_selected(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)
    _prepare_employee(root, "002_petrov_petr_petrovich")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.post(
            "/arm/employees/checklist/generate",
            json={
                "employee_rel_path": "02_personnel/employees/002_petrov_petr_petrovich",
                "mode": "selected",
                "codes": ["E03_ADMISSION_ORDERS"],
                "overwrite": True,
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert len(payload["created_files"]) == 1
    generated = root / payload["created_files"][0]
    assert generated.exists()


def test_arm_employee_checklist_generate_project_order_template(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)
    employee_root = _prepare_employee(root, "003_sidorov_sergey_alexeevich")
    (employee_root / "employee_profile.txt").write_text(
        "employee_id: 003\n"
        "last_name: Сидоров\n"
        "first_name: Сергей\n"
        "middle_name: Алексеевич\n"
        "position: производитель работ\n",
        encoding="utf-8",
    )

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.post(
            "/arm/employees/checklist/generate",
            json={
                "employee_rel_path": "02_personnel/employees/003_sidorov_sergey_alexeevich",
                "mode": "selected",
                "codes": ["P12_ORDER_PERMIT"],
                "overwrite": True,
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["created_files"]
    assert payload["created_files"][0].startswith("01_orders_and_appointments/drafts_from_checklist/")

    generated = root / payload["created_files"][0]
    assert generated.exists()

    text = generated.read_text(encoding="utf-8")
    assert "о назначении лиц, имеющих право выдачи наряда-допуска" in text
    assert "Индивидуальный предприниматель" in text
    assert "01.03.2026" in text


def test_arm_employee_checklist_generate_with_explicit_order_date(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)
    employee_root = _prepare_employee(root, "004_novikov_nikolay_sergeevich")
    (employee_root / "employee_profile.txt").write_text(
        "employee_id: 004\n"
        "last_name: Новиков\n"
        "first_name: Николай\n"
        "middle_name: Сергеевич\n"
        "position: производитель работ\n",
        encoding="utf-8",
    )

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.post(
            "/arm/employees/checklist/generate",
            json={
                "employee_rel_path": "02_personnel/employees/004_novikov_nikolay_sergeevich",
                "mode": "selected",
                "codes": ["P11_ORDER_PS"],
                "overwrite": True,
                "order_date": "05.02.2026",
            },
        )

    assert response.status_code == 200
    payload = response.json()
    generated = root / payload["created_files"][0]
    text = generated.read_text(encoding="utf-8")
    assert "05.02.2026" in text
    assert "01.03.2026" not in text


def test_arm_employee_catalog_and_overview_grouping(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    electric_root = _prepare_employee(root, "010_ivanov_ivan")
    supervisor_root = _prepare_employee(root, "011_petrov_petr")

    (electric_root / "employee_profile.txt").write_text(
        "employee_id: 010\n"
        "last_name: Иванов\n"
        "first_name: Иван\n"
        "position: электромонтажник\n",
        encoding="utf-8",
    )
    (supervisor_root / "employee_profile.txt").write_text(
        "employee_id: 011\n"
        "last_name: Петров\n"
        "first_name: Петр\n"
        "position: прораб\n",
        encoding="utf-8",
    )

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        catalog_response = client.get("/arm/employees/catalog")
        overview_response = client.get("/arm/employees/checklist/overview", params={"profession": "electric"})

    assert catalog_response.status_code == 200
    catalog_payload = catalog_response.json()
    assert catalog_payload["total"] == 2
    groups = {item["profession_group"] for item in catalog_payload["items"]}
    assert "electric" in groups
    assert "supervisor" in groups

    assert overview_response.status_code == 200
    overview_payload = overview_response.json()
    assert len(overview_payload["groups"]) == 1
    assert overview_payload["groups"][0]["profession_group"] == "electric"
    assert overview_payload["groups"][0]["employees_total"] == 1
    assert len(overview_payload["groups"][0]["missing_actions"]) > 0


def test_arm_export_orders_docx_classification_forwarded(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    office_dir = root / "01_orders_and_appointments" / "print_office"
    office_dir.mkdir(parents=True, exist_ok=True)
    bundle = office_dir / "bundle.zip"
    bundle.write_bytes(b"zip")

    observed: dict[str, str] = {}

    def fake_export_orders_docx_bundle(*, object_root: Path, classification: str) -> OfficeDocxExportResult:
        observed["classification"] = classification
        return OfficeDocxExportResult(output_dir=office_dir, bundle_path=bundle, files_count=1)

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.export_orders_docx_bundle", fake_export_orders_docx_bundle)

    with TestClient(app) as client:
        response = client.get("/arm/exports/orders-docx", params={"classification": "checklist-drafts"})

    assert response.status_code == 200
    assert observed["classification"] == "checklist-drafts"


def test_arm_export_orders_docx_invalid_classification(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)
    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get("/arm/exports/orders-docx", params={"classification": "unknown-group"})

    assert response.status_code == 400
    assert "Недопустимая классификация экспорта" in response.json()["detail"]


def test_arm_export_office_pack_classification_forwarded(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    office_dir = root / "01_orders_and_appointments" / "print_office"
    office_dir.mkdir(parents=True, exist_ok=True)
    docx_bundle = office_dir / "bundle.zip"
    xlsx_file = office_dir / "registers.xlsx"
    pack_file = office_dir / "office_pack.zip"
    docx_bundle.write_bytes(b"zip")
    xlsx_file.write_bytes(b"xlsx")
    pack_file.write_bytes(b"zip")

    observed: dict[str, str] = {}

    def fake_export_orders_docx_bundle(*, object_root: Path, classification: str) -> OfficeDocxExportResult:
        observed["classification"] = classification
        return OfficeDocxExportResult(output_dir=office_dir, bundle_path=docx_bundle, files_count=1)

    def fake_export_registers_xlsx(**_: object) -> OfficeXlsxExportResult:
        return OfficeXlsxExportResult(file_path=xlsx_file)

    def fake_build_office_pack_zip(**_: object) -> Path:
        return pack_file

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin._build_checklist", lambda root: [])
    monkeypatch.setattr("app.api.routes.arm_admin._collect_export_data", lambda db: ([], [], []))
    monkeypatch.setattr("app.api.routes.arm_admin.export_orders_docx_bundle", fake_export_orders_docx_bundle)
    monkeypatch.setattr("app.api.routes.arm_admin.export_registers_xlsx", fake_export_registers_xlsx)
    monkeypatch.setattr("app.api.routes.arm_admin.build_office_pack_zip", fake_build_office_pack_zip)

    with TestClient(app) as client:
        response = client.get("/arm/exports/office-pack", params={"classification": "employee-drafts"})

    assert response.status_code == 200
    assert observed["classification"] == "employee-drafts"


def test_arm_dashboard_contains_action_and_batch_controls(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)
    monkeypatch.setattr("app.api.routes.arm_admin.check_local_llm_available", lambda: (False, None))

    with TestClient(app) as client:
        response = client.get("/arm/dashboard")

    assert response.status_code == 200
    html = response.text
    assert 'id="treeBackBtn"' in html
    assert 'id="treeForwardBtn"' in html
    assert 'id="treeUpBtn"' in html
    assert 'id="treeRootBtn"' in html
    assert 'id="treeBreadcrumb"' in html
    assert 'id="taskActionLabel"' in html
    assert 'id="taskActionRunMissingBtn"' in html
    assert 'id="manualReviewList"' in html
    assert 'id="manualReviewMoveBtn"' in html
    assert 'id="batchGenerateMode"' in html
    assert 'id="batchGenerateAllFilteredBtn"' in html
    assert 'id="employeeChecklistToggleBtn"' in html
    assert 'id="backToTopBtn"' in html


def test_build_office_pack_zip_includes_employee_scans(tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    office_dir = root / "01_orders_and_appointments" / "print_office"
    docx_dir = office_dir / "docx"
    docx_dir.mkdir(parents=True, exist_ok=True)

    docx_bundle = office_dir / "пакет_приказов_docx.zip"
    docx_bundle.write_bytes(b"docx-bundle")
    xlsx_file = office_dir / "арм_реестры.xlsx"
    xlsx_file.write_bytes(b"xlsx")
    docx_file = docx_dir / "Приказ_07_Допуск_по_профессиям.docx"
    docx_file.write_bytes(b"docx")

    employee_scan = (
        root
        / "02_personnel"
        / "employees"
        / "007_исмагилов"
        / "01_identity_and_contract"
        / "20260311_PASSPORT_scan_from_arm_v01.jpg"
    )
    employee_scan.parent.mkdir(parents=True, exist_ok=True)
    employee_scan.write_bytes(b"scan")

    docx_result = OfficeDocxExportResult(output_dir=docx_dir, bundle_path=docx_bundle, files_count=1)
    xlsx_result = OfficeXlsxExportResult(file_path=xlsx_file)

    pack_path = build_office_pack_zip(root, docx_result, xlsx_result)

    assert pack_path.exists()

    import zipfile

    with zipfile.ZipFile(pack_path, "r") as archive:
        names = set(archive.namelist())

    assert "арм_реестры.xlsx" in names
    assert "пакет_приказов_docx.zip" in names
    assert "docx/Приказ_07_Допуск_по_профессиям.docx" in names
    assert (
        "scans/02_personnel/employees/007_исмагилов/01_identity_and_contract/"
        "20260311_PASSPORT_scan_from_arm_v01.jpg"
    ) in names


def test_arm_maintenance_reset_rebuild_preserves_employees(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    employee_root = _prepare_employee(root, "100_petrov_petr_petrovich")
    (employee_root / "employee_profile.txt").write_text(
        "employee_id: 100\n"
        "last_name: Петров\n"
        "first_name: Петр\n"
        "middle_name: Петрович\n"
        "position: прораб\n"
        "team: team_a\n",
        encoding="utf-8",
    )

    scan_file = root / "10_scan_inbox" / "temp_scan.jpg"
    scan_file.write_bytes(b"scan")
    draft_file = root / "01_orders_and_appointments" / "drafts_from_checklist" / "old.md"
    draft_file.parent.mkdir(parents=True, exist_ok=True)
    draft_file.write_text("old", encoding="utf-8")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.post(
            "/arm/maintenance/reset-rebuild",
            params={"regenerate_project_orders": 1, "overwrite_orders": 1},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert "Сотрудников сохранено: 1" in payload["message"]
    assert not scan_file.exists()
    assert not draft_file.exists()
    assert (employee_root / "employee_profile.txt").exists()

    rebuilt_dir = root / "01_orders_and_appointments" / "drafts_from_checklist"
    rebuilt_files = list(rebuilt_dir.glob("*.md"))
    assert rebuilt_files


def test_arm_speech_google_transcribe_success(monkeypatch) -> None:
    import io
    import wave

    import speech_recognition

    def fake_recognize_google(self, audio_data, language="ru-RU"):
        return "тестовая голосовая команда"

    monkeypatch.setattr(speech_recognition.Recognizer, "recognize_google", fake_recognize_google)

    buffer = io.BytesIO()
    with wave.open(buffer, "wb") as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(16000)
        wav.writeframes(b"\x00\x00" * 16000)

    with TestClient(app) as client:
        response = client.post(
            "/arm/speech/google-transcribe",
            files={"audio": ("voice.wav", buffer.getvalue(), "audio/wav")},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["provider"] == "google-webspeech"
    assert payload["text"] == "тестовая голосовая команда"


def test_arm_speech_google_transcribe_unknown_value(monkeypatch) -> None:
    import io
    import wave

    import speech_recognition

    def fake_recognize_google(self, audio_data, language="ru-RU"):
        raise speech_recognition.UnknownValueError()

    monkeypatch.setattr(speech_recognition.Recognizer, "recognize_google", fake_recognize_google)

    buffer = io.BytesIO()
    with wave.open(buffer, "wb") as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(16000)
        wav.writeframes(b"\x00\x00" * 8000)

    with TestClient(app) as client:
        response = client.post(
            "/arm/speech/google-transcribe",
            files={"audio": ("voice.wav", buffer.getvalue(), "audio/wav")},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is False
    assert "Речь не распознана" in payload["message"]


def test_arm_fs_file_write_persists_content_to_database(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = f"01_orders_and_appointments/{tmp_path.name}_editor_sync.md"
    first_content = "Первичная версия"
    second_content = "Обновленная версия"

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.post(
            "/arm/fs/file",
            json={"rel_path": rel_path, "content": first_content},
        )
        assert response.status_code == 200
        response = client.post(
            "/arm/fs/file",
            json={"rel_path": rel_path, "content": second_content},
        )

    assert response.status_code == 200
    assert "файл + БД" in response.json()["message"]
    assert (root / rel_path).read_text(encoding="utf-8") == second_content

    with SessionLocal() as db:
        snapshots = db.scalars(select(DocumentContent).where(DocumentContent.rel_path == rel_path)).all()

    assert len(snapshots) == 1
    assert snapshots[0].content == second_content


def test_arm_fs_print_preview_renders_text_content(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/20260312_ORDER_11_test_v01.md"
    target = root / rel_path
    target.write_text("## Заголовок\n\n| Колонка 1 | Колонка 2 |\n|---|---|\n| Строка 1 | Строка 2 |", encoding="utf-8")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/fs/print-preview",
            params={"rel_path": rel_path, "auto_print": 0},
        )

    assert response.status_code == 200
    assert 'class="doc-body"' in response.text
    assert "<table>" in response.text
    assert "Строка 1" in response.text


def test_arm_fs_print_preview_redirects_to_ready_pdf_for_order_markdown(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/20260312_ORDER_11_test_v01.md"
    target = root / rel_path
    target.write_text("# Приказ", encoding="utf-8")

    ready_pdf = root / "01_orders_and_appointments" / "print_pdf_ready" / "20260312_ORDER_11_test_v01.pdf"
    ready_pdf.write_bytes(b"%PDF-1.7\n")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/fs/print-preview",
            params={"rel_path": rel_path, "auto_print": 0},
            follow_redirects=False,
        )

    assert response.status_code == 307
    location = response.headers.get("location") or ""
    assert location.startswith("/arm/fs/view?rel_path=")
    assert "print_pdf_ready" in location
    assert "20260312_ORDER_11_test_v01.pdf" in location
    assert "&v=" in location


def test_arm_fs_print_preview_hides_employee_folder_paths(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/20260312_ORDER_01_test_v01.md"
    target = root / rel_path
    target.write_text(
        "## Приложение\n"
        "| Сотрудник | Скан (путь) |\n"
        "|---|---|\n"
        "| Иванов Иван | 02_personnel/employees/001_ivanov/04_attestation_and_certificates/ |\n",
        encoding="utf-8",
    )

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/fs/print-preview",
            params={"rel_path": rel_path, "auto_print": 0},
        )

    assert response.status_code == 200
    assert "02_personnel/employees" not in response.text
    assert "Приложение" not in response.text


def test_arm_fs_print_preview_strips_id_notes_and_appendix(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/20260312_ORDER_03_test_v01.md"
    target = root / rel_path
    target.write_text(
        "## Приказ\n"
        "| Направление | Основание |\n"
        "|---|---|\n"
        "| Охрана труда (новый состав ID 008-012) | Журнал ОТ (ручное ведение) |\n"
        "\n"
        "## Приложение: ссылки на сканы\n"
        "| ФИО | Путь |\n"
        "|---|---|\n"
        "| Иванов | 02_personnel/employees/001_ivanov/04_attestation_and_certificates/ |\n",
        encoding="utf-8",
    )

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/fs/print-preview",
            params={"rel_path": rel_path, "auto_print": 0},
        )

    assert response.status_code == 200
    assert "ID 008-012" not in response.text
    assert "ручное ведение" not in response.text
    assert "Приложение" not in response.text
    assert "02_personnel/employees" not in response.text


def test_arm_fs_print_preview_drops_id_table_column(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/20260312_ORDER_02_test_v01.md"
    target = root / rel_path
    target.write_text(
        "| ID | ФИО | Должность |\n"
        "|---|---|---|\n"
        "| 001 | Иванов Иван | Прораб |\n",
        encoding="utf-8",
    )

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/fs/print-preview",
            params={"rel_path": rel_path, "auto_print": 0},
        )

    assert response.status_code == 200
    assert ">ID<" not in response.text
    assert "Иванов Иван" in response.text


def test_arm_fs_print_preview_renders_image_inline(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "10_scan_inbox/sample_scan.jpg"
    target = root / rel_path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(b"\xff\xd8\xff\xd9")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get(
            "/arm/fs/print-preview",
            params={"rel_path": rel_path, "auto_print": 0},
        )

    assert response.status_code == 200
    assert '<img class="preview-image"' in response.text
    assert "/arm/fs/view?rel_path=" in response.text


def test_arm_fs_view_handles_cyrillic_pdf_filename(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/print_pdf_ready/20260312_ORDER_02_допуск_к_смр_v01.pdf"
    target = root / rel_path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(b"%PDF-1.7\n")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.get("/arm/fs/view", params={"rel_path": rel_path})

    assert response.status_code == 200
    disposition = response.headers.get("content-disposition") or ""
    assert "inline" in disposition.lower()
    assert "filename*=" in disposition.lower()


def test_arm_fs_print_returns_browser_preview_message(monkeypatch, tmp_path: Path) -> None:
    root = tmp_path / "object"
    _prepare_object_root(root)

    rel_path = "01_orders_and_appointments/20260312_ORDER_12_test_v01.md"
    target = root / rel_path
    target.write_text("content", encoding="utf-8")

    monkeypatch.setattr("app.api.routes.arm_admin.resolve_object_root", lambda: root)

    with TestClient(app) as client:
        response = client.post("/arm/fs/print", params={"rel_path": rel_path})

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert "/arm/fs/print-preview?rel_path=" in payload["message"]
    assert "auto_print=1" in payload["message"]
