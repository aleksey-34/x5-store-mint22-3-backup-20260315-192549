from pathlib import Path

from app.services.scan_archiver import ingest_inbox, parse_scan_filename, resolve_destination_folder


def test_parse_scan_filename_for_passport() -> None:
    parsed = parse_scan_filename(Path("20260310__PASSPORT__ivanov_ii__001.pdf"))

    assert parsed.doc_type == "employee_passport"
    assert parsed.doc_code == "PASSPORT"
    assert parsed.employee_id == "001"


def test_parse_scan_filename_for_order() -> None:
    parsed = parse_scan_filename(Path("20260310__ORDER__appoint_hse_responsible.pdf"))

    assert parsed.doc_type == "order"
    assert parsed.doc_code == "ORDER"
    assert parsed.employee_id is None


def test_resolve_destination_for_passport(tmp_path: Path) -> None:
    object_root = tmp_path / "object"
    employee_folder = object_root / "02_personnel" / "employees" / "001_ivanov"
    (employee_folder / "01_identity_and_contract").mkdir(parents=True)

    parsed = parse_scan_filename(Path("20260310__PASSPORT__ivanov_ii__001.pdf"))
    destination = resolve_destination_folder(parsed=parsed, object_root=object_root)

    assert destination == employee_folder / "01_identity_and_contract"


def test_ingest_inbox_routes_hidden_work_act(tmp_path: Path) -> None:
    object_root = tmp_path / "object"
    inbox = object_root / "10_scan_inbox"
    inbox.mkdir(parents=True)

    scan_file = inbox / "20260310__AWR__foundation_axis_1_7.pdf"
    scan_file.write_text("fake-scan", encoding="utf-8")

    results = ingest_inbox(object_root=object_root, inbox_folder=inbox, db=None)

    assert len(results) == 1
    assert results[0].status == "archived"
    archived_file = object_root / "05_execution_docs" / "hidden_work_acts" / "20260310_AWR_foundation_axis_1_7_v01.pdf"
    assert archived_file.exists()


def test_ingest_inbox_moves_invalid_file_to_manual_review(tmp_path: Path) -> None:
    object_root = tmp_path / "object"
    inbox = object_root / "10_scan_inbox"
    inbox.mkdir(parents=True)

    bad_file = inbox / "bad_name.pdf"
    bad_file.write_text("bad", encoding="utf-8")

    results = ingest_inbox(object_root=object_root, inbox_folder=inbox, db=None)

    assert len(results) == 1
    assert results[0].status == "manual_review"
    assert (inbox / "manual_review" / "bad_name.pdf").exists()
    assert (inbox / "manual_review" / "review_log.csv").exists()
