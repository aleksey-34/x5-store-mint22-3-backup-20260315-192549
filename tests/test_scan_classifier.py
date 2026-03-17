from app.services.scan_classifier import classify_scan_candidate, classify_scan_filename


def test_classify_scan_filename_detects_order() -> None:
    prediction = classify_scan_filename("20260310__PRIKAZ__appoint_hse.pdf")

    assert prediction.predicted_doc_type == "order"
    assert prediction.confidence > 0


def test_classify_scan_candidate_uses_ocr_text() -> None:
    prediction = classify_scan_candidate(
        filename="scan_001.pdf",
        ocr_text="Паспорт гражданина РФ и удостоверение личности",
    )

    assert prediction.predicted_doc_type == "employee_passport"
    assert prediction.source == "filename+ocr"
