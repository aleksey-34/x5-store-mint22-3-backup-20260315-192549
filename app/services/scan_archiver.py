from __future__ import annotations

import csv
import hashlib
import re
import shutil
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from sqlalchemy.orm import Session

from app.models.document import Document
from app.services.scan_classifier import classify_scan_filename

try:
    import pypdfium2 as pdfium
except Exception:  # pragma: no cover - optional runtime dependency
    pdfium = None

try:
    import pytesseract
except Exception:  # pragma: no cover - optional runtime dependency
    pytesseract = None

try:
    from PIL import Image
except Exception:  # pragma: no cover - optional runtime dependency
    Image = None

SUPPORTED_SCAN_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff"}

DOC_TYPE_ALIASES = {
    "AWR": "hidden_work_act",
    "HIDDEN_WORK_ACT": "hidden_work_act",
    "ACT_HIDDEN_WORK": "hidden_work_act",
    "ACT_OF_HIDDEN_WORKS": "hidden_work_act",
    "АКТ_СКРЫТЫХ_РАБОТ": "hidden_work_act",
    "ПАСПОРТ": "employee_passport",
    "PASSPORT": "employee_passport",
    "ORDER": "order",
    "ПРИКАЗ": "order",
}

DOC_TYPE_TO_CODE = {
    "hidden_work_act": "AWR",
    "employee_passport": "PASSPORT",
    "order": "ORDER",
}

DOC_TYPE_TO_FOLDER = {
    "hidden_work_act": Path("05_execution_docs/hidden_work_acts"),
    "order": Path("01_orders_and_appointments"),
}

DEFAULT_OCR_LANG = "rus+eng"
DEFAULT_OCR_MAX_PDF_PAGES = 4


@dataclass
class ParsedScan:
    source_file: Path
    doc_date: date
    doc_date_raw: str
    doc_type: str
    doc_code: str
    subject_raw: str
    employee_id: str | None


@dataclass
class IngestResult:
    source_name: str
    status: str
    message: str
    destination: str | None = None
    document_id: int | None = None
    ocr_status: str = "disabled"
    ocr_text_path: str | None = None
    suggested_doc_type: str | None = None
    suggested_confidence: float | None = None


@dataclass
class OCRResult:
    status: str
    text: str = ""
    message: str = ""


@dataclass
class ArchiveBundleResult:
    zip_path: Path
    included_files: int


def normalize_token(value: str) -> str:
    token = value.strip().upper()
    token = re.sub(r"[\s\-]+", "_", token)
    token = re.sub(r"[^0-9A-ZА-Я_]", "", token)
    return token


def slugify_filename_part(value: str) -> str:
    token = value.strip().lower()
    token = re.sub(r"[\s\-]+", "_", token)
    token = re.sub(r"[^0-9a-zа-я_]", "", token)
    token = token.strip("_")
    return token or "document"


def parse_scan_filename(file_path: Path) -> ParsedScan:
    parts = file_path.stem.split("__")
    if len(parts) < 3:
        raise ValueError(
            "Invalid filename format. Expected: YYYYMMDD__DOC_TYPE__SUBJECT__[EMPLOYEE_ID].ext"
        )

    date_raw = parts[0].strip()
    try:
        parsed_date = datetime.strptime(date_raw, "%Y%m%d").date()
    except ValueError as exc:
        raise ValueError("Invalid date in filename, expected YYYYMMDD") from exc

    type_token = normalize_token(parts[1])
    doc_type = DOC_TYPE_ALIASES.get(type_token)
    if doc_type is None:
        raise ValueError(
            "Unsupported DOC_TYPE. Supported values include AWR, PASSPORT, ORDER"
        )

    subject_parts = parts[2:]
    employee_id: str | None = None
    if len(parts) >= 4:
        employee_id = parts[-1].strip()
        subject_parts = parts[2:-1]

    subject_raw = "__".join(subject_parts).strip()
    if not subject_raw:
        raise ValueError("Subject segment is empty")

    return ParsedScan(
        source_file=file_path,
        doc_date=parsed_date,
        doc_date_raw=date_raw,
        doc_type=doc_type,
        doc_code=DOC_TYPE_TO_CODE[doc_type],
        subject_raw=subject_raw,
        employee_id=employee_id,
    )


def resolve_destination_folder(parsed: ParsedScan, object_root: Path) -> Path:
    if parsed.doc_type in DOC_TYPE_TO_FOLDER:
        return object_root / DOC_TYPE_TO_FOLDER[parsed.doc_type]

    if parsed.doc_type == "employee_passport":
        if not parsed.employee_id:
            raise ValueError("EMPLOYEE_ID is required for PASSPORT files")

        employees_root = object_root / "02_personnel" / "employees"
        employee_token = slugify_filename_part(parsed.employee_id)
        matches = sorted(employees_root.glob(f"{employee_token}_*"))

        if not matches:
            raise ValueError(
                f"Employee folder for EMPLOYEE_ID '{parsed.employee_id}' was not found"
            )
        if len(matches) > 1:
            raise ValueError(
                f"Multiple employee folders found for EMPLOYEE_ID '{parsed.employee_id}'"
            )

        return matches[0] / "01_identity_and_contract"

    raise ValueError(f"No destination mapping for doc_type '{parsed.doc_type}'")


def next_revision(destination_folder: Path, file_base_name: str, extension: str) -> int:
    pattern = re.compile(rf"^{re.escape(file_base_name)}_v(\d+){re.escape(extension)}$", re.IGNORECASE)
    max_revision = 0

    for existing in destination_folder.iterdir():
        if not existing.is_file():
            continue
        match = pattern.match(existing.name)
        if match:
            max_revision = max(max_revision, int(match.group(1)))

    return max_revision + 1


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def append_review_log(log_file: Path, source_name: str, reason: str) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    file_exists = log_file.exists()

    with log_file.open("a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["timestamp", "source_name", "reason"])
        if not file_exists:
            writer.writeheader()
        writer.writerow(
            {
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "source_name": source_name,
                "reason": reason,
            }
        )


def _configure_tesseract_cmd(tesseract_cmd: str | None) -> None:
    if not tesseract_cmd or pytesseract is None:
        return
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd


def _ocr_image_file(path: Path, lang: str) -> OCRResult:
    if pytesseract is None or Image is None:
        return OCRResult(status="unavailable", message="Install pytesseract and Pillow")

    try:
        with Image.open(path) as img:
            text = pytesseract.image_to_string(img, lang=lang).strip()
        if text:
            return OCRResult(status="ok", text=text)
        return OCRResult(status="empty", message="OCR returned empty text")
    except Exception as exc:  # noqa: BLE001
        return OCRResult(status="error", message=str(exc))


def _ocr_pdf_file(path: Path, lang: str, max_pages: int) -> OCRResult:
    if pytesseract is None or pdfium is None:
        return OCRResult(status="unavailable", message="Install pytesseract and pypdfium2")

    try:
        pdf = pdfium.PdfDocument(str(path))
        page_count = len(pdf)
        selected_pages = min(page_count, max_pages)

        texts: list[str] = []
        for page_idx in range(selected_pages):
            page = pdf.get_page(page_idx)
            image = page.render(scale=300 / 72).to_pil()
            texts.append(pytesseract.image_to_string(image, lang=lang).strip())
            page.close()

        full_text = "\n\n".join(part for part in texts if part)
        if full_text.strip():
            return OCRResult(status="ok", text=full_text)
        return OCRResult(status="empty", message="OCR returned empty text")
    except Exception as exc:  # noqa: BLE001
        return OCRResult(status="error", message=str(exc))


def run_ocr_for_file(
    archived_file: Path,
    ocr_lang: str,
    tesseract_cmd: str | None,
    max_pdf_pages: int,
) -> OCRResult:
    _configure_tesseract_cmd(tesseract_cmd)

    suffix = archived_file.suffix.lower()
    if suffix == ".pdf":
        return _ocr_pdf_file(archived_file, lang=ocr_lang, max_pages=max_pdf_pages)
    if suffix in {".jpg", ".jpeg", ".png", ".tif", ".tiff"}:
        return _ocr_image_file(archived_file, lang=ocr_lang)

    return OCRResult(status="unavailable", message=f"OCR unsupported for extension '{suffix}'")


def write_ocr_sidecar(archived_file: Path, text: str) -> Path:
    sidecar = archived_file.with_suffix(f"{archived_file.suffix}.ocr.txt")
    sidecar.write_text(text, encoding="utf-8")
    return sidecar


def register_document_record(
    db: Session,
    parsed: ParsedScan,
    relative_path: str,
    source_name: str,
    checksum: str,
    ocr_status: str,
    ocr_text_path: str | None,
    ocr_message: str,
) -> int:
    note_parts = [
        f"source_scan={source_name}",
        f"sha256={checksum}",
        f"ocr_status={ocr_status}",
    ]
    if ocr_text_path:
        note_parts.append(f"ocr_text_path={ocr_text_path}")
    if ocr_message:
        note_parts.append(f"ocr_message={ocr_message}")

    item = Document(
        title=parsed.subject_raw.replace("_", " "),
        doc_type=parsed.doc_type,
        status="archived",
        file_path=relative_path,
        notes="; ".join(note_parts),
    )
    db.add(item)
    db.commit()
    db.refresh(item)
    return item.id


def ingest_scan_file(
    source_file: Path,
    object_root: Path,
    manual_review_folder: Path,
    db: Session | None,
    enable_ocr: bool = True,
    ocr_lang: str = DEFAULT_OCR_LANG,
    tesseract_cmd: str | None = None,
    max_pdf_pages: int = DEFAULT_OCR_MAX_PDF_PAGES,
) -> IngestResult:
    if source_file.suffix.lower() not in SUPPORTED_SCAN_EXTENSIONS:
        message = f"Unsupported extension '{source_file.suffix}'"
        target = ensure_unique_path(manual_review_folder / source_file.name)
        shutil.move(str(source_file), str(target))
        append_review_log(manual_review_folder / "review_log.csv", source_file.name, message)
        return IngestResult(source_name=source_file.name, status="manual_review", message=message)

    try:
        parsed = parse_scan_filename(source_file)
        destination_folder = resolve_destination_folder(parsed, object_root)
        destination_folder.mkdir(parents=True, exist_ok=True)

        subject_slug = slugify_filename_part(parsed.subject_raw)
        extension = source_file.suffix.lower()
        base_name = f"{parsed.doc_date_raw}_{parsed.doc_code}_{subject_slug}"
        revision = next_revision(destination_folder, base_name, extension)
        destination_name = f"{base_name}_v{revision:02d}{extension}"

        destination_file = ensure_unique_path(destination_folder / destination_name)
        shutil.move(str(source_file), str(destination_file))

        checksum = sha256_file(destination_file)
        relative_path = str(destination_file.relative_to(object_root)).replace("\\", "/")

        ocr_status = "disabled"
        ocr_message = ""
        ocr_text_path: str | None = None

        if enable_ocr:
            ocr_result = run_ocr_for_file(
                archived_file=destination_file,
                ocr_lang=ocr_lang,
                tesseract_cmd=tesseract_cmd,
                max_pdf_pages=max_pdf_pages,
            )
            ocr_status = "warning" if ocr_result.status == "error" else ocr_result.status
            ocr_message = ocr_result.message

            if ocr_result.status == "ok" and ocr_result.text:
                sidecar = write_ocr_sidecar(destination_file, ocr_result.text)
                ocr_text_path = str(sidecar.relative_to(object_root)).replace("\\", "/")

        document_id: int | None = None
        if db is not None:
            document_id = register_document_record(
                db=db,
                parsed=parsed,
                relative_path=relative_path,
                source_name=source_file.name,
                checksum=checksum,
                ocr_status=ocr_status,
                ocr_text_path=ocr_text_path,
                ocr_message=ocr_message,
            )

        return IngestResult(
            source_name=source_file.name,
            status="archived",
            message="ok",
            destination=relative_path,
            document_id=document_id,
            ocr_status=ocr_status,
            ocr_text_path=ocr_text_path,
        )
    except Exception as exc:  # noqa: BLE001
        prediction = classify_scan_filename(source_file.name)
        suggestion_suffix = ""
        suggested_doc_type: str | None = None
        suggested_confidence: float | None = None
        if prediction.predicted_doc_type != "unknown":
            suggested_doc_type = prediction.predicted_doc_type
            suggested_confidence = prediction.confidence
            suggestion_suffix = (
                f"; suggested_doc_type={prediction.predicted_doc_type}; "
                f"confidence={prediction.confidence:.2f}"
            )

        final_message = f"{exc}{suggestion_suffix}"
        target = ensure_unique_path(manual_review_folder / source_file.name)
        shutil.move(str(source_file), str(target))
        append_review_log(manual_review_folder / "review_log.csv", source_file.name, final_message)
        return IngestResult(
            source_name=source_file.name,
            status="manual_review",
            message=final_message,
            suggested_doc_type=suggested_doc_type,
            suggested_confidence=suggested_confidence,
        )


def ingest_inbox(
    object_root: Path,
    inbox_folder: Path,
    db: Session | None,
    enable_ocr: bool = True,
    ocr_lang: str = DEFAULT_OCR_LANG,
    tesseract_cmd: str | None = None,
    max_pdf_pages: int = DEFAULT_OCR_MAX_PDF_PAGES,
) -> list[IngestResult]:
    inbox_folder.mkdir(parents=True, exist_ok=True)
    manual_review_folder = inbox_folder / "manual_review"
    manual_review_folder.mkdir(parents=True, exist_ok=True)

    results: list[IngestResult] = []
    for source_file in sorted(inbox_folder.iterdir()):
        if source_file.is_dir():
            continue
        result = ingest_scan_file(
            source_file=source_file,
            object_root=object_root,
            manual_review_folder=manual_review_folder,
            db=db,
            enable_ocr=enable_ocr,
            ocr_lang=ocr_lang,
            tesseract_cmd=tesseract_cmd,
            max_pdf_pages=max_pdf_pages,
        )
        results.append(result)

    return results


def parse_document_date_from_filename(path: Path) -> date | None:
    match = re.match(r"^(\d{8})_", path.name)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y%m%d").date()
    except ValueError:
        return None


def archive_candidates(object_root: Path) -> list[Path]:
    roots = [
        object_root / "01_orders_and_appointments",
        object_root / "05_execution_docs" / "hidden_work_acts",
        object_root / "02_personnel" / "employees",
    ]

    files: list[Path] = []
    for root in roots:
        if not root.exists():
            continue
        for file in root.rglob("*"):
            if file.is_file() and file.suffix.lower() in SUPPORTED_SCAN_EXTENSIONS:
                files.append(file)
    return files


def create_period_archive(
    object_root: Path,
    from_date: date,
    to_date: date,
    output_zip: Path | None = None,
) -> ArchiveBundleResult:
    if from_date > to_date:
        raise ValueError("from_date must be less or equal to to_date")

    if output_zip is None:
        archive_root = object_root / "09_archive" / "scan_bundles"
        archive_root.mkdir(parents=True, exist_ok=True)
        output_zip = archive_root / f"scan_bundle_{from_date:%Y%m%d}_{to_date:%Y%m%d}.zip"

    selected_files: list[Path] = []
    for file in archive_candidates(object_root):
        file_date = parse_document_date_from_filename(file)
        if file_date is None:
            continue
        if from_date <= file_date <= to_date:
            selected_files.append(file)

    output_zip.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_zip, mode="w", compression=ZIP_DEFLATED) as archive:
        for file in selected_files:
            archive.write(file, arcname=str(file.relative_to(object_root)).replace("\\", "/"))

    return ArchiveBundleResult(zip_path=output_zip, included_files=len(selected_files))
