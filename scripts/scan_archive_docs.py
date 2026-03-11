from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.db.init_db import init_db
from app.db.session import SessionLocal
from app.services.scan_classifier import classify_scan_candidate
from app.services.scan_archiver import (
    DEFAULT_OCR_LANG,
    DEFAULT_OCR_MAX_PDF_PAGES,
    SUPPORTED_SCAN_EXTENSIONS,
    create_period_archive,
    ingest_inbox,
)


def parse_date(value: str) -> datetime.date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Use date format YYYY-MM-DD") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Scan ingestion and archive automation")
    subparsers = parser.add_subparsers(dest="command", required=True)

    ingest = subparsers.add_parser("ingest", help="Ingest scan files from inbox")
    ingest.add_argument("--object-root", required=True, help="Object root folder path")
    ingest.add_argument(
        "--inbox",
        default=None,
        help="Inbox folder with fresh scans (default: <object-root>/10_scan_inbox)",
    )
    ingest.add_argument(
        "--no-db",
        action="store_true",
        help="Do not register processed files in the documents table",
    )
    ingest.add_argument(
        "--ocr",
        dest="enable_ocr",
        action="store_true",
        default=True,
        help="Enable OCR extraction for archived scans (default: enabled)",
    )
    ingest.add_argument(
        "--no-ocr",
        dest="enable_ocr",
        action="store_false",
        help="Disable OCR extraction",
    )
    ingest.add_argument(
        "--ocr-lang",
        default=DEFAULT_OCR_LANG,
        help="Tesseract language pack, e.g. rus+eng",
    )
    ingest.add_argument(
        "--tesseract-cmd",
        default=None,
        help="Optional full path to tesseract executable",
    )
    ingest.add_argument(
        "--max-pdf-pages",
        type=int,
        default=DEFAULT_OCR_MAX_PDF_PAGES,
        help="Max PDF pages to OCR per file",
    )

    bundle = subparsers.add_parser("bundle", help="Create period zip bundle")
    bundle.add_argument("--object-root", required=True, help="Object root folder path")
    bundle.add_argument("--from-date", required=True, type=parse_date, help="From date YYYY-MM-DD")
    bundle.add_argument("--to-date", required=True, type=parse_date, help="To date YYYY-MM-DD")
    bundle.add_argument("--output", default=None, help="Optional output zip path")

    classify = subparsers.add_parser("classify", help="Classify scan files for manual triage")
    classify.add_argument("--object-root", required=True, help="Object root folder path")
    classify.add_argument(
        "--path",
        default=None,
        help="Folder or file to classify (default: <object-root>/10_scan_inbox/manual_review)",
    )
    classify.add_argument(
        "--recursive",
        action="store_true",
        help="Recursively process nested folders",
    )

    return parser


def run_ingest(
    object_root: Path,
    inbox: Path,
    no_db: bool,
    enable_ocr: bool,
    ocr_lang: str,
    tesseract_cmd: str | None,
    max_pdf_pages: int,
) -> int:
    db = None
    if not no_db:
        init_db()
        db = SessionLocal()

    try:
        results = ingest_inbox(
            object_root=object_root,
            inbox_folder=inbox,
            db=db,
            enable_ocr=enable_ocr,
            ocr_lang=ocr_lang,
            tesseract_cmd=tesseract_cmd,
            max_pdf_pages=max_pdf_pages,
        )
        archived = sum(1 for item in results if item.status == "archived")
        manual_review = sum(1 for item in results if item.status == "manual_review")
        ocr_ok = sum(1 for item in results if item.ocr_status == "ok")
        ocr_unavailable = sum(1 for item in results if item.ocr_status == "unavailable")

        print(f"Inbox processed: {len(results)}")
        print(f"Archived: {archived}")
        print(f"Manual review: {manual_review}")
        print(f"OCR ok: {ocr_ok}")
        print(f"OCR unavailable: {ocr_unavailable}")

        for item in results:
            destination = item.destination or "-"
            ocr_info = item.ocr_status
            if item.ocr_text_path:
                ocr_info = f"{ocr_info} ({item.ocr_text_path})"
            ml_hint = ""
            if item.suggested_doc_type and item.suggested_confidence is not None:
                ml_hint = f"; ML={item.suggested_doc_type}({item.suggested_confidence:.2f})"
            print(f"[{item.status}] {item.source_name} -> {destination}; {item.message}; OCR={ocr_info}{ml_hint}")

        return 0
    finally:
        if db is not None:
            db.close()


def run_bundle(
    object_root: Path,
    from_date: datetime.date,
    to_date: datetime.date,
    output: str | None,
) -> int:
    output_path = Path(output) if output else None
    result = create_period_archive(
        object_root=object_root,
        from_date=from_date,
        to_date=to_date,
        output_zip=output_path,
    )

    print(f"Archive ready: {result.zip_path}")
    print(f"Included files: {result.included_files}")
    return 0


def _iter_files_for_classification(path: Path, recursive: bool) -> list[Path]:
    if path.is_file():
        if path.suffix.lower() in SUPPORTED_SCAN_EXTENSIONS:
            return [path]
        return []

    iterator = path.rglob("*") if recursive else path.iterdir()
    files = [item for item in iterator if item.is_file() and item.suffix.lower() in SUPPORTED_SCAN_EXTENSIONS]
    return sorted(files)


def _read_sidecar_ocr(file_path: Path) -> str | None:
    sidecar = file_path.with_suffix(f"{file_path.suffix}.ocr.txt")
    if not sidecar.exists():
        return None
    try:
        return sidecar.read_text(encoding="utf-8")
    except Exception:  # noqa: BLE001
        return None


def run_classify(object_root: Path, path_arg: str | None, recursive: bool) -> int:
    path = Path(path_arg) if path_arg else (object_root / "10_scan_inbox" / "manual_review")
    if not path.exists():
        print(f"Path does not exist: {path}")
        return 2

    files = _iter_files_for_classification(path=path, recursive=recursive)
    print(f"Files to classify: {len(files)}")

    for file in files:
        ocr_text = _read_sidecar_ocr(file)
        prediction = classify_scan_candidate(filename=file.name, ocr_text=ocr_text)
        hits = ",".join(prediction.matched_keywords) if prediction.matched_keywords else "-"
        print(
            f"{file.name}; predicted={prediction.predicted_doc_type}; "
            f"confidence={prediction.confidence:.2f}; source={prediction.source}; hits={hits}"
        )

    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    object_root = Path(args.object_root)
    if not object_root.exists():
        parser.error(f"Object root does not exist: {object_root}")

    if args.command == "ingest":
        inbox = Path(args.inbox) if args.inbox else object_root / "10_scan_inbox"
        return run_ingest(
            object_root=object_root,
            inbox=inbox,
            no_db=args.no_db,
            enable_ocr=args.enable_ocr,
            ocr_lang=args.ocr_lang,
            tesseract_cmd=args.tesseract_cmd,
            max_pdf_pages=args.max_pdf_pages,
        )

    if args.command == "bundle":
        return run_bundle(
            object_root=object_root,
            from_date=args.from_date,
            to_date=args.to_date,
            output=args.output,
        )

    if args.command == "classify":
        return run_classify(
            object_root=object_root,
            path_arg=args.path,
            recursive=args.recursive,
        )

    parser.error(f"Unsupported command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
