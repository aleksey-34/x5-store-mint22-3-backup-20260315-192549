from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path

FORMAT_GUIDS = {
    "jpg": "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}",
    "png": "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}",
    "tif": "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}",
}

DOC_TYPES = {"AWR", "PASSPORT", "ORDER", "INVOICE", "UPD", "TTN", "ACT", "OTHER"}


def _slug(value: str) -> str:
    token = "_".join(value.strip().split()).lower()
    clean = []
    for ch in token:
        if ch.isalnum() or ch == "_":
            clean.append(ch)
    result = "".join(clean).strip("_")
    return result or "document"


def _import_win32com_client():
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is not installed. Install it with: pip install pywin32"
        ) from exc
    return win32com.client


def _wia_prop(obj, name: str, default: str = "") -> str:
    try:
        for prop in obj.Properties:
            if str(prop.Name).lower() == name.lower():
                return str(prop.Value)
    except Exception:  # noqa: BLE001
        return default
    return default


def _set_wia_property(item, property_id: int, value: int) -> None:
    try:
        for prop in item.Properties:
            if int(prop.PropertyID) == property_id:
                prop.Value = value
                return
    except Exception:
        return


def list_scanners() -> list[tuple[int, str, str]]:
    win32 = _import_win32com_client()
    manager = win32.Dispatch("WIA.DeviceManager")

    devices: list[tuple[int, str, str]] = []
    count = int(manager.DeviceInfos.Count)
    for idx in range(1, count + 1):
        info = manager.DeviceInfos.Item(idx)
        name = _wia_prop(info, "Name", default=f"Device-{idx}")
        device_id = _wia_prop(info, "Device ID", default="") or _wia_prop(info, "DeviceID", default="")
        devices.append((idx, name, device_id))

    return devices


def capture_scan(
    output_file: Path,
    device_index: int,
    image_format: str,
    dpi: int,
    grayscale: bool,
) -> Path:
    win32 = _import_win32com_client()

    manager = win32.Dispatch("WIA.DeviceManager")
    total = int(manager.DeviceInfos.Count)
    if device_index < 1 or device_index > total:
        raise ValueError(f"device_index must be in range 1..{total}")

    info = manager.DeviceInfos.Item(device_index)
    device = info.Connect()
    item = device.Items.Item(1)

    _set_wia_property(item, 6147, dpi)  # Horizontal Resolution
    _set_wia_property(item, 6148, dpi)  # Vertical Resolution
    _set_wia_property(item, 6146, 2 if grayscale else 1)  # Current Intent

    common_dialog = win32.Dispatch("WIA.CommonDialog")
    image = common_dialog.ShowTransfer(item, FORMAT_GUIDS[image_format], False)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    image.SaveFile(str(output_file))
    return output_file


def detect_windows_scanner_stack() -> str:
    commands = [
        (
            "WIA service",
            "Get-Service -Name stisvc | Select-Object Name,Status,StartType | Format-Table -AutoSize | Out-String",
        ),
        (
            "Image devices",
            "Get-PnpDevice -Class Image | Select-Object Status,Class,FriendlyName,InstanceId | Format-Table -AutoSize | Out-String",
        ),
    ]

    chunks: list[str] = []
    for title, ps_cmd in commands:
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
        body = completed.stdout.strip() or completed.stderr.strip() or "No output"
        chunks.append(f"[{title}]\n{body}")

    return "\n\n".join(chunks)


def build_inbox_filename(
    doc_type: str,
    subject: str,
    employee_id: str | None,
    scan_date: str,
    image_format: str,
) -> str:
    normalized_type = doc_type.upper().strip()
    if normalized_type not in DOC_TYPES:
        allowed = ", ".join(sorted(DOC_TYPES))
        raise ValueError(f"Unsupported doc_type '{doc_type}'. Allowed: {allowed}")

    subject_slug = _slug(subject)
    ext = image_format.lower()

    if normalized_type == "PASSPORT":
        if not employee_id:
            raise ValueError("employee_id is required for PASSPORT")
        return f"{scan_date}__{normalized_type}__{subject_slug}__{employee_id}.{ext}"

    return f"{scan_date}__{normalized_type}__{subject_slug}.{ext}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="WIA scanner control for construction docflow")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("list", help="List available WIA scanners")
    subparsers.add_parser("diagnose", help="Show Windows scanner stack status")

    scan = subparsers.add_parser("scan-to-inbox", help="Capture one scan and save to object inbox")
    scan.add_argument("--object-root", required=True, help="Object root path")
    scan.add_argument("--doc-type", required=True, choices=sorted(DOC_TYPES), help="Document type")
    scan.add_argument("--subject", required=True, help="Short subject for file naming")
    scan.add_argument("--employee-id", default=None, help="Employee ID (required for PASSPORT)")
    scan.add_argument("--device-index", type=int, default=1, help="WIA device index from list command")
    scan.add_argument("--format", dest="image_format", choices=sorted(FORMAT_GUIDS), default="jpg")
    scan.add_argument("--dpi", type=int, default=300, help="Scan resolution in DPI")
    scan.add_argument("--grayscale", action="store_true", help="Capture in grayscale")
    scan.add_argument(
        "--scan-date",
        default=datetime.now().strftime("%Y%m%d"),
        help="Date token for filename (YYYYMMDD)",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.command == "diagnose":
        print(detect_windows_scanner_stack())
        return 0

    if args.command == "list":
        scanners = list_scanners()
        if not scanners:
            print("No WIA scanners found")
            return 1

        print("Available WIA scanners:")
        for idx, name, device_id in scanners:
            print(f"  [{idx}] {name}")
            if device_id:
                print(f"      {device_id}")
        return 0

    if args.command == "scan-to-inbox":
        object_root = Path(args.object_root)
        if not object_root.exists():
            raise SystemExit(f"Object root not found: {object_root}")

        inbox = object_root / "10_scan_inbox"
        filename = build_inbox_filename(
            doc_type=args.doc_type,
            subject=args.subject,
            employee_id=args.employee_id,
            scan_date=args.scan_date,
            image_format=args.image_format,
        )
        output_file = inbox / filename

        saved = capture_scan(
            output_file=output_file,
            device_index=args.device_index,
            image_format=args.image_format,
            dpi=args.dpi,
            grayscale=args.grayscale,
        )
        print(f"Scan saved: {saved}")
        return 0

    raise SystemExit(f"Unsupported command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
