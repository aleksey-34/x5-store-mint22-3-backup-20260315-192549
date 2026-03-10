from __future__ import annotations

import argparse
import csv
import re
from dataclasses import dataclass
from pathlib import Path


OBJECT_STATIC_FOLDERS = [
    "00_incoming_requests",
    "01_orders_and_appointments",
    "02_personnel/00_registries",
    "02_personnel/employees",
    "03_hse_and_fire_safety/instructions",
    "03_hse_and_fire_safety/briefings",
    "03_hse_and_fire_safety/incidents_and_microtrauma",
    "03_hse_and_fire_safety/ppe_and_equipment_checks",
    "03_hse_and_fire_safety/permits",
    "04_journals/production",
    "04_journals/labor_safety",
    "05_execution_docs/ppr",
    "05_execution_docs/pprv_work_at_height",
    "05_execution_docs/admission_acts",
    "05_execution_docs/hidden_work_acts",
    "05_execution_docs/work_reports",
    "06_normative_base",
    "07_monthly_control",
    "08_outgoing_submissions",
    "09_archive",
    "09_archive/scan_bundles",
    "10_scan_inbox",
    "10_scan_inbox/manual_review",
]

EMPLOYEE_SUBFOLDERS = [
    "01_identity_and_contract",
    "02_admission_orders",
    "03_briefings_and_training",
    "04_attestation_and_certificates",
    "05_ppe_issue",
    "06_permits_and_work_admission",
    "07_medical_and_first_aid",
]


POSITION_NORMALIZATION = {
    "Монтажник": "Монтажник стальных и железобетонных конструкций",
    "Сварщик": "Электрогазосварщик",
}


@dataclass
class Employee:
    employee_id: str
    last_name: str
    first_name: str
    middle_name: str
    position: str
    grade: str
    birth_date: str
    team: str

    @property
    def folder_name(self) -> str:
        safe_last_name = slugify(self.last_name) or "employee"
        safe_id = slugify(self.employee_id) or "id"
        return f"{safe_id}_{safe_last_name}"


def slugify(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"[^a-zа-я0-9_\-]", "", value)
    return value


def normalize_position(value: str) -> str:
    clean_value = value.strip()
    return POSITION_NORMALIZATION.get(clean_value, clean_value)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create object document folders and employee dossiers."
    )
    parser.add_argument(
        "--output-root",
        default="docflow/objects",
        help="Root folder for generated object directories (default: docflow/objects)",
    )
    parser.add_argument("--object-code", required=True, help="Object code, e.g. X5-UFA-E2")
    parser.add_argument(
        "--object-name",
        required=True,
        help="Short object name, e.g. logistics_park",
    )
    parser.add_argument(
        "--employees-csv",
        default="docs/templates/personnel/employees_sample.csv",
        help=(
            "CSV with employees. Columns: "
            "employee_id,last_name,first_name,middle_name,position,grade,birth_date,team"
        ),
    )
    return parser.parse_args()


def load_employees(csv_path: Path) -> list[Employee]:
    if not csv_path.exists():
        raise FileNotFoundError(f"Employees CSV not found: {csv_path}")

    employees: list[Employee] = []
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not row.get("employee_id") or not row.get("last_name"):
                continue
            employees.append(
                Employee(
                    employee_id=row.get("employee_id", "").strip(),
                    last_name=row.get("last_name", "").strip(),
                    first_name=row.get("first_name", "").strip(),
                    middle_name=row.get("middle_name", "").strip(),
                    position=normalize_position(row.get("position", "")),
                    grade=row.get("grade", "").strip(),
                    birth_date=row.get("birth_date", "").strip(),
                    team=row.get("team", "").strip(),
                )
            )
    return employees


def ensure_object_folders(object_root: Path) -> None:
    object_root.mkdir(parents=True, exist_ok=True)
    for rel_path in OBJECT_STATIC_FOLDERS:
        (object_root / rel_path).mkdir(parents=True, exist_ok=True)


def create_employee_dossiers(object_root: Path, employees: list[Employee]) -> int:
    employees_root = object_root / "02_personnel" / "employees"
    count = 0

    for employee in employees:
        employee_root = employees_root / employee.folder_name
        employee_root.mkdir(parents=True, exist_ok=True)
        for folder in EMPLOYEE_SUBFOLDERS:
            (employee_root / folder).mkdir(parents=True, exist_ok=True)

        profile_file = employee_root / "employee_profile.txt"
        profile_content = (
            f"employee_id: {employee.employee_id}\n"
            f"last_name: {employee.last_name}\n"
            f"first_name: {employee.first_name}\n"
            f"middle_name: {employee.middle_name}\n"
            f"position: {employee.position}\n"
            f"grade: {employee.grade}\n"
            f"birth_date: {employee.birth_date}\n"
            f"team: {employee.team}\n"
        )
        profile_file.write_text(profile_content, encoding="utf-8")
        count += 1

    return count


def main() -> None:
    args = parse_args()

    object_code = slugify(args.object_code)
    object_name = slugify(args.object_name)
    object_folder_name = f"{object_code}_{object_name}"

    output_root = Path(args.output_root)
    object_root = output_root / object_folder_name

    employees = load_employees(Path(args.employees_csv))

    ensure_object_folders(object_root)
    created = create_employee_dossiers(object_root, employees)

    print(f"Object structure ready: {object_root}")
    print(f"Employee dossiers created: {created}")


if __name__ == "__main__":
    main()
