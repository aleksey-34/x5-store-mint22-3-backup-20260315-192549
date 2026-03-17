"""
seed_db_from_filesystem.py — однократный импорт данных из файловой структуры в SQLite.

Читает:
  - docflow/objects/…/02_personnel/employees/**/employee_profile.txt  → документы типа «employee»
  - docflow/objects/…/01_orders_and_appointments/*.md                 → документы типа «order» / «permit»

Повторный запуск безопасен: пропускает уже существующие записи (по file_path).
"""

from __future__ import annotations

import sys
from pathlib import Path

# --- путь до корня проекта -----------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# импортируем только после добавления ROOT в sys.path
from app.db.session import SessionLocal, engine  # noqa: E402
from app.models import document as _doc_module  # noqa: E402
from app.models.document import Document  # noqa: E402

# убеждаемся, что таблицы созданы
from app.db.session import Base  # noqa: E402
Base.metadata.create_all(bind=engine)

# -------------------------------------------------------------------------

def _rel(path: Path) -> str:
    """Возвращает путь относительно корня проекта (прямые слэши)."""
    try:
        return path.relative_to(ROOT).as_posix()
    except ValueError:
        return path.as_posix()


def parse_profile(profile_path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    for line in profile_path.read_text(encoding="utf-8").splitlines():
        if ":" in line:
            key, _, value = line.partition(":")
            data[key.strip()] = value.strip()
    return data


def seed_employees(db, employees_root: Path) -> tuple[int, int]:
    added = skipped = 0
    for profile in sorted(employees_root.rglob("employee_profile.txt")):
        rel = _rel(profile)
        # Пропускаем дубли
        if db.query(Document).filter_by(file_path=rel).first():
            skipped += 1
            continue

        data = parse_profile(profile)
        full_name = " ".join(
            filter(None, [data.get("last_name"), data.get("first_name"), data.get("middle_name")])
        ) or profile.parent.name
        position = data.get("position", "")
        title = f"{full_name} — {position}" if position else full_name

        doc = Document(
            title=title,
            doc_type="employee",
            status="new",
            file_path=rel,
            notes=(
                f"ID: {data.get('employee_id', '')}  "
                f"Дата рождения: {data.get('birth_date', '')}  "
                f"Бригада: {data.get('team', '')}"
            ),
            fix_comment=None,
        )
        db.add(doc)
        added += 1

    db.flush()
    return added, skipped


def _guess_title_from_md(md_path: Path) -> str:
    """Читает первую строку H1 из .md файла, иначе использует имя файла."""
    try:
        for line in md_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("# "):
                return line[2:].strip()
    except Exception:
        pass
    return md_path.stem.replace("_", " ")


def seed_orders(db, orders_root: Path) -> tuple[int, int]:
    added = skipped = 0
    for md_file in sorted(orders_root.glob("*.md")):
        rel = _rel(md_file)
        if db.query(Document).filter_by(file_path=rel).first():
            skipped += 1
            continue

        stem_upper = md_file.stem.upper()
        if "PERMIT" in stem_upper:
            doc_type = "permit"
        elif "REGISTER" in stem_upper:
            doc_type = "order_register"
        else:
            doc_type = "order"

        doc = Document(
            title=_guess_title_from_md(md_file),
            doc_type=doc_type,
                status="approved",
            file_path=rel,
                fix_comment=None,
            notes=None,
        )
        db.add(doc)
        added += 1

    db.flush()
    return added, skipped


def main() -> None:
    from app.core.config import settings

    object_root = ROOT / settings.object_root
    employees_root = object_root / "02_personnel" / "employees"
    orders_root = object_root / "01_orders_and_appointments"

    if not employees_root.exists():
        print(f"[ERROR] employees dir not found: {employees_root}")
        sys.exit(1)
    if not orders_root.exists():
        print(f"[ERROR] orders dir not found: {orders_root}")
        sys.exit(1)

    db = SessionLocal()
    try:
        emp_add, emp_skip = seed_employees(db, employees_root)
        ord_add, ord_skip = seed_orders(db, orders_root)
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()

    print("=" * 55)
    print("  DB seed completed")
    print(f"  Employees : added={emp_add}  skipped={emp_skip}")
    print(f"  Orders    : added={ord_add}  skipped={ord_skip}")
    print("=" * 55)


if __name__ == "__main__":
    main()
