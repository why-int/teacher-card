from __future__ import annotations

import sys
from pathlib import Path

APP_TITLE = "Teacher Card Builder"
WINDOW_TITLE = APP_TITLE
TEMPLATE_FILE_NAMES = ("Пример.xlsx", "Шаблон и пример.xlsm")


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def bundle_root() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parents[1]


def app_runtime_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resource_path(relative_path: str) -> Path:
    return bundle_root() / relative_path


def default_save_dir() -> Path:
    documents = Path.home() / "Documents"
    if documents.exists():
        return documents
    return Path.home()


def find_template_file() -> Path | None:
    candidates: list[Path] = []

    for name in TEMPLATE_FILE_NAMES:
        candidates.append(app_runtime_dir() / name)
        candidates.append(Path.cwd() / name)
        candidates.append(resource_path(f"templates/{name}"))

    for path in candidates:
        if path.is_file():
            return path
    return None
