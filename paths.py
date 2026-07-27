from __future__ import annotations

import os
import shutil
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
APP_DATA_DIR = Path(os.environ["LOCALAPPDATA"]) / "ResearchReportAutomation"
CONFIG_DIR = APP_DATA_DIR / "config"
CURRENT_COLLECTION_DIR = APP_DATA_DIR / "current_collection"
RUNS_DIR = APP_DATA_DIR / "runs"
RESOURCES_DIR = APP_DATA_DIR / "resources"


def ensure_app_dirs() -> None:
    for path in [APP_DATA_DIR, CONFIG_DIR, CURRENT_COLLECTION_DIR, RUNS_DIR, RESOURCES_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def migrate_legacy_data() -> None:
    legacy_data = APP_DIR / "data"
    if not legacy_data.exists():
        return
    ensure_app_dirs()
    for name in ["config", "current_collection", "runs"]:
        source = legacy_data / name
        target = APP_DATA_DIR / name
        if source.exists():
            copy_missing(source, target)


def copy_default_resource(filename: str, default_text: str | None = None, overwrite: bool = False) -> Path:
    ensure_app_dirs()
    target = RESOURCES_DIR / filename
    if target.exists() and not overwrite:
        return target
    source = APP_DIR / filename
    if source.exists():
        shutil.copy2(source, target)
    elif default_text is not None and (overwrite or not target.exists()):
        target.write_text(default_text, encoding="utf-8")
    return target


def copy_missing(source: Path, target: Path) -> None:
    if source.is_dir():
        target.mkdir(parents=True, exist_ok=True)
        for child in source.iterdir():
            copy_missing(child, target / child.name)
        return
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists():
        shutil.copy2(source, target)
