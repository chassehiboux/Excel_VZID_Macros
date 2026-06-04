from __future__ import annotations

import os
import shutil
import sys
import argparse
from pathlib import Path
from tkinter import Tk, messagebox

import pythoncom
import win32com.client


APP_NAME = "VZID"
LOADER_FILE = "LoaderVZID.xlam"
MAIN_FILE = "MainVZID.xlam"
CONFIG_TEMPLATE = "config.template.json"
INSTALL_BASE_DIR: Path | None = None
NO_UI = False


def base_dir() -> Path:
    if INSTALL_BASE_DIR is not None:
        return INSTALL_BASE_DIR

    root = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if not root:
        root = str(Path.home())
    return Path(root) / APP_NAME


def loader_dir() -> Path:
    return base_dir() / "loader"


def current_dir() -> Path:
    return base_dir() / "versions" / "current"


def pending_dir() -> Path:
    return base_dir() / "versions" / "pending"


def config_dir() -> Path:
    return base_dir() / "config"


def logs_dir() -> Path:
    return base_dir() / "logs"


def config_path() -> Path:
    return config_dir() / "config.json"


def loader_path() -> Path:
    return loader_dir() / LOADER_FILE


def main_path() -> Path:
    return current_dir() / MAIN_FILE


def bundled_path(relative_path: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path

    repo_root = Path(__file__).resolve().parents[1]
    direct_path = repo_root / relative_path
    if direct_path.exists():
        return direct_path

    build_path = repo_root / "build" / Path(relative_path).name
    if build_path.exists():
        return build_path

    return direct_path


def ensure_dirs() -> None:
    for path in (loader_dir(), current_dir(), pending_dir(), config_dir(), logs_dir()):
        path.mkdir(parents=True, exist_ok=True)


def copy_assets() -> None:
    shutil.copy2(bundled_path(LOADER_FILE), loader_path())
    shutil.copy2(bundled_path(MAIN_FILE), main_path())
    if not config_path().exists():
        shutil.copy2(bundled_path(f"config/{CONFIG_TEMPLATE}"), config_path())


def register_loader_addin() -> None:
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        addin_ref = None
        target_path = str(loader_path()).lower()
        for index in range(1, excel.AddIns.Count + 1):
            candidate = excel.AddIns.Item(index)
            if str(candidate.FullName).lower() == target_path:
                addin_ref = candidate
                break

        if addin_ref is None:
            addin_ref = excel.AddIns.Add(str(loader_path()), False)

        addin_ref.Installed = True
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def show_info(title: str, text: str) -> None:
    if NO_UI:
        print(f"{title}: {text}")
        return

    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, text)
    root.destroy()


def show_error(title: str, text: str) -> None:
    if NO_UI:
        print(f"{title}: {text}")
        return

    root = Tk()
    root.withdraw()
    messagebox.showerror(title, text)
    root.destroy()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target-root")
    parser.add_argument("--skip-register", action="store_true")
    parser.add_argument("--no-ui", action="store_true")
    return parser.parse_args()


def main() -> int:
    global INSTALL_BASE_DIR
    global NO_UI

    args = parse_args()
    NO_UI = args.no_ui
    if args.target_root:
        INSTALL_BASE_DIR = Path(args.target_root).resolve()

    try:
        ensure_dirs()
        copy_assets()
        if not args.skip_register:
            register_loader_addin()

        show_info(
            "VZID Setup",
            "Надстройка установлена для текущего пользователя.\n\n"
            "Если Excel сейчас открыт, полностью закройте все его окна и откройте Excel заново.",
        )
        return 0
    except Exception as exc:  # pragma: no cover
        show_error("VZID Setup", f"Установка завершилась ошибкой:\n{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
