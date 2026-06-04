from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import winreg
from pathlib import Path
from tkinter import Tk, messagebox

import pythoncom
import win32com.client


APP_NAME = "VZID"
LEGACY_LOADER_FILE = "LoaderVZID.xlam"
MAIN_FILE = "MainVZID.xlam"
UPDATER_FILE = "updater.exe"
CONFIG_TEMPLATE = "config.template.json"
STARTUP_FILE_NAMES = {LEGACY_LOADER_FILE.lower(), MAIN_FILE.lower()}
INSTALL_BASE_DIR: Path | None = None
NO_UI = False


def base_dir() -> Path:
    if INSTALL_BASE_DIR is not None:
        return INSTALL_BASE_DIR

    return local_cache_root() / APP_NAME


def appdata_root() -> Path:
    if INSTALL_BASE_DIR is not None:
        return INSTALL_BASE_DIR.parent

    root = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
    if not root:
        root = str(Path.home())
    return Path(root)


def localappdata_root() -> Path:
    if INSTALL_BASE_DIR is not None:
        return INSTALL_BASE_DIR.parent / "LegacyLocalAppData"

    root = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if not root:
        root = str(Path.home())
    return Path(root)


def excel_root() -> Path:
    return appdata_root() / "Microsoft" / "Excel"


def local_cache_root() -> Path:
    return excel_root() / "LocalCache"


def xlstart_dir() -> Path:
    return excel_root() / "XLSTART"


def addin_dir() -> Path:
    return base_dir() / "addin"


def updater_dir() -> Path:
    return base_dir() / "updater"


def updates_dir() -> Path:
    return base_dir() / "updates"


def backup_dir() -> Path:
    return base_dir() / "backup"


def legacy_loader_dir() -> Path:
    return base_dir() / "loader"


def legacy_versions_dir() -> Path:
    return base_dir() / "versions"


def legacy_local_vzid_dir() -> Path:
    return localappdata_root() / APP_NAME


def config_dir() -> Path:
    return base_dir() / "config"


def logs_dir() -> Path:
    return base_dir() / "logs"


def config_path() -> Path:
    return config_dir() / "config.json"


def config_template_path() -> Path:
    return bundled_path(f"config/{CONFIG_TEMPLATE}")


def main_path() -> Path:
    return addin_dir() / MAIN_FILE


def updater_path() -> Path:
    return updater_dir() / UPDATER_FILE


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
    for path in (addin_dir(), updater_dir(), updates_dir(), backup_dir(), config_dir(), logs_dir()):
        path.mkdir(parents=True, exist_ok=True)


def load_json_file(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_json_file(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def clear_prepared_updates() -> None:
    for prepared_file in updates_dir().glob("MainVZID*.xlam"):
        try:
            prepared_file.unlink()
        except FileNotFoundError:
            pass


def cleanup_legacy_layout() -> None:
    for legacy_file in legacy_file_candidates():
        cleanup_file(legacy_file)

    for pending_root in (
        legacy_versions_dir() / "pending",
        legacy_local_vzid_dir() / "versions" / "pending",
    ):
        for pending_file in pending_root.glob("MainVZID*.xlam"):
            cleanup_file(pending_file)


def legacy_file_candidates() -> list[Path]:
    candidates = [
        legacy_loader_dir() / LEGACY_LOADER_FILE,
        legacy_versions_dir() / "current" / MAIN_FILE,
        legacy_local_vzid_dir() / "loader" / LEGACY_LOADER_FILE,
        legacy_local_vzid_dir() / "versions" / "current" / MAIN_FILE,
        legacy_local_vzid_dir() / "addin" / MAIN_FILE,
        local_cache_root() / MAIN_FILE,
        local_cache_root() / APP_NAME / MAIN_FILE,
        local_cache_root() / APP_NAME / "addin" / MAIN_FILE,
        xlstart_dir() / LEGACY_LOADER_FILE,
        xlstart_dir() / MAIN_FILE,
    ]
    return candidates


def cleanup_file(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        return


def cleanup_excel_registration(excel_version: str) -> None:
    disconnect_legacy_addins()
    remove_startup_entries_from_registry(excel_version)
    remove_addin_manager_entries(excel_version)


def sync_config() -> None:
    template = load_json_file(config_template_path())
    current = load_json_file(config_path()) if config_path().exists() else {}

    merged = dict(template)
    if isinstance(current, dict):
        merged.update(current)

    for obsolete_key in ("activeLoaderVersion", "pendingMainVersion", "pendingMainPath"):
        merged.pop(obsolete_key, None)

    active_main_version = str(template.get("activeMainVersion", "0.0.0"))
    active_updater_version = str(template.get("activeUpdaterVersion", "0.0.0"))

    merged["schemaVersion"] = str(template.get("schemaVersion", "2"))
    merged["activeMainVersion"] = active_main_version
    merged["activeUpdaterVersion"] = active_updater_version
    merged["availableMainVersion"] = ""
    merged["availableMainDownloadUrl"] = ""
    merged["preparedMainVersion"] = ""
    merged["preparedMainPath"] = ""
    merged["lastUpdateCheckAt"] = ""
    merged["lastUpdateStatus"] = "up_to_date"
    merged["lastUpdateMessage"] = f"Установлена версия {active_main_version} через setup.exe."

    save_json_file(config_path(), merged)


def copy_assets() -> None:
    shutil.copy2(bundled_path(MAIN_FILE), main_path())
    shutil.copy2(bundled_path(f"updater/{UPDATER_FILE}"), updater_path())
    if not config_path().exists():
        shutil.copy2(config_template_path(), config_path())

    clear_prepared_updates()
    sync_config()


def detect_excel_version() -> str:
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        return str(excel.Version)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def register_main_addin_in_registry(excel_version: str, addin_path: Path) -> None:
    options_key_path = f"Software\\Microsoft\\Office\\{excel_version}\\Excel\\Options"
    value_to_store = f'"{addin_path.resolve()}"'
    preserved_values = read_preserved_startup_entries(excel_version)

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, options_key_path) as key:
        clear_startup_registry_values(key)
        preserved_values.append(value_to_store)

        for index, current_value in enumerate(preserved_values):
            value_name = "OPEN" if index == 0 else f"OPEN{index}"
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, current_value)


def remove_startup_entries_from_registry(excel_version: str) -> None:
    options_key_path = f"Software\\Microsoft\\Office\\{excel_version}\\Excel\\Options"
    preserved_values = read_preserved_startup_entries(excel_version)

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, options_key_path) as key:
        clear_startup_registry_values(key)
        for index, current_value in enumerate(preserved_values):
            value_name = "OPEN" if index == 0 else f"OPEN{index}"
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, current_value)


def read_preserved_startup_entries(excel_version: str) -> list[str]:
    options_key_path = f"Software\\Microsoft\\Office\\{excel_version}\\Excel\\Options"
    preserved_values: list[str] = []

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, options_key_path) as key:
        index = 0
        while index < 64:
            value_name = "OPEN" if index == 0 else f"OPEN{index}"
            try:
                current_value, _ = winreg.QueryValueEx(key, value_name)
            except FileNotFoundError:
                index += 1
                continue

            if startup_entry_basename(str(current_value)) not in STARTUP_FILE_NAMES:
                preserved_values.append(str(current_value))
            index += 1

    return preserved_values


def clear_startup_registry_values(key) -> None:
    index = 0
    while index < 64:
        value_name = "OPEN" if index == 0 else f"OPEN{index}"
        try:
            winreg.DeleteValue(key, value_name)
        except FileNotFoundError:
            pass
        index += 1


def remove_addin_manager_entries(excel_version: str) -> None:
    key_path = f"Software\\Microsoft\\Office\\{excel_version}\\Excel\\Add-in Manager"

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            value_names: list[str] = []
            value_index = 0
            while True:
                try:
                    value_name, _, _ = winreg.EnumValue(key, value_index)
                except OSError:
                    break
                value_names.append(value_name)
                value_index += 1

            for value_name in value_names:
                if startup_entry_basename(value_name) in STARTUP_FILE_NAMES:
                    try:
                        winreg.DeleteValue(key, value_name)
                    except FileNotFoundError:
                        pass
    except FileNotFoundError:
        return


def disconnect_legacy_addins() -> None:
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook_count = excel.Workbooks.Count
        for index in range(workbook_count, 0, -1):
            workbook = excel.Workbooks(index)
            if Path(str(workbook.FullName)).name.lower() in STARTUP_FILE_NAMES:
                workbook.Close(False)

        addin_count = excel.AddIns.Count
        for index in range(1, addin_count + 1):
            addin_ref = excel.AddIns(index)
            addin_name = Path(str(addin_ref.FullName or addin_ref.Name)).name.lower()
            if addin_name not in STARTUP_FILE_NAMES:
                continue
            try:
                if addin_ref.Installed:
                    addin_ref.Installed = False
            except Exception:
                pass
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def startup_entry_basename(value: str) -> str:
    normalized_value = str(value).strip()
    if normalized_value.count('"') >= 2:
        first_quote = normalized_value.find('"')
        last_quote = normalized_value.rfind('"')
        if last_quote > first_quote:
            normalized_value = normalized_value[first_quote + 1:last_quote]
    else:
        normalized_value = normalized_value.strip('"')

    return Path(normalized_value).name.lower()


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
        excel_version = ""
        if not args.skip_register:
            excel_version = detect_excel_version()
            cleanup_excel_registration(excel_version)

        ensure_dirs()
        cleanup_legacy_layout()
        copy_assets()
        if not args.skip_register:
            register_main_addin_in_registry(excel_version, main_path())

        show_info(
            "VZID Setup",
            "Надстройка установлена для текущего пользователя.\n\n"
            "Если Excel сейчас открыт, полностью закройте все его окна и откройте Excel заново.",
        )
        return 0
    except OSError as exc:  # pragma: no cover
        if getattr(exc, "winerror", None) == 32:
            show_error(
                "VZID Setup",
                "Не удалось заменить файлы надстройки, потому что Excel все еще держит их открытыми.\n\n"
                "Полностью закройте все окна Excel и запустите setup.exe снова.",
            )
            return 1

        show_error("VZID Setup", f"Установка завершилась ошибкой:\n{exc}")
        return 1
    except Exception as exc:  # pragma: no cover
        show_error("VZID Setup", f"Установка завершилась ошибкой:\n{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
