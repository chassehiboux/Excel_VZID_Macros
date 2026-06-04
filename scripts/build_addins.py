from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
LOADER_SRC = ROOT / "src" / "vzid-loader"
MAIN_SRC = ROOT / "src" / "main-vzid"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_UI_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"

ET.register_namespace("", CONTENT_TYPES_NS)
ET.register_namespace("", RELS_NS)


def import_components(vb_project, source_dir: Path) -> None:
    if not source_dir.exists():
        return

    for extension in (".bas", ".cls", ".frm"):
        for path in sorted(source_dir.glob(f"*{extension}")):
            vb_project.VBComponents.Import(str(path))


def replace_thisworkbook_code(workbook, vb_project, workbook_code_path: Path) -> None:
    code_module = vb_project.VBComponents(workbook.CodeName).CodeModule
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(workbook_code_path.read_text(encoding="utf-8"))


def create_addin(source_root: Path, output_path: Path) -> None:
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Add()
        workbook.IsAddin = True

        vb_project = workbook.VBProject
        import_components(vb_project, source_root / "modules")
        import_components(vb_project, source_root / "forms")
        replace_thisworkbook_code(workbook, vb_project, source_root / "workbook" / "ThisWorkbook.cls")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            output_path.unlink()
        workbook.SaveAs(str(output_path), FileFormat=55)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def patch_custom_ui(xlam_path: Path, custom_ui_path: Path) -> None:
    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        with zipfile.ZipFile(xlam_path, "r") as archive:
            archive.extractall(temp_dir)

        custom_ui_dir = temp_dir / "customUI"
        custom_ui_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(custom_ui_path, custom_ui_dir / "customUI14.xml")

        content_types_path = temp_dir / "[Content_Types].xml"
        content_tree = ET.parse(content_types_path)
        content_root = content_tree.getroot()
        override_tag = f"{{{CONTENT_TYPES_NS}}}Override"
        override_exists = any(
            node.attrib.get("PartName") == "/customUI/customUI14.xml"
            for node in content_root.findall(override_tag)
        )
        if not override_exists:
            ET.SubElement(
                content_root,
                override_tag,
                {
                    "PartName": "/customUI/customUI14.xml",
                    "ContentType": "application/vnd.ms-office.customUI+xml",
                },
            )
            content_tree.write(content_types_path, encoding="utf-8", xml_declaration=True)

        rels_path = temp_dir / "_rels" / ".rels"
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        relationship_tag = f"{{{RELS_NS}}}Relationship"
        rel_exists = any(
            node.attrib.get("Type") == CUSTOM_UI_REL_TYPE
            for node in rels_root.findall(relationship_tag)
        )
        if not rel_exists:
            existing_ids = [node.attrib.get("Id", "") for node in rels_root.findall(relationship_tag)]
            next_index = 1
            while f"rId{next_index}" in existing_ids:
                next_index += 1

            ET.SubElement(
                rels_root,
                relationship_tag,
                {
                    "Id": f"rId{next_index}",
                    "Type": CUSTOM_UI_REL_TYPE,
                    "Target": "customUI/customUI14.xml",
                },
            )
            rels_tree.write(rels_path, encoding="utf-8", xml_declaration=True)

        rebuilt_path = xlam_path.with_suffix(".rebuilt")
        if rebuilt_path.exists():
            rebuilt_path.unlink()

        with zipfile.ZipFile(rebuilt_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in sorted(temp_dir.rglob("*")):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_dir).as_posix())

        rebuilt_path.replace(xlam_path)


def build_all(output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)

    loader_path = output_dir / "LoaderVZID.xlam"
    main_path = output_dir / "MainVZID.xlam"

    create_addin(LOADER_SRC, loader_path)
    create_addin(MAIN_SRC, main_path)
    patch_custom_ui(main_path, MAIN_SRC / "customui" / "customUI14.xml")

    return loader_path, main_path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default=str(ROOT / "build"))
    args = parser.parse_args()

    output_dir = Path(args.output_dir).resolve()
    loader_path, main_path = build_all(output_dir)
    print(loader_path)
    print(main_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
