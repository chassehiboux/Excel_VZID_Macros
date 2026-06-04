from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
MAIN_SRC = ROOT / "src" / "main-vzid"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_UI_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
CUSTOM_UI_PART_NAME = "/customUI/customUI14.xml"
CUSTOM_UI_TARGET = "customUI/customUI14.xml"
CUSTOM_UI_CONTENT_TYPE = "application/xml"

VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VB_NAME_PATTERN = re.compile(r'^Attribute VB_Name = "([^"]+)"$', re.MULTILINE)
SOURCE_TEXT_ENCODINGS = ("utf-8", "cp1251")
SKIPPED_FORM_NAMES = {
    "frmVZID_KGN.frm",
    "frmVZID_TMN.frm",
    "frmVZID_EKB.frm",
    "frmVZID_CHLB.frm",
}


def import_components(vb_project, source_dir: Path) -> None:
    if not source_dir.exists():
        return

    for extension in (".bas", ".cls"):
        for path in sorted(source_dir.glob(f"*{extension}")):
            import_text_component(vb_project, path)

    for path in sorted(source_dir.glob("*.frm")):
        if path.name in SKIPPED_FORM_NAMES:
            continue
        validate_form_dependencies(path)
        vb_project.VBComponents.Import(str(path))


def import_text_component(vb_project, source_path: Path) -> None:
    source_text, source_encoding = read_source_text_with_encoding(source_path)
    if source_encoding != "utf-8":
        vb_project.VBComponents.Import(str(source_path))
        return

    component_name = extract_vb_name(source_text, source_path)
    component_type = component_type_for_path(source_path)
    module_text = remove_attribute_block(source_text)

    component = vb_project.VBComponents.Add(component_type)
    component.Name = component_name

    code_module = component.CodeModule
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(module_text)


def extract_vb_name(source_text: str, source_path: Path) -> str:
    match = VB_NAME_PATTERN.search(source_text)
    if match is None:
        raise ValueError(f"Missing Attribute VB_Name in {source_path}")
    return match.group(1)


def component_type_for_path(source_path: Path) -> int:
    if source_path.suffix.lower() == ".bas":
        return VBEXT_CT_STD_MODULE
    if source_path.suffix.lower() == ".cls":
        return VBEXT_CT_CLASS_MODULE
    raise ValueError(f"Unsupported text component type: {source_path}")


def remove_attribute_block(source_text: str) -> str:
    lines = source_text.splitlines()
    body_start = 0
    for index, line in enumerate(lines):
        if line.startswith("Attribute VB_"):
            body_start = index + 1
            continue
        if body_start and line == "":
            body_start = index + 1
            continue
        break

    return "\n".join(lines[body_start:])


def validate_form_dependencies(form_path: Path) -> None:
    form_text = read_source_text(form_path)
    frx_path = form_path.with_suffix(".frx")

    if not form_references_frx(form_text):
        return

    if frx_path.exists():
        return

    raise FileNotFoundError(
        f"UserForm {form_path.name} references {frx_path.name}, but the .frx file is missing. "
        "Export and commit both files together."
    )


def form_references_frx(form_text: str) -> bool:
    return ".frx\":" in form_text.lower()


def replace_thisworkbook_code(workbook, vb_project, workbook_code_path: Path) -> None:
    code_module = vb_project.VBComponents(workbook.CodeName).CodeModule
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(read_source_text(workbook_code_path))


def read_source_text(source_path: Path) -> str:
    return read_source_text_with_encoding(source_path)[0]


def read_source_text_with_encoding(source_path: Path) -> tuple[str, str]:
    source_bytes = source_path.read_bytes()
    for encoding in SOURCE_TEXT_ENCODINGS:
        try:
            return source_bytes.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(
        "unknown",
        source_bytes,
        0,
        len(source_bytes),
        f"Unable to decode {source_path} as UTF-8 or CP1251",
    )


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
        content_changed = False
        override_node = None
        for node in content_root.findall(override_tag):
            if node.attrib.get("PartName") == CUSTOM_UI_PART_NAME:
                override_node = node
                break

        if override_node is None:
            override_node = ET.SubElement(
                content_root,
                override_tag,
                {
                    "PartName": CUSTOM_UI_PART_NAME,
                    "ContentType": CUSTOM_UI_CONTENT_TYPE,
                },
            )
            content_changed = True
        elif override_node.attrib.get("ContentType") != CUSTOM_UI_CONTENT_TYPE:
            override_node.set("ContentType", CUSTOM_UI_CONTENT_TYPE)
            content_changed = True

        if content_changed:
            write_xml_with_default_namespace(content_tree, content_types_path, CONTENT_TYPES_NS)

        rels_path = temp_dir / "_rels" / ".rels"
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        relationship_tag = f"{{{RELS_NS}}}Relationship"
        rel_changed = False
        custom_ui_rel = None
        existing_ids = []
        for node in rels_root.findall(relationship_tag):
            existing_ids.append(node.attrib.get("Id", ""))
            if node.attrib.get("Target") in ("customUI/customUI14.xml", CUSTOM_UI_TARGET, CUSTOM_UI_PART_NAME):
                custom_ui_rel = node

        if custom_ui_rel is None:
            next_index = 1
            while f"rId{next_index}" in existing_ids:
                next_index += 1

            custom_ui_rel = ET.SubElement(
                rels_root,
                relationship_tag,
                {
                    "Id": f"rId{next_index}",
                    "Type": CUSTOM_UI_REL_TYPE,
                    "Target": CUSTOM_UI_TARGET,
                },
            )
            rel_changed = True
        else:
            if custom_ui_rel.attrib.get("Type") != CUSTOM_UI_REL_TYPE:
                custom_ui_rel.set("Type", CUSTOM_UI_REL_TYPE)
                rel_changed = True
            if custom_ui_rel.attrib.get("Target") != CUSTOM_UI_TARGET:
                custom_ui_rel.set("Target", CUSTOM_UI_TARGET)
                rel_changed = True

        if rel_changed:
            write_xml_with_default_namespace(rels_tree, rels_path, RELS_NS)

        rebuilt_path = xlam_path.with_suffix(".rebuilt")
        if rebuilt_path.exists():
            rebuilt_path.unlink()

        with zipfile.ZipFile(rebuilt_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in sorted(temp_dir.rglob("*")):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_dir).as_posix())

        rebuilt_path.replace(xlam_path)


def build_main(output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    main_path = output_dir / "MainVZID.xlam"
    create_addin(MAIN_SRC, main_path)
    patch_custom_ui(main_path, MAIN_SRC / "customui" / "customUI14.xml")
    return main_path


def write_xml_with_default_namespace(tree: ET.ElementTree, target_path: Path, namespace_uri: str) -> None:
    xml_text = ET.tostring(tree.getroot(), encoding="unicode")
    xml_text = xml_text.replace("ns0:", "")
    xml_text = xml_text.replace(":ns0", "")
    xml_text = xml_text.replace('xmlns:ns0="' + namespace_uri + '"', 'xmlns="' + namespace_uri + '"')
    target_path.write_text("<?xml version='1.0' encoding='utf-8'?>\n" + xml_text, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default=str(ROOT / "build"))
    args = parser.parse_args()

    output_dir = Path(args.output_dir).resolve()
    main_path = build_main(output_dir)
    print(main_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
