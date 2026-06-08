from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from posixpath import join as posix_join
from posixpath import normpath as posix_normpath

MISSING_CNPJ_COLUMN_MESSAGE = (
    "nao tem cnpj na sua planilha adicione uma coluna CNPJ e os cnpj em baixo"
)

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class MissingCnpjColumnError(ValueError):
    pass


class InvalidXlsxError(ValueError):
    pass


def _cell_column(cell_ref: str) -> int:
    letters = "".join(char for char in cell_ref if char.isalpha())
    result = 0
    for char in letters:
        result = result * 26 + (ord(char.upper()) - 64)
    return result


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        shared_xml = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(shared_xml)
    values: list[str] = []
    for item in root.findall(f".//{{{SPREADSHEET_NS}}}si"):
        values.append(
            "".join(
                node.text or ""
                for node in item.findall(f".//{{{SPREADSHEET_NS}}}t")
            )
        )
    return values


def _resolve_workbook_target(target: str) -> str:
    target = target.replace("\\", "/").strip()
    if not target:
        raise InvalidXlsxError("Planilha XLSX invalida")

    if target.startswith("/"):
        path = posix_normpath(target.lstrip("/"))
    else:
        path = posix_normpath(posix_join("xl", target))

    if path.startswith("../") or path == "..":
        raise InvalidXlsxError("Planilha XLSX invalida")
    return path


def _first_sheet_path(archive: zipfile.ZipFile) -> str:
    workbook_xml = archive.read("xl/workbook.xml")
    rels_xml = archive.read("xl/_rels/workbook.xml.rels")
    workbook_root = ET.fromstring(workbook_xml)
    rels_root = ET.fromstring(rels_xml)

    first_sheet = workbook_root.find(
        f"./{{{SPREADSHEET_NS}}}sheets/{{{SPREADSHEET_NS}}}sheet"
    )
    if first_sheet is None:
        raise InvalidXlsxError("Planilha XLSX invalida")

    relationship_id = first_sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id")
    if not relationship_id:
        raise InvalidXlsxError("Planilha XLSX invalida")

    for relationship in rels_root.findall(f"./{{{PACKAGE_REL_NS}}}Relationship"):
        if relationship.attrib.get("Id") != relationship_id:
            continue
        return _resolve_workbook_target(relationship.attrib.get("Target", ""))

    raise InvalidXlsxError("Planilha XLSX invalida")


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "s":
        value = cell.find(f"./{{{SPREADSHEET_NS}}}v")
        if value is None or value.text is None:
            return ""
        try:
            return shared_strings[int(value.text)].strip()
        except (ValueError, IndexError) as exc:
            raise InvalidXlsxError("Planilha XLSX invalida") from exc

    texts = [
        node.text or ""
        for node in cell.findall(f".//{{{SPREADSHEET_NS}}}t")
    ]
    if texts:
        return "".join(texts).strip()

    value = cell.find(f"./{{{SPREADSHEET_NS}}}v")
    return (value.text or "").strip() if value is not None else ""


def _read_rows(content: bytes) -> list[list[str]]:
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            shared_strings = _read_shared_strings(archive)
            sheet_xml = archive.read(_first_sheet_path(archive))
    except (
        KeyError,
        zipfile.BadZipFile,
        ET.ParseError,
        InvalidXlsxError,
    ) as exc:
        raise InvalidXlsxError("Planilha XLSX invalida") from exc

    try:
        root = ET.fromstring(sheet_xml)
    except ET.ParseError as exc:
        raise InvalidXlsxError("Planilha XLSX invalida") from exc

    rows: list[list[str]] = []
    for row in root.findall(f".//{{{SPREADSHEET_NS}}}row"):
        values_by_column: dict[int, str] = {}
        max_column = 0
        for cell in row.findall(f"./{{{SPREADSHEET_NS}}}c"):
            ref = cell.attrib.get("r", "")
            column = _cell_column(ref)
            if column <= 0:
                continue
            values_by_column[column] = _cell_text(cell, shared_strings)
            max_column = max(max_column, column)
        rows.append(
            [values_by_column.get(index, "") for index in range(1, max_column + 1)]
        )
    return rows


def extract_cnpj_values(content: bytes) -> list[str]:
    rows = _read_rows(content)
    if not rows:
        raise MissingCnpjColumnError(MISSING_CNPJ_COLUMN_MESSAGE)

    header = [value.strip().upper() for value in rows[0]]
    try:
        cnpj_index = header.index("CNPJ")
    except ValueError as exc:
        raise MissingCnpjColumnError(MISSING_CNPJ_COLUMN_MESSAGE) from exc

    values: list[str] = []
    for row in rows[1:]:
        if cnpj_index >= len(row):
            continue
        value = row[cnpj_index].strip()
        if value:
            values.append(value)
    return values
