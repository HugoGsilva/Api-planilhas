from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

MISSING_CNPJ_COLUMN_MESSAGE = (
    "nao tem cnpj na sua planilha adicione uma coluna CNPJ e os cnpj em baixo"
)

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


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
            sheet_xml = archive.read("xl/worksheets/sheet1.xml")
    except (KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
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
