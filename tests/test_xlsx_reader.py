import sys
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.xlsx_reader import (
    InvalidXlsxError,
    MISSING_CNPJ_COLUMN_MESSAGE,
    MissingCnpjColumnError,
    extract_cnpj_values,
)


def _column_name(index):
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def build_minimal_xlsx(rows, *, sheet_target="worksheets/sheet1.xml"):
    def cell_ref(row_index, col_index):
        return f"{_column_name(col_index)}{row_index}"

    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            ref = cell_ref(row_index, col_index)
            cells.append(
                f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'
            )
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    sheet = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        "</worksheet>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Entrada" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="worksheet" Target="{sheet_target}"/>'
        "</Relationships>"
    )
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        archive.writestr(f"xl/{sheet_target}", sheet)
    return buffer.getvalue()


def build_shared_strings_xlsx(rows, *, sheet_target="worksheets/sheet1.xml"):
    values = []
    indexes = {}

    def shared_index(value):
        if value not in indexes:
            indexes[value] = len(values)
            values.append(value)
        return indexes[value]

    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            ref = f"{_column_name(col_index)}{row_index}"
            cells.append(f'<c r="{ref}" t="s"><v>{shared_index(value)}</v></c>')
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    sheet = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        "</worksheet>"
    )
    shared_strings = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(f"<si><t>{value}</t></si>" for value in values)
        + "</sst>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Entrada" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="worksheet" Target="{sheet_target}"/>'
        "</Relationships>"
    )
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        archive.writestr(f"xl/{sheet_target}", sheet)
        archive.writestr("xl/sharedStrings.xml", shared_strings)
    return buffer.getvalue()


class XlsxReaderTest(unittest.TestCase):
    def test_extracts_cnpj_column_values(self):
        content = build_minimal_xlsx(
            [
                ["Nome", " CNPJ "],
                ["Empresa A", "12.345.678/0001-90"],
                ["Empresa B", "11222333000144"],
                ["Empresa C", ""],
            ]
        )

        values = extract_cnpj_values(content)

        self.assertEqual(values, ["12.345.678/0001-90", "11222333000144"])

    def test_rejects_spreadsheet_without_cnpj_column(self):
        content = build_minimal_xlsx([["Nome", "Documento"], ["Empresa A", "123"]])

        with self.assertRaises(MissingCnpjColumnError) as ctx:
            extract_cnpj_values(content)

        self.assertEqual(str(ctx.exception), MISSING_CNPJ_COLUMN_MESSAGE)

    def test_extracts_shared_string_values(self):
        content = build_shared_strings_xlsx(
            [["Nome", "CNPJ"], ["Empresa A", "12.345.678/0001-90"]]
        )

        values = extract_cnpj_values(content)

        self.assertEqual(values, ["12.345.678/0001-90"])

    def test_extracts_cnpj_column_after_z(self):
        header = [""] * 26 + ["CNPJ"]
        data = [""] * 26 + ["11222333000144"]
        content = build_minimal_xlsx([header, data])

        values = extract_cnpj_values(content)

        self.assertEqual(values, ["11222333000144"])

    def test_maps_invalid_xml_to_invalid_xlsx_error(self):
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/workbook.xml", "<workbook")
            archive.writestr("xl/_rels/workbook.xml.rels", "<Relationships />")

        with self.assertRaises(InvalidXlsxError) as ctx:
            extract_cnpj_values(buffer.getvalue())

        self.assertEqual(str(ctx.exception), "Planilha XLSX invalida")

    def test_maps_invalid_shared_string_index_to_invalid_xlsx_error(self):
        workbook = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Entrada" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>"
        )
        rels = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>"
        )
        shared_strings = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<si><t>CNPJ</t></si>"
            "</sst>"
        )
        sheet = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData><row r="1"><c r="A1" t="s"><v>99</v></c></row></sheetData>'
            "</worksheet>"
        )
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/workbook.xml", workbook)
            archive.writestr("xl/_rels/workbook.xml.rels", rels)
            archive.writestr("xl/sharedStrings.xml", shared_strings)
            archive.writestr(
                "xl/worksheets/sheet1.xml",
                sheet,
            )

        with self.assertRaises(InvalidXlsxError) as ctx:
            extract_cnpj_values(buffer.getvalue())

        self.assertEqual(str(ctx.exception), "Planilha XLSX invalida")

    def test_reads_first_real_sheet_from_workbook_relationships(self):
        buffer = BytesIO()
        workbook = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>'
            '<sheet name="Primeira" sheetId="1" r:id="rIdSheet2"/>'
            '<sheet name="Segunda" sheetId="2" r:id="rIdSheet1"/>'
            "</sheets>"
            "</workbook>"
        )
        rels = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdSheet1" Type="worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rIdSheet2" Type="worksheet" Target="worksheets/sheet2.xml"/>'
            "</Relationships>"
        )
        sheet1 = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<sheetData><row><c r=\"A1\" t=\"inlineStr\"><is><t>Nome</t></is></c></row></sheetData>"
            "</worksheet>"
        )
        sheet2 = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<sheetData>"
            '<row><c r="A1" t="inlineStr"><is><t>CNPJ</t></is></c></row>'
            '<row><c r="A2" t="inlineStr"><is><t>12.345.678/0001-90</t></is></c></row>'
            "</sheetData>"
            "</worksheet>"
        )
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/workbook.xml", workbook)
            archive.writestr("xl/_rels/workbook.xml.rels", rels)
            archive.writestr("xl/worksheets/sheet1.xml", sheet1)
            archive.writestr("xl/worksheets/sheet2.xml", sheet2)

        values = extract_cnpj_values(buffer.getvalue())

        self.assertEqual(values, ["12.345.678/0001-90"])


if __name__ == "__main__":
    unittest.main()
