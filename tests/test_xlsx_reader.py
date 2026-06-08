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
    MISSING_CNPJ_COLUMN_MESSAGE,
    MissingCnpjColumnError,
    extract_cnpj_values,
)


def build_minimal_xlsx(rows):
    def cell_ref(row_index, col_index):
        return f"{chr(64 + col_index)}{row_index}"

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
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
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


if __name__ == "__main__":
    unittest.main()
