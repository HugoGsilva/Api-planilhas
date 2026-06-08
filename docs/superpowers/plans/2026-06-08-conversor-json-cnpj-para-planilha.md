# Conversor JSON CNPJ Para Planilha Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Python converter that reads DirectD CNPJ JSON files from `cnpjs/cnpjs` and generates an `.xlsx` import spreadsheet with the approved columns.

**Architecture:** Use a small Python package with pure functions for extraction and a CLI wrapper for batch conversion. Generate XLSX using the Python standard library (`zipfile` + XML) to avoid installing dependencies during the first delivery.

**Tech Stack:** Python 3, standard library only, `unittest` for tests.

---

## File Structure

- Create `src/api_planilhas/__init__.py`: package marker and public version.
- Create `src/api_planilhas/converter.py`: approved headers, JSON extraction, row mapping, XLSX writer and batch conversion.
- Create `scripts/converter_json_para_planilha.py`: PowerShell-friendly CLI entrypoint.
- Create `tests/test_converter.py`: focused tests for mapping, missing data, invalid JSON handling and XLSX output.

No commits are included in the steps because the project-level instructions prohibit committing without explicit user approval.

---

### Task 1: Core Mapping

**Files:**
- Create: `src/api_planilhas/__init__.py`
- Create: `src/api_planilhas/converter.py`
- Test: `tests/test_converter.py`

- [ ] **Step 1: Create the package marker**

Create `src/api_planilhas/__init__.py`:

```python
"""Utilities for converting DirectD CNPJ JSON files into spreadsheets."""

__version__ = "0.1.0"
```

- [ ] **Step 2: Write failing mapping tests**

Create `tests/test_converter.py` with:

```python
import unittest

from api_planilhas.converter import HEADERS, extract_row


class ExtractRowTest(unittest.TestCase):
    def test_extracts_complete_json_to_approved_columns(self):
        payload = {
            "metaDados": {"resultado": "OK"},
            "retorno": {
                "cnpj": "00019000000133",
                "razaoSocial": "EMPRESA TESTE LTDA",
                "nomeFantasia": "EMPRESA TESTE",
                "cnaeCodigo": 6201501,
                "cnaeDescricao": "Desenvolvimento de programas de computador sob encomenda",
                "quantidadeFuncionarios": 42,
                "situacaoCadastral": "ATIVA",
                "telefones": [
                    {"telefoneComDDD": "1130000001"},
                    {"telefoneComDDD": "1130000002"},
                    {"telefoneComDDD": "1130000003"},
                ],
                "emails": [
                    {"enderecoEmail": "principal@example.com"},
                    {"enderecoEmail": "secundario@example.com"},
                    {"enderecoEmail": "terceiro@example.com"},
                ],
                "enderecos": [
                    {
                        "cep": "01001000",
                        "logradouro": "PRACA DA SE",
                        "numero": "100",
                        "complemento": "CJ 10",
                        "bairro": "SE",
                        "cidade": "SAO PAULO",
                        "uf": "SP",
                    }
                ],
                "socios": [
                    {
                        "documento": "12345678900",
                        "cargo": "SOCIO ADMINISTRADOR",
                    }
                ],
            },
        }

        row = extract_row(payload)

        self.assertEqual(len(row), len(HEADERS))
        self.assertEqual(row[HEADERS.index("Pessoa")], "EMPRESA TESTE LTDA")
        self.assertEqual(row[HEADERS.index("Telefone1")], "1130000001")
        self.assertEqual(row[HEADERS.index("Telefone2")], "1130000002")
        self.assertEqual(row[HEADERS.index("email Prin")], "principal@example.com")
        self.assertEqual(row[HEADERS.index("email Secu")], "secundario@example.com")
        self.assertEqual(row[HEADERS.index("Cargo")], "SOCIO ADMINISTRADOR")
        self.assertEqual(row[HEADERS.index("CPF")], "12345678900")
        self.assertEqual(row[HEADERS.index("Organizacao")], "EMPRESA TESTE LTDA")
        self.assertEqual(row[HEADERS.index("CNPJ")], "00019000000133")
        self.assertEqual(row[HEADERS.index("CEP")], "01001000")
        self.assertEqual(row[HEADERS.index("Logradouro")], "PRACA DA SE")
        self.assertEqual(row[HEADERS.index("Numero")], "100")
        self.assertEqual(row[HEADERS.index("Complemento")], "CJ 10")
        self.assertEqual(row[HEADERS.index("Bairro")], "SE")
        self.assertEqual(row[HEADERS.index("Cidade")], "SAO PAULO")
        self.assertEqual(row[HEADERS.index("UF")], "SP")
        self.assertEqual(row[HEADERS.index("IE")], "")
        self.assertEqual(row[HEADERS.index("Segmento")], "")
        self.assertEqual(row[HEADERS.index("Valor Estimado")], "")
        self.assertEqual(row[HEADERS.index("*Nome Fantasia")], "EMPRESA TESTE")
        self.assertEqual(
            row[HEADERS.index("*CNAE")],
            "6201501 - Desenvolvimento de programas de computador sob encomenda",
        )
        self.assertEqual(row[HEADERS.index("*Vendedor")], "")
        self.assertEqual(row[HEADERS.index("Titulo da oportunidade")], "")
        self.assertEqual(row[HEADERS.index("Status")], "ATIVA")
        self.assertEqual(row[HEADERS.index("Data Ganho/Perdido")], "")
        self.assertEqual(row[HEADERS.index("*QUANTIDADE DE FUNCIONARIOS")], 42)

    def test_extracts_missing_lists_as_empty_cells(self):
        payload = {
            "retorno": {
                "cnpj": "00000000000000",
                "razaoSocial": "SEM LISTAS LTDA",
                "cnaeDescricao": "Comercio varejista",
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("Telefone1")], "")
        self.assertEqual(row[HEADERS.index("Telefone2")], "")
        self.assertEqual(row[HEADERS.index("email Prin")], "")
        self.assertEqual(row[HEADERS.index("email Secu")], "")
        self.assertEqual(row[HEADERS.index("Cargo")], "")
        self.assertEqual(row[HEADERS.index("CPF")], "")
        self.assertEqual(row[HEADERS.index("CEP")], "")
        self.assertEqual(row[HEADERS.index("*CNAE")], "Comercio varejista")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3: Run tests and verify they fail because implementation is missing**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter -v
```

Expected: failure importing `api_planilhas.converter` or `extract_row`.

- [ ] **Step 4: Implement mapping functions**

Create `src/api_planilhas/converter.py` with:

```python
from __future__ import annotations

from typing import Any


HEADERS = [
    "Pessoa",
    "Telefone1",
    "Telefone2",
    "email Prin",
    "email Secu",
    "Cargo",
    "CPF",
    "Organizacao",
    "CNPJ",
    "CEP",
    "Logradouro",
    "Numero",
    "Complemento",
    "Bairro",
    "Cidade",
    "UF",
    "IE",
    "Segmento",
    "Valor Estimado",
    "*Nome Fantasia",
    "*CNAE",
    "*Vendedor",
    "Titulo da oportunidade",
    "Status",
    "Data Ganho/Perdido",
    "*QUANTIDADE DE FUNCIONARIOS",
]


def _retorno(payload: dict[str, Any]) -> dict[str, Any]:
    retorno = payload.get("retorno")
    return retorno if isinstance(retorno, dict) else {}


def _list_item(data: dict[str, Any], key: str, index: int) -> dict[str, Any]:
    items = data.get(key)
    if not isinstance(items, list) or index >= len(items):
        return {}
    item = items[index]
    return item if isinstance(item, dict) else {}


def _value(data: dict[str, Any], key: str) -> Any:
    value = data.get(key)
    return "" if value is None else value


def _cnae(data: dict[str, Any]) -> str:
    codigo = _value(data, "cnaeCodigo")
    descricao = _value(data, "cnaeDescricao")
    if codigo != "" and descricao != "":
        return f"{codigo} - {descricao}"
    if descricao != "":
        return str(descricao)
    if codigo != "":
        return str(codigo)
    return ""


def extract_row(payload: dict[str, Any]) -> list[Any]:
    data = _retorno(payload)
    telefone_1 = _list_item(data, "telefones", 0)
    telefone_2 = _list_item(data, "telefones", 1)
    email_1 = _list_item(data, "emails", 0)
    email_2 = _list_item(data, "emails", 1)
    endereco = _list_item(data, "enderecos", 0)
    socio = _list_item(data, "socios", 0)

    return [
        _value(data, "razaoSocial"),
        _value(telefone_1, "telefoneComDDD"),
        _value(telefone_2, "telefoneComDDD"),
        _value(email_1, "enderecoEmail"),
        _value(email_2, "enderecoEmail"),
        _value(socio, "cargo"),
        _value(socio, "documento"),
        _value(data, "razaoSocial"),
        _value(data, "cnpj"),
        _value(endereco, "cep"),
        _value(endereco, "logradouro"),
        _value(endereco, "numero"),
        _value(endereco, "complemento"),
        _value(endereco, "bairro"),
        _value(endereco, "cidade"),
        _value(endereco, "uf"),
        "",
        "",
        "",
        _value(data, "nomeFantasia"),
        _cnae(data),
        "",
        "",
        _value(data, "situacaoCadastral"),
        "",
        _value(data, "quantidadeFuncionarios"),
    ]
```

- [ ] **Step 5: Run mapping tests and verify they pass**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter -v
```

Expected: `OK`.

---

### Task 2: XLSX Writer

**Files:**
- Modify: `src/api_planilhas/converter.py`
- Modify: `tests/test_converter.py`

- [ ] **Step 1: Add failing XLSX writer test**

Append this test method inside `ExtractRowTest` in `tests/test_converter.py`:

```python
    def test_writes_xlsx_file_with_headers_and_rows(self):
        import tempfile
        import zipfile
        from pathlib import Path

        from api_planilhas.converter import write_xlsx

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "saida.xlsx"
            write_xlsx(output, [["EMPRESA TESTE LTDA"] + [""] * (len(HEADERS) - 1)])

            self.assertTrue(output.exists())
            with zipfile.ZipFile(output) as archive:
                names = set(archive.namelist())
                self.assertIn("[Content_Types].xml", names)
                self.assertIn("xl/workbook.xml", names)
                self.assertIn("xl/worksheets/sheet1.xml", names)
                sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

            self.assertIn("Pessoa", sheet)
            self.assertIn("EMPRESA TESTE LTDA", sheet)
```

- [ ] **Step 2: Run test and verify it fails because writer is missing**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter.ExtractRowTest.test_writes_xlsx_file_with_headers_and_rows -v
```

Expected: failure importing `write_xlsx`.

- [ ] **Step 3: Add XLSX writer implementation**

Append this code to `src/api_planilhas/converter.py`:

```python
import html
import zipfile
from pathlib import Path


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _cell_xml(row_index: int, column_index: int, value: Any) -> str:
    ref = f"{_column_name(column_index)}{row_index}"
    if value == "":
        return f'<c r="{ref}"/>'
    escaped = html.escape(str(value), quote=False)
    return f'<c r="{ref}" t="inlineStr"><is><t>{escaped}</t></is></c>'


def _row_xml(row_index: int, values: list[Any]) -> str:
    cells = "".join(_cell_xml(row_index, index, value) for index, value in enumerate(values, start=1))
    return f'<row r="{row_index}">{cells}</row>'


def _sheet_xml(rows: list[list[Any]]) -> str:
    all_rows = [HEADERS, *rows]
    row_xml = "".join(_row_xml(index, row) for index, row in enumerate(all_rows, start=1))
    dimension = f"A1:{_column_name(len(HEADERS))}{len(all_rows)}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        '<sheetData>'
        f"{row_xml}"
        '</sheetData>'
        '</worksheet>'
    )


def write_xlsx(output_path: str | Path, rows: list[list[Any]]) -> None:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>',
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="importacao_oportunidades" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '</Relationships>',
        )
        archive.writestr("xl/worksheets/sheet1.xml", _sheet_xml(rows))
```

- [ ] **Step 4: Run XLSX writer test**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter.ExtractRowTest.test_writes_xlsx_file_with_headers_and_rows -v
```

Expected: `OK`.

- [ ] **Step 5: Run all current tests**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter -v
```

Expected: `OK`.

---

### Task 3: Batch Conversion

**Files:**
- Modify: `src/api_planilhas/converter.py`
- Modify: `tests/test_converter.py`

- [ ] **Step 1: Add failing batch conversion test**

Append this test method inside `ExtractRowTest` in `tests/test_converter.py`:

```python
    def test_converts_directory_and_reports_invalid_json(self):
        import json
        import tempfile
        import zipfile
        from pathlib import Path

        from api_planilhas.converter import convert_directory

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "cnpjs"
            source.mkdir()
            valid_payload = {
                "retorno": {
                    "cnpj": "00019000000133",
                    "razaoSocial": "EMPRESA TESTE LTDA",
                    "telefones": [],
                    "emails": [],
                    "enderecos": [],
                    "socios": [],
                }
            }
            (source / "00019000000133.json").write_text(json.dumps(valid_payload), encoding="utf-8")
            (source / "invalido.json").write_text("{json invalido", encoding="utf-8")
            output = root / "saida.xlsx"

            result = convert_directory(source, output)

            self.assertEqual(result.processed_files, 1)
            self.assertEqual(result.generated_rows, 1)
            self.assertEqual(len(result.errors), 1)
            self.assertEqual(result.errors[0].file_name, "invalido.json")
            with zipfile.ZipFile(output) as archive:
                sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("EMPRESA TESTE LTDA", sheet)
```

- [ ] **Step 2: Run test and verify it fails because batch conversion is missing**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter.ExtractRowTest.test_converts_directory_and_reports_invalid_json -v
```

Expected: failure importing `convert_directory`.

- [ ] **Step 3: Add batch conversion implementation**

Append this code to `src/api_planilhas/converter.py`:

```python
import json
from dataclasses import dataclass


@dataclass(frozen=True)
class ConversionError:
    file_name: str
    message: str


@dataclass(frozen=True)
class ConversionResult:
    processed_files: int
    generated_rows: int
    errors: list[ConversionError]
    output_path: Path


def convert_directory(source_dir: str | Path, output_path: str | Path) -> ConversionResult:
    source = Path(source_dir)
    output = Path(output_path)
    if not source.exists():
        raise FileNotFoundError(f"Diretorio fonte nao encontrado: {source}")
    if not source.is_dir():
        raise NotADirectoryError(f"Caminho fonte nao e um diretorio: {source}")

    rows: list[list[Any]] = []
    errors: list[ConversionError] = []

    for file_path in sorted(source.glob("*.json")):
        try:
            payload = json.loads(file_path.read_text(encoding="utf-8-sig"))
            if not isinstance(payload, dict):
                raise ValueError("JSON raiz precisa ser um objeto")
            rows.append(extract_row(payload))
        except Exception as exc:
            errors.append(ConversionError(file_name=file_path.name, message=str(exc)))

    write_xlsx(output, rows)
    return ConversionResult(
        processed_files=len(rows),
        generated_rows=len(rows),
        errors=errors,
        output_path=output,
    )
```

- [ ] **Step 4: Run batch conversion test**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter.ExtractRowTest.test_converts_directory_and_reports_invalid_json -v
```

Expected: `OK`.

- [ ] **Step 5: Run all tests**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter -v
```

Expected: `OK`.

---

### Task 4: CLI Entrypoint

**Files:**
- Create: `scripts/converter_json_para_planilha.py`
- Modify: `tests/test_converter.py`

- [ ] **Step 1: Add failing CLI smoke test**

Append this test method inside `ExtractRowTest` in `tests/test_converter.py`:

```python
    def test_cli_script_exists(self):
        from pathlib import Path

        script = Path("scripts/converter_json_para_planilha.py")

        self.assertTrue(script.exists())
        self.assertIn("convert_directory", script.read_text(encoding="utf-8"))
```

- [ ] **Step 2: Run test and verify it fails because script is missing**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter.ExtractRowTest.test_cli_script_exists -v
```

Expected: failure because `scripts/converter_json_para_planilha.py` does not exist.

- [ ] **Step 3: Create CLI script**

Create `scripts/converter_json_para_planilha.py`:

```python
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from api_planilhas.converter import convert_directory


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Converte JSONs de CNPJ da DirectD em planilha XLSX de importacao."
    )
    parser.add_argument(
        "--source",
        default="cnpjs/cnpjs",
        help="Diretorio com arquivos JSON. Padrao: cnpjs/cnpjs",
    )
    parser.add_argument(
        "--output",
        default="dist/importacao_oportunidades.xlsx",
        help="Caminho da planilha XLSX gerada. Padrao: dist/importacao_oportunidades.xlsx",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    result = convert_directory(args.source, args.output)
    print(f"Arquivos processados: {result.processed_files}")
    print(f"Linhas geradas: {result.generated_rows}")
    print(f"Planilha: {result.output_path}")
    if result.errors:
        print(f"Arquivos com erro: {len(result.errors)}")
        for error in result.errors:
            print(f"- {error.file_name}: {error.message}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run CLI smoke test**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter.ExtractRowTest.test_cli_script_exists -v
```

Expected: `OK`.

- [ ] **Step 5: Run all tests**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter -v
```

Expected: `OK`.

---

### Task 5: Real Fixture Conversion Verification

**Files:**
- No source changes expected unless verification exposes a bug.

- [ ] **Step 1: Run the converter against the real local JSON folder**

Run:

```powershell
python scripts/converter_json_para_planilha.py --source cnpjs/cnpjs --output dist/importacao_oportunidades.xlsx
```

Expected output pattern:

```text
Arquivos processados: 961
Linhas geradas: 961
Planilha: dist\importacao_oportunidades.xlsx
```

If some fixture JSON is invalid, expected output must also list `Arquivos com erro: N` and each invalid file name.

- [ ] **Step 2: Inspect generated XLSX package**

Run:

```powershell
python -c "import zipfile; z=zipfile.ZipFile('dist/importacao_oportunidades.xlsx'); print('xl/worksheets/sheet1.xml' in z.namelist())"
```

Expected:

```text
True
```

- [ ] **Step 3: Run final tests**

Run:

```powershell
$env:PYTHONPATH='src'; python -m unittest tests.test_converter -v
```

Expected: `OK`.

- [ ] **Step 4: Check repository status**

Run:

```powershell
git status --short
```

Expected: source, test, script, docs and generated `dist/importacao_oportunidades.xlsx` changes are visible. Report all remaining modified/untracked files in the final response. Do not commit unless the user explicitly asks.

---

## Self-Review

- Spec coverage: the plan covers local JSON input, exact approved columns, one row per JSON, list truncation rules, empty cells for missing data, invalid JSON reporting and XLSX generation.
- Placeholder scan: no placeholders remain; fields intentionally empty are explicitly listed in the mapping.
- Type consistency: `extract_row`, `write_xlsx`, `convert_directory`, `ConversionError` and `ConversionResult` are defined before use in later tasks.
