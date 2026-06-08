from __future__ import annotations

import json
from dataclasses import dataclass
from html import escape
from pathlib import Path
from io import BytesIO
import zipfile
from typing import Any


SPREADSHEET_NS = (
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)


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


@dataclass(frozen=True)
class ConversionError:
    file_name: str
    message: str


@dataclass(frozen=True)
class ConversionResult:
    """Resultado da conversao em lote.

    processed_files: quantidade de arquivos JSON processados com sucesso.
    Apenas arquivos com JSON válido (raiz do payload do tipo dict) contam.
    generated_rows: quantidade de linhas geradas no XLSX (mesma quantidade de
    arquivos processados com sucesso).
    """

    processed_files: int
    generated_rows: int
    errors: tuple[ConversionError, ...]
    output_path: Path


def _as_dict(payload: Any) -> dict[str, Any]:
    if isinstance(payload, dict):
        return payload
    return {}


def _retorno(payload: dict[str, Any]) -> dict[str, Any]:
    retorno = payload.get("retorno")
    return retorno if isinstance(retorno, dict) else {}


def _list_item(data: dict[str, Any], key: str, index: int) -> dict[str, Any]:
    items = data.get(key)
    if not isinstance(items, list) or index < 0 or index >= len(items):
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


def extract_row(payload: Any) -> list[Any]:
    data = _retorno(_as_dict(payload))

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


def _column_name(index: int) -> str:
    if index < 1:
        raise ValueError("index must be >= 1")

    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _sanitize_xml_text(text: str) -> str:
    return "".join(
        char
        for char in text
        if char in {"\t", "\n", "\r"} or ord(char) >= 0x20
    )


def _cell_xml(row: int, column: int, value: Any) -> str:
    ref = f"{_column_name(column)}{row}"
    text = "" if value is None else str(value)
    text = _sanitize_xml_text(text)
    escaped = escape(text, quote=True)

    if text == "":
        return f"<c r=\"{ref}\"/>"
    return (
        f"<c r=\"{ref}\" t=\"inlineStr\"><is><t>{escaped}</t></is></c>"
    )


def _row_xml(row_index: int, cells: list[Any]) -> str:
    cells_xml = "".join(
        _cell_xml(row_index, idx + 1, value) for idx, value in enumerate(cells)
    )
    return f"<row r=\"{row_index}\">{cells_xml}</row>"


def _sheet_xml(rows: list[list[Any]]) -> str:
    rows_xml = "".join(_row_xml(index + 1, row) for index, row in enumerate(rows))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{SPREADSHEET_NS}">'
        f"<sheetData>{rows_xml}</sheetData>"
        "</worksheet>"
    )


def _content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>"
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/_rels/.rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )


def _workbook_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<workbook xmlns="{SPREADSHEET_NS}" '
        f'xmlns:r="{OFFICE_REL_NS}">'
        '<sheets>'
        '<sheet name="importacao_oportunidades" sheetId="1" r:id="rId1"/>'
        "</sheets></workbook>"
    )


def _workbook_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{REL_NS}">'
        f'<Relationship Id="rId1" Type="{OFFICE_REL_NS}/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )


def _package_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{REL_NS}">'
        f'<Relationship Id="rId1" Type="{OFFICE_REL_NS}/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def _write_xlsx_archive(
    target: str | Path | BytesIO, rows: list[list[Any]]
) -> None:
    worksheet_rows = [HEADERS, *rows]

    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types_xml())
        archive.writestr("_rels/.rels", _package_rels_xml())
        archive.writestr("xl/workbook.xml", _workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels_xml())
        archive.writestr("xl/worksheets/sheet1.xml", _sheet_xml(worksheet_rows))


def write_xlsx(output_path: str | Path, rows: list[list[Any]]) -> None:
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    _write_xlsx_archive(output_file, rows)


def build_xlsx_bytes(rows: list[list[Any]]) -> bytes:
    buffer = BytesIO()
    _write_xlsx_archive(buffer, rows)
    return buffer.getvalue()


def convert_directory(
    source_dir: str | Path, output_path: str | Path
) -> ConversionResult:
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
        errors=tuple(errors),
        output_path=output,
    )
