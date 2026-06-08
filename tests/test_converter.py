import sys
import subprocess
import json
import unittest
import zipfile
import xml.etree.ElementTree as ET
from tempfile import TemporaryDirectory
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.converter import (
    HEADERS,
    convert_directory,
    extract_row,
    extract_rows,
    write_xlsx,
)


class ExtractRowTest(unittest.TestCase):
    def test_headers_order_is_exactly_approved(self):
        expected_headers = [
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

        self.assertEqual(HEADERS, expected_headers)

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
                        "nome": "JOAO SOCIO",
                        "documento": "12345678900",
                        "cargo": "SOCIO ADMINISTRADOR",
                    }
                ],
            },
        }

        row = extract_row(payload)

        self.assertEqual(len(row), len(HEADERS))
        self.assertEqual(row[HEADERS.index("Pessoa")], "JOAO SOCIO")
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

    def test_extracts_missing_phones_with_emails_and_socio(self):
        payload = {
            "retorno": {
                "razaoSocial": "SEM TELEFONE LTDA",
                "telefones": [],
                "emails": [
                    {"enderecoEmail": "contato@semsms.com"},
                    {"enderecoEmail": "backup@semsms.com"},
                ],
                "socios": [
                    {
                        "documento": "98765432100",
                        "cargo": "DIRETOR ADMINISTRATIVO",
                    }
                ],
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("Telefone1")], "")
        self.assertEqual(row[HEADERS.index("Telefone2")], "")
        self.assertEqual(row[HEADERS.index("email Prin")], "contato@semsms.com")
        self.assertEqual(row[HEADERS.index("email Secu")], "backup@semsms.com")
        self.assertEqual(row[HEADERS.index("Pessoa")], "")
        self.assertEqual(row[HEADERS.index("Cargo")], "DIRETOR ADMINISTRATIVO")
        self.assertEqual(row[HEADERS.index("CPF")], "98765432100")

    def test_extracts_missing_emails_with_phones(self):
        payload = {
            "retorno": {
                "razaoSocial": "SEM EMAIL LTDA",
                "telefones": [
                    {"telefoneComDDD": "11999887766"},
                ],
                "socios": [
                    {
                        "documento": "11122233344",
                        "cargo": "GESTOR",
                    }
                ],
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("Telefone1")], "11999887766")
        self.assertEqual(row[HEADERS.index("Telefone2")], "")
        self.assertEqual(row[HEADERS.index("email Prin")], "")
        self.assertEqual(row[HEADERS.index("email Secu")], "")
        self.assertEqual(row[HEADERS.index("Pessoa")], "")
        self.assertEqual(row[HEADERS.index("Cargo")], "GESTOR")
        self.assertEqual(row[HEADERS.index("CPF")], "11122233344")

    def test_extracts_missing_socio(self):
        payload = {
            "retorno": {
                "razaoSocial": "SEM SOCIO LTDA",
                "telefones": [
                    {"telefoneComDDD": "11988776655"},
                ],
                "emails": [
                    {"enderecoEmail": "contato@semsocio.com"},
                ],
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("Telefone1")], "11988776655")
        self.assertEqual(row[HEADERS.index("email Prin")], "contato@semsocio.com")
        self.assertEqual(row[HEADERS.index("Pessoa")], "")
        self.assertEqual(row[HEADERS.index("Cargo")], "")
        self.assertEqual(row[HEADERS.index("CPF")], "")

    def test_extracts_socio_name_as_pessoa(self):
        payload = {
            "retorno": {
                "razaoSocial": "EMPRESA COM SOCIO LTDA",
                "socios": [
                    {
                        "nome": "MARIA SOCIA",
                        "documento": "12345678900",
                        "cargo": "SOCIA ADMINISTRADORA",
                    }
                ],
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("Pessoa")], "MARIA SOCIA")
        self.assertEqual(row[HEADERS.index("Cargo")], "SOCIA ADMINISTRADORA")
        self.assertEqual(row[HEADERS.index("CPF")], "12345678900")
        self.assertEqual(row[HEADERS.index("Organizacao")], "EMPRESA COM SOCIO LTDA")

    def test_extract_rows_repeats_company_data_for_each_socio(self):
        payload = {
            "retorno": {
                "cnpj": "00019000000133",
                "razaoSocial": "EMPRESA MULTI SOCIOS LTDA",
                "telefones": [{"telefoneComDDD": "1130000001"}],
                "socios": [
                    {
                        "nome": "SOCIO UM",
                        "documento": "11122233344",
                        "cargo": "SOCIO",
                    },
                    {
                        "nome": "SOCIO DOIS",
                        "documento": "55566677788",
                        "cargo": "ADMINISTRADOR",
                    },
                ],
            }
        }

        rows = extract_rows(payload)

        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][HEADERS.index("Pessoa")], "SOCIO UM")
        self.assertEqual(rows[0][HEADERS.index("CPF")], "11122233344")
        self.assertEqual(rows[0][HEADERS.index("Cargo")], "SOCIO")
        self.assertEqual(rows[1][HEADERS.index("Pessoa")], "SOCIO DOIS")
        self.assertEqual(rows[1][HEADERS.index("CPF")], "55566677788")
        self.assertEqual(rows[1][HEADERS.index("Cargo")], "ADMINISTRADOR")
        for row in rows:
            self.assertEqual(row[HEADERS.index("Organizacao")], "EMPRESA MULTI SOCIOS LTDA")
            self.assertEqual(row[HEADERS.index("CNPJ")], "00019000000133")
            self.assertEqual(row[HEADERS.index("Telefone1")], "1130000001")

    def test_extract_rows_keeps_one_company_row_without_socios(self):
        payload = {
            "retorno": {
                "cnpj": "00019000000133",
                "razaoSocial": "EMPRESA SEM SOCIOS LTDA",
                "socios": [],
            }
        }

        rows = extract_rows(payload)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][HEADERS.index("Pessoa")], "")
        self.assertEqual(rows[0][HEADERS.index("Cargo")], "")
        self.assertEqual(rows[0][HEADERS.index("CPF")], "")
        self.assertEqual(rows[0][HEADERS.index("Organizacao")], "EMPRESA SEM SOCIOS LTDA")

    def test_uses_first_address_when_multiple_addresses(self):
        payload = {
            "retorno": {
                "razaoSocial": "ENDERECO MULTIPLO LTDA",
                "enderecos": [
                    {
                        "cep": "01000001",
                        "logradouro": "Rua Primeiro",
                        "numero": "10",
                        "complemento": "Sala 1",
                        "bairro": "Centro",
                        "cidade": "Sao Paulo",
                        "uf": "SP",
                    },
                    {
                        "cep": "02000002",
                        "logradouro": "Rua Segundo",
                        "numero": "20",
                        "complemento": "Sala 2",
                        "bairro": "Jardim",
                        "cidade": "Sao Paulo",
                        "uf": "SP",
                    },
                ],
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("CEP")], "01000001")
        self.assertEqual(row[HEADERS.index("Logradouro")], "Rua Primeiro")
        self.assertEqual(row[HEADERS.index("Numero")], "10")
        self.assertEqual(row[HEADERS.index("Complemento")], "Sala 1")
        self.assertEqual(row[HEADERS.index("Bairro")], "Centro")
        self.assertEqual(row[HEADERS.index("Cidade")], "Sao Paulo")
        self.assertEqual(row[HEADERS.index("UF")], "SP")

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

    def test_handles_broken_structures_without_exception(self):
        payload: dict[str, Any] = {
            "retorno": "INVALIDA"
        }

        row = extract_row(payload)

        self.assertEqual(len(row), len(HEADERS))
        self.assertTrue(all(cell == "" for cell in row))

    def test_handles_wrong_list_item_types_without_exception(self):
        payload = {
            "retorno": {
                "razaoSocial": "LISTA QUEBRADA LTDA",
                "telefones": ["nao-dict"],
                "emails": [1, 2],
                "enderecos": [3],
                "socios": [4],
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
        self.assertEqual(row[HEADERS.index("Logradouro")], "")

    def test_converts_directory_and_reports_invalid_json(self):
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "cnpjs"
            output = Path(temp_dir) / "saida" / "modelo.xlsx"
            source.mkdir()

            valid_payload = {
                "retorno": {
                    "cnpj": "00019000000133",
                    "razaoSocial": "EMPRESA TESTE LTDA",
                    "telefones": [],
                    "emails": [],
                    "enderecos": [],
                    "socios": [
                        {"nome": "SOCIO UM", "documento": "11122233344"},
                        {"nome": "SOCIO DOIS", "documento": "55566677788"},
                    ],
                }
            }
            (source / "valido.json").write_text(
                json.dumps(valid_payload),
                encoding="utf-8-sig",
            )
            (source / "invalido.json").write_text("{", encoding="utf-8-sig")

            result = convert_directory(source, output)

            self.assertEqual(result.processed_files, 1)
            self.assertEqual(result.generated_rows, 2)
            self.assertEqual(len(result.errors), 1)
            self.assertEqual(result.errors[0].file_name, "invalido.json")
            self.assertTrue(result.output_path == output)
            self.assertTrue(output.exists())

            with zipfile.ZipFile(output) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("EMPRESA TESTE LTDA", sheet_xml)
                self.assertIn("SOCIO UM", sheet_xml)
                self.assertIn("SOCIO DOIS", sheet_xml)

    def test_convert_directory_raises_when_source_does_not_exist(self):
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "cnpjs_ausentes"
            output = Path(temp_dir) / "saida" / "modelo.xlsx"

            with self.assertRaises(FileNotFoundError) as ctx:
                convert_directory(source, output)
            self.assertIn("Diretorio fonte nao encontrado", str(ctx.exception))

    def test_convert_directory_raises_when_source_is_not_a_directory(self):
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "nao_e_diretorio.json"
            output = Path(temp_dir) / "saida" / "modelo.xlsx"
            source.write_text("{}", encoding="utf-8-sig")

            with self.assertRaises(NotADirectoryError) as ctx:
                convert_directory(source, output)
            self.assertIn("Caminho fonte nao e um diretorio", str(ctx.exception))

    def test_convert_directory_empty_source_generates_only_header_row(self):
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "cnpjs"
            output = Path(temp_dir) / "saida" / "modelo.xlsx"
            source.mkdir()

            result = convert_directory(source, output)

            self.assertEqual(result.processed_files, 0)
            self.assertEqual(result.generated_rows, 0)
            self.assertEqual(result.errors, ())
            self.assertTrue(result.output_path == output)
            self.assertTrue(output.exists())

            with zipfile.ZipFile(output) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml")
                worksheet = ET.fromstring(sheet_xml)
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

                header_row = worksheet.find(f".//{{{ns}}}row[@r='1']")
                second_row = worksheet.find(f".//{{{ns}}}row[@r='2']")

                self.assertIsNotNone(header_row)
                self.assertIsNone(second_row)
                cells = header_row.findall(f"./{{{ns}}}c")
                self.assertEqual(len(cells), len(HEADERS))


class CLIEntrypointSmokeTest(unittest.TestCase):
    def test_converter_json_para_planilha_script_exists_and_contains_convert_directory(self):
        script_path = PROJECT_ROOT / "scripts" / "converter_json_para_planilha.py"
        self.assertTrue(
            script_path.is_file(),
            f"Arquivo esperado nao encontrado: {script_path}",
        )
        content = script_path.read_text(encoding="utf-8")
        self.assertIn("convert_directory", content)

    def test_entrypoint_returns_error_when_source_is_missing(self):
        with TemporaryDirectory() as temp_dir:
            script_path = PROJECT_ROOT / "scripts" / "converter_json_para_planilha.py"
            output_path = Path(temp_dir) / "temp_out.xlsx"

            result = subprocess.run(
                [
                    sys.executable,
                    str(script_path),
                    "--source",
                    "nao-existe-dir",
                    "--output",
                    str(output_path),
                ],
                capture_output=True,
                text=True,
                cwd=PROJECT_ROOT,
            )

            self.assertNotEqual(result.returncode, 0)
            combined_output = f"{result.stdout}\n{result.stderr}"
            self.assertIn("Erro:", combined_output)

    def test_entrypoint_runs_successfully_with_valid_directory(self):
        with TemporaryDirectory() as temp_dir:
            script_path = PROJECT_ROOT / "scripts" / "converter_json_para_planilha.py"
            source_dir = Path(temp_dir) / "cnpjs"
            output_path = Path(temp_dir) / "temp_out.xlsx"
            source_dir.mkdir()

            payload = {
                "retorno": {
                    "cnpj": "00019000000133",
                    "razaoSocial": "EMPRESA TESTE LTDA",
                    "telefones": [],
                    "emails": [],
                    "enderecos": [],
                    "socios": [],
                }
            }
            (source_dir / "valido.json").write_text(
                json.dumps(payload),
                encoding="utf-8-sig",
            )

            result = subprocess.run(
                [
                    sys.executable,
                    str(script_path),
                    "--source",
                    str(source_dir),
                    "--output",
                    str(output_path),
                ],
                capture_output=True,
                text=True,
                cwd=PROJECT_ROOT,
            )

            self.assertEqual(result.returncode, 0)
            self.assertTrue(output_path.exists())
            self.assertIn("Arquivos processados: 1", result.stdout)
            self.assertIn("Linhas geradas: 1", result.stdout)
            self.assertIn("Planilha:", result.stdout)
            combined_output = f"{result.stdout}\n{result.stderr}"
            self.assertNotIn("Traceback", combined_output)


class WriteXlsxTest(unittest.TestCase):
    def test_write_xlsx_creates_package_with_headers_and_data(self):
        with TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "saida" / "modelo.xlsx"

            write_xlsx(output, [["Empresa Exemplo LTDA"]])

            self.assertTrue(output.exists())

            with zipfile.ZipFile(output) as archive:
                names = set(archive.namelist())

                self.assertIn("[Content_Types].xml", names)
                self.assertIn("xl/workbook.xml", names)
                self.assertIn("xl/worksheets/sheet1.xml", names)

                workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
                self.assertIn(
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
                    workbook_xml,
                )

                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("Pessoa", sheet_xml)
                self.assertIn("Empresa Exemplo LTDA", sheet_xml)

    def test_write_xlsx_writes_numbers_as_inline_string_and_escapes_xml(self):
        with TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "saida" / "modelo.xlsx"

            write_xlsx(
                output,
                [
                    [
                        12345,
                        "EMPRESA & <TESTE>",
                        "Linha com controle:\x00final",
                    ]
                ],
            )

            self.assertTrue(output.exists())

            with zipfile.ZipFile(output) as archive:
                sheet_xml_bytes = archive.read("xl/worksheets/sheet1.xml")
                sheet_xml = sheet_xml_bytes.decode("utf-8")

                self.assertNotIn(' t="n"', sheet_xml)
                self.assertIn(
                    "<c r=\"A2\" t=\"inlineStr\"><is><t>12345</t></is></c>",
                    sheet_xml,
                )
                self.assertIn("EMPRESA &amp; &lt;TESTE&gt;", sheet_xml)

                worksheet = ET.fromstring(sheet_xml_bytes)
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

                cell_a2 = worksheet.find(f".//{{{ns}}}c[@r='A2']")
                cell_b2 = worksheet.find(f".//{{{ns}}}c[@r='B2']")
                cell_c2 = worksheet.find(f".//{{{ns}}}c[@r='C2']")

                self.assertIsNotNone(cell_a2)
                self.assertIsNotNone(cell_b2)
                self.assertIsNotNone(cell_c2)

                self.assertEqual(cell_a2.attrib.get("t"), "inlineStr")
                text_a2 = cell_a2.find(f"./{{{ns}}}is/{{{ns}}}t")
                self.assertEqual(text_a2.text, "12345")

                text_b2 = cell_b2.find(f"./{{{ns}}}is/{{{ns}}}t")
                self.assertEqual(text_b2.text, "EMPRESA & <TESTE>")

                text_c2 = cell_c2.find(f"./{{{ns}}}is/{{{ns}}}t")
                self.assertEqual(text_c2.text, "Linha com controle:final")

    def test_build_xlsx_bytes_returns_valid_package(self):
        import zipfile
        from io import BytesIO

        from api_planilhas.converter import build_xlsx_bytes

        content = build_xlsx_bytes(
            [["EMPRESA BYTES LTDA"] + [""] * (len(HEADERS) - 1)]
        )

        self.assertIsInstance(content, bytes)
        with zipfile.ZipFile(BytesIO(content)) as archive:
            self.assertIn("xl/worksheets/sheet1.xml", archive.namelist())
            sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("Pessoa", sheet_xml)
            self.assertIn("EMPRESA BYTES LTDA", sheet_xml)

if __name__ == "__main__":
    unittest.main()
