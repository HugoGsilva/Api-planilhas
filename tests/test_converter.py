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
            "cnpj",
            "razaoSocial",
            "nomeFantasia",
            "dataFundacao",
            "cnaeCodigo",
            "cnaeDescricao",
            "cnaEsSecundarios",
            "quantidadeFuncionarios",
            "situacaoCadastral",
            "naturezaJuridicaCodigo",
            "naturezaJuridicaDescricao",
            "naturezaJuridicaTipo",
            "porte",
            "faixaFuncionarios",
            "faixaFaturamento",
            "matriz",
            "orgaoPublico",
            "ramo",
            "tipoEmpresa",
            "telefones",
            "enderecos",
            "emails",
            "ultimaAtualizacaoPJ",
            "socios",
            "Nome Socio",
            "CPF Socio",
            "Telefones",
        ]

        self.assertEqual(HEADERS, expected_headers)

    def test_extracts_complete_json_to_csv_model_columns(self):
        payload = {
            "metaDados": {"resultado": "OK"},
            "retorno": {
                "cnpj": "00019000000133",
                "razaoSocial": "EMPRESA TESTE LTDA",
                "nomeFantasia": "EMPRESA TESTE",
                "cnaeCodigo": 6201501,
                "cnaeDescricao": "Desenvolvimento de programas de computador sob encomenda",
                "cnaEsSecundarios": [
                    {
                        "cnaeCodigo": 7112000,
                        "cnaeDescricao": "Servicos de engenharia",
                    }
                ],
                "quantidadeFuncionarios": 42,
                "situacaoCadastral": "ATIVA",
                "naturezaJuridicaCodigo": 2062,
                "naturezaJuridicaDescricao": "Sociedade Empresaria Limitada",
                "naturezaJuridicaTipo": "Privada",
                "porte": "MICRO EMPRESA",
                "faixaFuncionarios": "Ate 9 Funcionarios",
                "faixaFaturamento": "Ate R$ 240,0 Mil",
                "matriz": True,
                "orgaoPublico": "Nao",
                "ramo": "Engenharia",
                "tipoEmpresa": "LTDA",
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
                    },
                    {
                        "nome": "MARIA SOCIA",
                        "documento": "98765432100",
                        "cargo": "SOCIA",
                    }
                ],
                "ultimaAtualizacaoPJ": "01/01/2026 10:00:00",
            },
        }

        row = extract_row(payload)

        self.assertEqual(len(row), len(HEADERS))
        self.assertEqual(row[HEADERS.index("cnpj")], "00019000000133")
        self.assertEqual(row[HEADERS.index("razaoSocial")], "EMPRESA TESTE LTDA")
        self.assertEqual(row[HEADERS.index("nomeFantasia")], "EMPRESA TESTE")
        self.assertEqual(row[HEADERS.index("dataFundacao")], "")
        self.assertEqual(row[HEADERS.index("cnaeCodigo")], 6201501)
        self.assertEqual(
            row[HEADERS.index("cnaeDescricao")],
            "Desenvolvimento de programas de computador sob encomenda",
        )
        self.assertEqual(
            row[HEADERS.index("cnaEsSecundarios")],
            "7112000 - Servicos de engenharia",
        )
        self.assertEqual(row[HEADERS.index("quantidadeFuncionarios")], 42)
        self.assertEqual(row[HEADERS.index("situacaoCadastral")], "ATIVA")
        self.assertEqual(row[HEADERS.index("naturezaJuridicaCodigo")], 2062)
        self.assertEqual(
            row[HEADERS.index("naturezaJuridicaDescricao")],
            "Sociedade Empresaria Limitada",
        )
        self.assertEqual(row[HEADERS.index("naturezaJuridicaTipo")], "Privada")
        self.assertEqual(row[HEADERS.index("porte")], "MICRO EMPRESA")
        self.assertEqual(row[HEADERS.index("faixaFuncionarios")], "Ate 9 Funcionarios")
        self.assertEqual(row[HEADERS.index("faixaFaturamento")], "Ate R$ 240,0 Mil")
        self.assertEqual(row[HEADERS.index("matriz")], "Sim")
        self.assertEqual(row[HEADERS.index("orgaoPublico")], "Nao")
        self.assertEqual(row[HEADERS.index("ramo")], "Engenharia")
        self.assertEqual(row[HEADERS.index("tipoEmpresa")], "LTDA")
        self.assertEqual(
            row[HEADERS.index("telefones")],
            "1130000001, 1130000002, 1130000003",
        )
        self.assertEqual(
            row[HEADERS.index("enderecos")],
            "PRACA DA SE, 100, CJ 10, SE, SAO PAULO, SP, 01001000",
        )
        self.assertEqual(
            row[HEADERS.index("emails")],
            "principal@example.com, secundario@example.com, terceiro@example.com",
        )
        self.assertEqual(row[HEADERS.index("ultimaAtualizacaoPJ")], "01/01/2026 10:00:00")
        self.assertEqual(
            row[HEADERS.index("socios")],
            "JOAO SOCIO (SOCIO ADMINISTRADOR), MARIA SOCIA (SOCIA)",
        )
        self.assertEqual(row[HEADERS.index("Nome Socio")], "JOAO SOCIO")
        self.assertEqual(row[HEADERS.index("CPF Socio")], "")
        self.assertEqual(row[HEADERS.index("Telefones")], "")

    def test_extracts_one_row_per_socio_with_cpf_phones(self):
        payload = {
            "retorno": {
                "cnpj": "00019000000133",
                "razaoSocial": "EMPRESA TESTE LTDA",
                "socios": [
                    {
                        "nome": "JOAO SOCIO",
                        "documento": "111.111.111-11",
                    },
                    {
                        "nome": "MARIA SOCIA",
                        "documento": "222.222.222-22",
                    },
                ],
            }
        }

        rows = extract_rows(payload)

        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][HEADERS.index("cnpj")], "00019000000133")
        self.assertEqual(rows[1][HEADERS.index("cnpj")], "00019000000133")
        self.assertEqual(rows[0][HEADERS.index("Nome Socio")], "JOAO SOCIO")
        self.assertEqual(rows[0][HEADERS.index("CPF Socio")], "")
        self.assertEqual(rows[0][HEADERS.index("Telefones")], "")
        self.assertEqual(rows[1][HEADERS.index("Nome Socio")], "MARIA SOCIA")
        self.assertEqual(rows[1][HEADERS.index("CPF Socio")], "")
        self.assertEqual(rows[1][HEADERS.index("Telefones")], "")

    def test_extracts_empty_cells_when_optional_json_fields_are_missing(self):
        payload = {
            "retorno": {
                "cnpj": "00000000000000",
                "razaoSocial": "SEM LISTAS LTDA",
                "telefones": [],
                "emails": [],
                "enderecos": [],
                "socios": [],
            }
        }

        row = extract_row(payload)

        self.assertEqual(row[HEADERS.index("cnpj")], "00000000000000")
        self.assertEqual(row[HEADERS.index("razaoSocial")], "SEM LISTAS LTDA")
        self.assertEqual(row[HEADERS.index("telefones")], "")
        self.assertEqual(row[HEADERS.index("emails")], "")
        self.assertEqual(row[HEADERS.index("enderecos")], "")
        self.assertEqual(row[HEADERS.index("socios")], "")
        self.assertEqual(row[HEADERS.index("Nome Socio")], "")
        self.assertEqual(row[HEADERS.index("CPF Socio")], "")
        self.assertEqual(row[HEADERS.index("Telefones")], "")
        self.assertEqual(row[HEADERS.index("matriz")], "")

    def test_joins_multiple_addresses_with_pipe_separator(self):
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

        self.assertEqual(
            row[HEADERS.index("enderecos")],
            (
                "Rua Primeiro, 10, Sala 1, Centro, Sao Paulo, SP, 01000001"
                " | Rua Segundo, 20, Sala 2, Jardim, Sao Paulo, SP, 02000002"
            ),
        )

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

        self.assertEqual(row[HEADERS.index("telefones")], "")
        self.assertEqual(row[HEADERS.index("emails")], "")
        self.assertEqual(row[HEADERS.index("enderecos")], "")
        self.assertEqual(row[HEADERS.index("socios")], "")
        self.assertEqual(row[HEADERS.index("Nome Socio")], "")
        self.assertEqual(row[HEADERS.index("CPF Socio")], "")
        self.assertEqual(row[HEADERS.index("Telefones")], "")

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
                self.assertIn("cnpj", sheet_xml)
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
            self.assertIn("cnpj", sheet_xml)
            self.assertIn("EMPRESA BYTES LTDA", sheet_xml)

if __name__ == "__main__":
    unittest.main()
