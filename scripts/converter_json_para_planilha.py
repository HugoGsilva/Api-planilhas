from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.converter import convert_directory


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source",
        default="cnpjs/cnpjs",
        help="Diretorio com arquivos JSON de CNPJs",
    )
    parser.add_argument(
        "--output",
        default="dist/importacao_oportunidades.xlsx",
        help="Caminho de saida do XLSX",
    )
    return parser.parse_args(args=argv)


def main(argv: list[str] | None = None) -> int:
    try:
        args = parse_args(argv)
        result = convert_directory(args.source, args.output)
    except (FileNotFoundError, NotADirectoryError, ValueError, OSError) as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1

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
