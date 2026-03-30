from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "output" / "remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx"
OUTPUT_FILE = PROJECT_ROOT / "data" / "output" / "validacao.xlsx"


def auto_fit_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))

    for index, column in enumerate(dataframe.columns):
        max_length = max(
            len(str(column)),
            dataframe[column].astype(str).map(len).max() if not dataframe.empty else 0,
        )
        worksheet.set_column(index, index, min(max_length + 2, 32))

    workbook = writer.book
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    integer_format = workbook.add_format({"num_format": "#,##0"})

    worksheet.set_column(1, 1, 18, money_format)
    worksheet.set_column(2, 2, 14, integer_format)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera um arquivo de validacao a partir da planilha industrial consolidada."
    )
    parser.add_argument(
        "--input-path",
        default=str(INPUT_FILE),
        help="Caminho do arquivo consolidado de remuneracao industrial.",
    )
    parser.add_argument(
        "--output-path",
        default=str(OUTPUT_FILE),
        help="Caminho de saida do arquivo de validacao.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    input_path = Path(args.input_path)
    output_path = Path(args.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    dataframe = pd.read_excel(input_path)

    total_vinculos = int(dataframe["qtd_vinculos"].sum())
    media_salarial_geral = (
        (dataframe["remuneracao_media_nom_ano"] * dataframe["qtd_vinculos"]).sum() / total_vinculos
    )

    validation = pd.DataFrame(
        [
            {
                "indicador": "municipios_unicos",
                "valor": int(dataframe["municipio_codigo"].nunique()),
                "observacao": "Quantidade de municipios unicos no arquivo geral SC.",
            },
            {
                "indicador": "total_vinculos_geral",
                "valor": total_vinculos,
                "observacao": "Soma da coluna qtd_vinculos.",
            },
            {
                "indicador": "media_salarial_geral",
                "valor": float(media_salarial_geral),
                "observacao": "Media ponderada por qtd_vinculos da coluna remuneracao_media_nom_ano.",
            },
            {
                "indicador": "mesorregioes_unicas",
                "valor": int(dataframe["mesorregiao_nome"].nunique()),
                "observacao": "Quantidade de mesorregioes unicas no arquivo geral SC.",
            },
        ]
    )

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        validation.to_excel(writer, sheet_name="Validacao", index=False)
        auto_fit_columns(writer, "Validacao", validation)


if __name__ == "__main__":
    main()
