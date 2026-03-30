from __future__ import annotations

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
EXAMPLES_DIR = PROJECT_ROOT / "examples"


def auto_fit_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))

    workbook = writer.book
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    integer_format = workbook.add_format({"num_format": "#,##0"})

    for index, column in enumerate(dataframe.columns):
        max_length = max(
            len(str(column)),
            dataframe[column].astype(str).map(len).max() if not dataframe.empty else 0,
        )
        worksheet.set_column(index, index, min(max_length + 2, 40))

    for numeric_column in ["qtd_vinculos", "valor"]:
        if numeric_column in dataframe.columns:
            column_index = dataframe.columns.get_loc(numeric_column)
            worksheet.set_column(column_index, column_index, 16, integer_format)

    for money_column in ["remuneracao_media_nom_ano"]:
        if money_column in dataframe.columns:
            column_index = dataframe.columns.get_loc(money_column)
            worksheet.set_column(column_index, column_index, 20, money_format)


def main() -> None:
    examples_dir = EXAMPLES_DIR
    examples_dir.mkdir(parents=True, exist_ok=True)

    remuneracao_csv = examples_dir / "sample_remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.csv"
    validacao_csv = examples_dir / "sample_validacao.csv"

    remuneracao_df = pd.read_csv(remuneracao_csv)
    validacao_df = pd.read_csv(validacao_csv)

    remuneracao_xlsx = examples_dir / "sample_remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx"
    validacao_xlsx = examples_dir / "sample_validacao.xlsx"

    with pd.ExcelWriter(remuneracao_xlsx, engine="xlsxwriter") as writer:
        remuneracao_df.to_excel(writer, sheet_name="Exemplo_Industrial_SC", index=False)
        auto_fit_columns(writer, "Exemplo_Industrial_SC", remuneracao_df)

    with pd.ExcelWriter(validacao_xlsx, engine="xlsxwriter") as writer:
        validacao_df.to_excel(writer, sheet_name="Exemplo_Validacao", index=False)
        auto_fit_columns(writer, "Exemplo_Validacao", validacao_df)


if __name__ == "__main__":
    main()
