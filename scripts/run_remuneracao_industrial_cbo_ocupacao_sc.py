from __future__ import annotations

import argparse
import logging
import unicodedata
from pathlib import Path

import pandas as pd


LOGGER = logging.getLogger("remuneracao_industrial_cbo_ocupacao_sc")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "raw" / "2024_vinc_sul" / "extracted" / "RAIS_VINC_PUB_SUL.COMT"
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
MUNICIPALITY_REFERENCE_FILE = PROJECT_ROOT / "data" / "reference" / "municipios_sc_mesorregioes.csv"
CBO_DICTIONARY_FILE = PROJECT_ROOT / "data" / "dict" / "dicionario_cbo.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
MAIN_BASENAME = "remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024"
ABDON_BASENAME = "remuneracao_media_ano_industrial_abdon_batista_cbo_ocupacao_2024"
CHUNK_SIZE = 500_000
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLS = 16_384


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def normalize_text(value: object) -> str:
    normalized = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return normalized.lower().strip()


def clean_digits(value: object, width: int | None = None) -> str | None:
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        formatted_value = format(value, "f").rstrip("0").rstrip(".")
    else:
        formatted_value = str(value)

    digits = "".join(character for character in formatted_value if character.isdigit())
    if not digits:
        return None

    if width is not None:
        return digits.zfill(width)
    return digits


def detect_separator(file_path: Path) -> str:
    with file_path.open("r", encoding="latin-1", newline="") as file_handle:
        sample = file_handle.readline()
    return ";" if sample.count(";") > sample.count(",") else ","


def resolve_columns(file_path: Path, separator: str) -> dict[str, str]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    required_columns = {
        "municipio_codigo": "municipio - codigo",
        "cbo_ocupacao_codigo": "cbo 2002 ocupacao - codigo",
        "cnae_subclasse_codigo": "cnae 2.0 subclasse - codigo",
        "vl_rem_media_nom": "vl rem media nom",
        "ind_vinculo_abandonado": "ind vinculo abandonado - codigo",
        "ind_vinculo_ativo_3112": "ind vinculo ativo 31/12 - codigo",
    }

    resolved: dict[str, str] = {}
    for alias, normalized_name in required_columns.items():
        matched_column = normalized_columns.get(normalized_name)
        if matched_column is None:
            raise ValueError(f"Coluna obrigatoria nao encontrada: {normalized_name}")
        resolved[alias] = matched_column

    return resolved


def load_cnae_dimension(path: Path) -> pd.DataFrame:
    dimension = pd.read_excel(path, sheet_name="CNAE-SUB")
    normalized_columns = {normalize_text(column): column for column in dimension.columns}
    dimension = dimension.rename(
        columns={
            normalized_columns["cod_div"]: "cnae_divisao_codigo",
            normalized_columns["cod_subcsp"]: "cnae_subclasse_codigo",
            next(column for key, column in normalized_columns.items() if key.startswith("divis")): "cnae_divisao_nome",
            normalized_columns["subclasse"]: "cnae_subclasse_nome",
        }
    )
    dimension["cnae_subclasse_codigo"] = dimension["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )
    dimension["cnae_divisao_codigo"] = dimension["cnae_divisao_codigo"].apply(
        lambda value: clean_digits(value, width=2)
    )
    return (
        dimension[
            [
                "cnae_subclasse_codigo",
                "cnae_subclasse_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ]
        ]
        .dropna(subset=["cnae_subclasse_codigo", "cnae_divisao_codigo"])
        .drop_duplicates(subset=["cnae_subclasse_codigo"])
        .reset_index(drop=True)
    )


def load_municipality_reference(path: Path) -> pd.DataFrame:
    reference = pd.read_csv(path)
    reference["municipio_codigo"] = pd.to_numeric(reference["municipio_codigo"], errors="coerce").astype("Int64")
    return reference[["municipio_codigo", "municipio_nome", "mesorregiao_nome"]].drop_duplicates(
        subset=["municipio_codigo"]
    )


def load_cbo_dictionary(path: Path) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)
    normalized_sheets = {normalize_text(sheet_name): sheet_name for sheet_name in workbook.sheet_names}
    sheet_name = normalized_sheets.get("ocupacao", workbook.sheet_names[0])
    dictionary = pd.read_excel(path, sheet_name=sheet_name)
    value_column = dictionary.columns[0]
    parsed = dictionary[value_column].astype(str).str.split(":", n=1, expand=True)
    dictionary = pd.DataFrame(
        {
            "cbo_ocupacao_codigo": parsed[0].apply(lambda value: clean_digits(value, width=6)),
            "cbo_ocupacao_nome": parsed[1].fillna("").astype(str).str.strip(),
        }
    )
    return (
        dictionary.loc[dictionary["cbo_ocupacao_codigo"].notna()]
        .drop_duplicates(subset=["cbo_ocupacao_codigo"])
        .reset_index(drop=True)
    )


def aggregate_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
) -> pd.DataFrame:
    chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce").astype("Int64")
    chunk["vl_rem_media_nom"] = pd.to_numeric(chunk["vl_rem_media_nom"], errors="coerce")
    chunk["ind_vinculo_abandonado"] = pd.to_numeric(chunk["ind_vinculo_abandonado"], errors="coerce")
    chunk["ind_vinculo_ativo_3112"] = pd.to_numeric(chunk["ind_vinculo_ativo_3112"], errors="coerce")
    chunk["cbo_ocupacao_codigo"] = chunk["cbo_ocupacao_codigo"].apply(lambda value: clean_digits(value, width=6))
    chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )

    filtered = chunk.loc[
        chunk["municipio_codigo"].astype(str).str.startswith("42")
        & chunk["ind_vinculo_abandonado"].eq(0)
        & chunk["ind_vinculo_ativo_3112"].eq(1)
        & chunk["vl_rem_media_nom"].gt(0)
    ].copy()

    if filtered.empty:
        return filtered

    enriched = filtered.merge(cnae_dimension, on="cnae_subclasse_codigo", how="left")
    enriched = enriched.merge(municipality_reference, on="municipio_codigo", how="left")
    enriched = enriched.loc[enriched["cnae_divisao_codigo"].between("05", "43")].copy()

    if enriched.empty:
        return enriched

    enriched["municipio_nome"] = enriched["municipio_nome"].fillna("Municipio nao identificado")
    enriched["mesorregiao_nome"] = enriched["mesorregiao_nome"].fillna("Mesorregiao nao identificada")
    enriched["cnae_divisao_nome"] = enriched["cnae_divisao_nome"].fillna("Divisao nao identificada")
    enriched["cnae_subclasse_nome"] = enriched["cnae_subclasse_nome"].fillna("Subclasse nao identificada")
    enriched["qtd_vinculos"] = 1

    return (
        enriched.groupby(
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cbo_ocupacao_codigo",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
                "cnae_subclasse_codigo",
                "cnae_subclasse_nome",
            ],
            as_index=False,
        )
        .agg(
            remuneracao_soma=("vl_rem_media_nom", "sum"),
            qtd_vinculos=("qtd_vinculos", "sum"),
        )
    )


def build_final_dataframe(aggregated_chunks: list[pd.DataFrame], cbo_dictionary: pd.DataFrame) -> pd.DataFrame:
    combined = pd.concat(aggregated_chunks, ignore_index=True)
    final = (
        combined.groupby(
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cbo_ocupacao_codigo",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
                "cnae_subclasse_codigo",
                "cnae_subclasse_nome",
            ],
            as_index=False,
        )[["remuneracao_soma", "qtd_vinculos"]]
        .sum()
    )
    final["ano_referencia"] = 2024
    final["remuneracao_media_nom_ano"] = final["remuneracao_soma"] / final["qtd_vinculos"]
    final = final.merge(cbo_dictionary, on="cbo_ocupacao_codigo", how="left")
    final["cbo_ocupacao_nome"] = final["cbo_ocupacao_nome"].fillna("CBO nao identificado")
    return final[
        [
            "ano_referencia",
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "cbo_ocupacao_codigo",
            "cbo_ocupacao_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
            "cnae_subclasse_codigo",
            "cnae_subclasse_nome",
            "qtd_vinculos",
            "remuneracao_media_nom_ano",
        ]
    ].sort_values(
        [
            "municipio_nome",
            "cbo_ocupacao_codigo",
            "cnae_subclasse_codigo",
        ]
    ).reset_index(drop=True)


def auto_fit_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))

    for index, column in enumerate(dataframe.columns):
        max_length = max(
            len(str(column)),
            dataframe[column].astype(str).map(len).max() if not dataframe.empty else 0,
        )
        worksheet.set_column(index, index, min(max_length + 2, 40))

    workbook = writer.book
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    integer_format = workbook.add_format({"num_format": "#,##0"})

    if "remuneracao_media_nom_ano" in dataframe.columns:
        column_index = dataframe.columns.get_loc("remuneracao_media_nom_ano")
        worksheet.set_column(column_index, column_index, 20, money_format)

    if "qtd_vinculos" in dataframe.columns:
        column_index = dataframe.columns.get_loc("qtd_vinculos")
        worksheet.set_column(column_index, column_index, 14, integer_format)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera a remuneracao media anual nominal da RAIS industrial para SC por CBO ocupacao."
    )
    parser.add_argument("--input-path", default=str(INPUT_FILE), help="Caminho do arquivo RAIS de vinculos.")
    parser.add_argument(
        "--cnae-dimension-path",
        default=str(CNAE_DIMENSION_FILE),
        help="Caminho da planilha de dimensao CNAE.",
    )
    parser.add_argument(
        "--municipality-reference-path",
        default=str(MUNICIPALITY_REFERENCE_FILE),
        help="Caminho do arquivo de referencia territorial.",
    )
    parser.add_argument(
        "--cbo-dictionary-path",
        default=str(CBO_DICTIONARY_FILE),
        help="Caminho do dicionario de CBO.",
    )
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="Diretorio base das saidas.")
    parser.add_argument(
        "--main-output-path",
        default=None,
        help="Caminho do arquivo principal. Se omitido, usa o nome padrao no diretorio de saida.",
    )
    parser.add_argument(
        "--abdon-output-path",
        default=None,
        help="Caminho do arquivo de Abdon Batista. Se omitido, usa o nome padrao no diretorio de saida.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=CHUNK_SIZE,
        help="Quantidade de linhas por chunk durante a leitura.",
    )
    return parser


def export_main_file(dataframe: pd.DataFrame, output_dir: Path, main_output_path: Path | None) -> Path:
    csv_path = output_dir / f"{MAIN_BASENAME}.csv"
    xlsx_path = main_output_path or (output_dir / f"{MAIN_BASENAME}.xlsx")

    if len(dataframe) > EXCEL_MAX_ROWS or len(dataframe.columns) > EXCEL_MAX_COLS:
        dataframe.to_csv(csv_path, index=False, encoding="utf-8-sig")
        LOGGER.info("Arquivo principal gerado em CSV: %s", csv_path)
        return csv_path

    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, sheet_name="Industrial_CBO_Ocupacao", index=False)
        auto_fit_columns(writer, "Industrial_CBO_Ocupacao", dataframe)
    LOGGER.info("Arquivo principal gerado em Excel: %s", xlsx_path)
    return xlsx_path


def export_abdon_batista_file(dataframe: pd.DataFrame, output_dir: Path, abdon_output_path: Path | None) -> Path:
    abdon = dataframe.loc[dataframe["municipio_codigo"].eq(420005)].copy()
    xlsx_path = abdon_output_path or (output_dir / f"{ABDON_BASENAME}.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        abdon.to_excel(writer, sheet_name="Abdon_Batista", index=False)
        auto_fit_columns(writer, "Abdon_Batista", abdon)
    LOGGER.info("Arquivo de Abdon Batista gerado em %s", xlsx_path)
    return xlsx_path


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()
    input_file = Path(args.input_path)
    cnae_dimension_file = Path(args.cnae_dimension_path)
    municipality_reference_file = Path(args.municipality_reference_path)
    cbo_dictionary_file = Path(args.cbo_dictionary_path)
    output_dir = Path(args.output_dir)
    main_output_path = Path(args.main_output_path) if args.main_output_path else None
    abdon_output_path = Path(args.abdon_output_path) if args.abdon_output_path else None
    if not input_file.exists():
        raise FileNotFoundError(f"Arquivo de entrada nao encontrado: {input_file}")

    output_dir.mkdir(parents=True, exist_ok=True)
    if main_output_path is not None:
        main_output_path.parent.mkdir(parents=True, exist_ok=True)
    if abdon_output_path is not None:
        abdon_output_path.parent.mkdir(parents=True, exist_ok=True)

    separator = detect_separator(input_file)
    resolved_columns = resolve_columns(input_file, separator)
    cnae_dimension = load_cnae_dimension(cnae_dimension_file)
    municipality_reference = load_municipality_reference(municipality_reference_file)
    cbo_dictionary = load_cbo_dictionary(cbo_dictionary_file)

    LOGGER.info("Lendo %s com separador '%s'", input_file.name, separator)
    reader = pd.read_csv(
        input_file,
        encoding="latin-1",
        sep=separator,
        usecols=list(resolved_columns.values()),
        chunksize=args.chunk_size,
        low_memory=False,
    )

    aggregated_chunks: list[pd.DataFrame] = []
    for chunk_number, chunk in enumerate(reader, start=1):
        renamed_chunk = chunk.rename(columns={source: target for target, source in resolved_columns.items()})
        aggregated = aggregate_chunk(renamed_chunk, cnae_dimension, municipality_reference)
        if not aggregated.empty:
            aggregated_chunks.append(aggregated)
        LOGGER.info("Chunk %s processado", chunk_number)

    if not aggregated_chunks:
        raise ValueError("Nenhum registro valido foi encontrado apos aplicar os filtros.")

    final = build_final_dataframe(aggregated_chunks, cbo_dictionary)
    export_main_file(final, output_dir, main_output_path)
    export_abdon_batista_file(final, output_dir, abdon_output_path)


if __name__ == "__main__":
    main()
