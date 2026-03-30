from __future__ import annotations

import argparse
import logging
import unicodedata
from pathlib import Path

import pandas as pd


LOGGER = logging.getLogger("remuneracao_media_sc")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "raw" / "2024_vinc_sul" / "extracted" / "RAIS_VINC_PUB_SUL.COMT"
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
MUNICIPALITY_REFERENCE_FILE = PROJECT_ROOT / "data" / "reference" / "municipios_sc_mesorregioes.csv"
OUTPUT_FILE = PROJECT_ROOT / "data" / "output" / "remuneracao_media_sc.xlsx"
CHUNK_SIZE = 500_000

CBO_GRANDE_GRUPO_NOMES = {
    "0": "Forças Armadas, Policiais e Bombeiros Militares",
    "1": "Membros superiores do poder público, dirigentes e gerentes",
    "2": "Profissionais das ciências e das artes",
    "3": "Técnicos de nível médio",
    "4": "Trabalhadores de serviços administrativos",
    "5": "Trabalhadores dos serviços e vendedores do comércio",
    "6": "Trabalhadores agropecuários, florestais, da caça e pesca",
    "7": "Trabalhadores da produção de bens e serviços industriais",
    "8": "Trabalhadores da produção de bens e serviços industriais",
    "9": "Trabalhadores de manutenção e reparação",
}


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
        "vl_rem_dezembro_nom": "vl rem dezembro nom",
        "ind_vinculo_abandonado": "ind vinculo abandonado - codigo",
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
    dimension = dimension.rename(
        columns={
            "COD_DIV": "cnae_divisao_codigo",
            "Divisão": "cnae_divisao_nome",
            "COD_SUBCsp": "cnae_subclasse_codigo",
        }
    )
    dimension["cnae_subclasse_codigo"] = dimension["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )
    dimension["cnae_divisao_codigo"] = dimension["cnae_divisao_codigo"].apply(
        lambda value: clean_digits(value, width=2)
    )
    return (
        dimension[["cnae_subclasse_codigo", "cnae_divisao_codigo", "cnae_divisao_nome"]]
        .dropna(subset=["cnae_subclasse_codigo"])
        .drop_duplicates(subset=["cnae_subclasse_codigo"])
        .reset_index(drop=True)
    )


def load_municipality_reference(path: Path) -> pd.DataFrame:
    reference = pd.read_csv(path)
    reference["municipio_codigo"] = pd.to_numeric(reference["municipio_codigo"], errors="coerce").astype("Int64")
    return reference[["municipio_codigo", "municipio_nome", "mesorregiao_nome"]].drop_duplicates(
        subset=["municipio_codigo"]
    )


def aggregate_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
    active_3112_only: bool,
) -> pd.DataFrame:
    chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce").astype("Int64")
    chunk["vl_rem_dezembro_nom"] = pd.to_numeric(chunk["vl_rem_dezembro_nom"], errors="coerce")
    chunk["ind_vinculo_abandonado"] = pd.to_numeric(chunk["ind_vinculo_abandonado"], errors="coerce")
    chunk["cbo_ocupacao_codigo"] = chunk["cbo_ocupacao_codigo"].apply(lambda value: clean_digits(value, width=6))
    chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )

    filtered = chunk.loc[
        chunk["municipio_codigo"].astype(str).str.startswith("42")
        & chunk["ind_vinculo_abandonado"].eq(0)
        & chunk["vl_rem_dezembro_nom"].gt(0)
    ].copy()

    if active_3112_only:
        filtered = filtered.loc[filtered["ind_vinculo_ativo_3112"].eq(1)].copy()

    if filtered.empty:
        return filtered

    filtered["cbo_grande_grupo_codigo"] = filtered["cbo_ocupacao_codigo"].str[0].fillna("NA")
    filtered["cbo_grande_grupo_nome"] = filtered["cbo_grande_grupo_codigo"].map(CBO_GRANDE_GRUPO_NOMES).fillna(
        "Grande grupo nao identificado"
    )

    enriched = filtered.merge(cnae_dimension, on="cnae_subclasse_codigo", how="left")
    enriched = enriched.merge(municipality_reference, on="municipio_codigo", how="left")

    enriched["municipio_nome"] = enriched["municipio_nome"].fillna("Municipio nao identificado")
    enriched["mesorregiao_nome"] = enriched["mesorregiao_nome"].fillna("Mesorregiao nao identificada")
    enriched["cnae_divisao_codigo"] = enriched["cnae_divisao_codigo"].fillna("NA")
    enriched["cnae_divisao_nome"] = enriched["cnae_divisao_nome"].fillna("Divisao nao identificada")
    enriched["qtd_vinculos"] = 1

    return (
        enriched.groupby(
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cbo_grande_grupo_codigo",
                "cbo_grande_grupo_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )
        .agg(
            remuneracao_soma=("vl_rem_dezembro_nom", "sum"),
            qtd_vinculos=("qtd_vinculos", "sum"),
        )
    )


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

    if "remuneracao_media_dezembro_nom" in dataframe.columns:
        column_index = dataframe.columns.get_loc("remuneracao_media_dezembro_nom")
        worksheet.set_column(column_index, column_index, 20, money_format)

    if "qtd_vinculos" in dataframe.columns:
        column_index = dataframe.columns.get_loc("qtd_vinculos")
        worksheet.set_column(column_index, column_index, 14, integer_format)


def resolve_output_path(path: Path) -> Path:
    if not path.exists():
        return path

    try:
        with path.open("ab"):
            return path
    except PermissionError:
        return path.with_stem(f"{path.stem} corrigido")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Gera a remuneracao media de dezembro da RAIS 2024 para SC.")
    parser.add_argument(
        "--input-path",
        default=str(INPUT_FILE),
        help="Caminho do arquivo RAIS de vinculos a ser processado.",
    )
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
        "--active-3112-only",
        action="store_true",
        help="Aplica tambem o filtro Ind Vinculo Ativo 31/12 = 1.",
    )
    parser.add_argument(
        "--output-path",
        default=None,
        help="Caminho de saida do arquivo Excel.",
    )
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()
    input_file = Path(args.input_path)
    cnae_dimension_file = Path(args.cnae_dimension_path)
    municipality_reference_file = Path(args.municipality_reference_path)
    if not input_file.exists():
        raise FileNotFoundError(f"Arquivo de entrada nao encontrado: {input_file}")

    output_file = Path(args.output_path) if args.output_path else OUTPUT_FILE
    output_file.parent.mkdir(parents=True, exist_ok=True)

    separator = detect_separator(input_file)
    resolved_columns = resolve_columns(input_file, separator)
    header = pd.read_csv(input_file, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    active_column = normalized_columns.get("ind vinculo ativo 31/12 - codigo")
    if active_column is None:
        raise ValueError("Coluna obrigatoria nao encontrada: ind vinculo ativo 31/12 - codigo")
    resolved_columns["ind_vinculo_ativo_3112"] = active_column
    cnae_dimension = load_cnae_dimension(cnae_dimension_file)
    municipality_reference = load_municipality_reference(municipality_reference_file)

    LOGGER.info("Lendo %s com separador '%s'", input_file.name, separator)
    reader = pd.read_csv(
        input_file,
        encoding="latin-1",
        sep=separator,
        usecols=list(resolved_columns.values()),
        chunksize=CHUNK_SIZE,
        low_memory=False,
    )

    aggregated_chunks: list[pd.DataFrame] = []
    for chunk_number, chunk in enumerate(reader, start=1):
        renamed_chunk = chunk.rename(columns={source: target for target, source in resolved_columns.items()})
        aggregated = aggregate_chunk(
            renamed_chunk,
            cnae_dimension,
            municipality_reference,
            active_3112_only=args.active_3112_only,
        )
        if not aggregated.empty:
            aggregated_chunks.append(aggregated)
        LOGGER.info("Chunk %s processado", chunk_number)

    if not aggregated_chunks:
        raise ValueError("Nenhum registro valido foi encontrado apos aplicar os filtros.")

    combined = pd.concat(aggregated_chunks, ignore_index=True)
    final = (
        combined.groupby(
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cbo_grande_grupo_codigo",
                "cbo_grande_grupo_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )[["remuneracao_soma", "qtd_vinculos"]]
        .sum()
    )
    final["ano_referencia"] = 2024
    final["remuneracao_media_dezembro_nom"] = final["remuneracao_soma"] / final["qtd_vinculos"]
    final = final[
        [
            "ano_referencia",
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "cbo_grande_grupo_codigo",
            "cbo_grande_grupo_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
            "qtd_vinculos",
            "remuneracao_media_dezembro_nom",
        ]
    ].sort_values(
        [
            "municipio_nome",
            "cbo_grande_grupo_codigo",
            "cnae_divisao_codigo",
        ]
    ).reset_index(drop=True)

    output_path = resolve_output_path(output_file)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        final.to_excel(writer, sheet_name="Remuneracao_Media_SC", index=False)
        auto_fit_columns(writer, "Remuneracao_Media_SC", final)

    LOGGER.info("Arquivo gerado em %s", output_path)


if __name__ == "__main__":
    main()
