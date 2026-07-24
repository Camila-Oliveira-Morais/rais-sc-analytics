from __future__ import annotations

import argparse
import logging
import unicodedata
from pathlib import Path

import pandas as pd
import py7zr


LOGGER = logging.getLogger("remuneracao_media_divisao_sc")

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
MUNICIPALITY_REFERENCE_FILE = PROJECT_ROOT / "data" / "reference" / "municipios_sc_mesorregioes.csv"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
DEFAULT_YEARS = ["2022", "2023", "2024", "2025"]
CHUNK_SIZE = 500_000


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


def resolve_first_available_column(
    normalized_columns: dict[str, str],
    candidates: list[str],
) -> str | None:
    for candidate in candidates:
        matched_column = normalized_columns.get(candidate)
        if matched_column is not None:
            return matched_column
    return None


def resolve_columns(file_path: Path, separator: str) -> dict[str, str]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    required_columns = {
        "municipio_codigo": ["municipio - codigo", "municipio"],
        "cnae_subclasse_codigo": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
        "vl_rem_media_nom": ["vl rem media nom", "vl remun media nom"],
        "ind_vinculo_abandonado": ["ind vinculo abandonado - codigo", "ind vinculo abandonado"],
        "ind_vinculo_ativo_3112": ["ind vinculo ativo 31/12 - codigo", "vinculo ativo 31/12"],
    }

    resolved: dict[str, str] = {}
    for alias, candidates in required_columns.items():
        matched_column = resolve_first_available_column(normalized_columns, candidates)
        if matched_column is None:
            raise ValueError(f"Coluna obrigatoria nao encontrada: {candidates[0]}")
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


def ensure_input_file(project_root: Path, year: str) -> Path:
    extracted_dir = project_root / "data" / "raw" / f"{year}_vinc_sul" / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if extracted_files:
        return extracted_files[0]

    archive_path = project_root / "data" / "raw" / f"{year}_vinc_sul" / "RAIS_VINC_PUB_SUL.7z"
    if not archive_path.exists():
        raise FileNotFoundError(f"Arquivo bruto nao encontrado para o ano {year}: {archive_path}")

    LOGGER.info("Extraindo arquivo de vinculos de %s", year)
    with py7zr.SevenZipFile(archive_path, mode="r") as archive:
        archive.extractall(path=extracted_dir)

    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if not extracted_files:
        raise FileNotFoundError(f"Nenhum arquivo foi extraido para o ano {year}.")
    return extracted_files[0]


def aggregate_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
    industrial_only: bool,
) -> pd.DataFrame:
    chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce").astype("Int64")
    chunk["vl_rem_media_nom"] = pd.to_numeric(chunk["vl_rem_media_nom"], errors="coerce")
    chunk["ind_vinculo_abandonado"] = pd.to_numeric(chunk["ind_vinculo_abandonado"], errors="coerce")
    chunk["ind_vinculo_ativo_3112"] = pd.to_numeric(chunk["ind_vinculo_ativo_3112"], errors="coerce")
    chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )

    filtered = chunk.loc[
        chunk["municipio_codigo"].notna()
        & chunk["municipio_codigo"].astype(str).str.startswith("42")
        & chunk["ind_vinculo_abandonado"].eq(0)
        & chunk["ind_vinculo_ativo_3112"].eq(1)
        & chunk["vl_rem_media_nom"].gt(0)
    ].copy()

    if filtered.empty:
        return filtered

    filtered["municipio_codigo"] = filtered["municipio_codigo"].astype("int64")
    enriched = filtered.merge(cnae_dimension, on="cnae_subclasse_codigo", how="left")
    enriched = enriched.merge(municipality_reference, on="municipio_codigo", how="left")
    enriched = enriched.loc[enriched["cnae_divisao_codigo"].notna()].copy()

    if industrial_only:
        enriched = enriched.loc[enriched["cnae_divisao_codigo"].between("05", "43")].copy()

    if enriched.empty:
        return enriched

    enriched["municipio_nome"] = enriched["municipio_nome"].fillna("Municipio nao identificado")
    enriched["mesorregiao_nome"] = enriched["mesorregiao_nome"].fillna("Mesorregiao nao identificada")
    enriched["cnae_divisao_nome"] = enriched["cnae_divisao_nome"].fillna("Divisao nao identificada")
    enriched["qtd_vinculos"] = 1

    return (
        enriched.groupby(
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )
        .agg(
            remuneracao_soma=("vl_rem_media_nom", "sum"),
            qtd_vinculos=("qtd_vinculos", "sum"),
        )
    )


def process_year(
    year: str,
    input_file: Path,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
    chunk_size: int,
    industrial_only: bool,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    separator = detect_separator(input_file)
    try:
        resolved_columns = resolve_columns(input_file, separator)
    except ValueError as error:
        raise ValueError(f"Ano {year}: {error}") from error

    LOGGER.info("Lendo %s (%s) com separador '%s'", input_file.name, year, separator)
    reader = pd.read_csv(
        input_file,
        encoding="latin-1",
        sep=separator,
        usecols=list(resolved_columns.values()),
        chunksize=chunk_size,
        low_memory=False,
    )

    aggregated_chunks: list[pd.DataFrame] = []
    total_rows = 0
    total_used = 0

    for chunk_number, chunk in enumerate(reader, start=1):
        total_rows += len(chunk)
        renamed_chunk = chunk.rename(columns={source: target for target, source in resolved_columns.items()})
        aggregated = aggregate_chunk(renamed_chunk, cnae_dimension, municipality_reference, industrial_only)
        if not aggregated.empty:
            total_used += int(aggregated["qtd_vinculos"].sum())
            aggregated_chunks.append(aggregated)
        LOGGER.info("Ano %s | chunk %s processado", year, chunk_number)

    if not aggregated_chunks:
        raise ValueError(f"Nenhum registro valido foi encontrado para o ano {year}.")

    combined = pd.concat(aggregated_chunks, ignore_index=True)
    final = (
        combined.groupby(
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )[["remuneracao_soma", "qtd_vinculos"]]
        .sum()
    )
    final["ano_referencia"] = int(year)
    final["remuneracao_media_nom_ano"] = final["remuneracao_soma"] / final["qtd_vinculos"]

    ordered = final[
        [
            "ano_referencia",
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
            "qtd_vinculos",
            "remuneracao_media_nom_ano",
            "remuneracao_soma",
        ]
    ].sort_values(["municipio_nome", "cnae_divisao_codigo"]).reset_index(drop=True)

    stats = {
        "ano_referencia": int(year),
        "arquivo_entrada": str(input_file),
        "linhas_lidas": total_rows,
        "vinculos_utilizados": total_used,
    }
    return ordered, stats


def build_summary(dataframe: pd.DataFrame) -> pd.DataFrame:
    summary = (
        dataframe.groupby(
            [
                "ano_referencia",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )[["remuneracao_soma", "qtd_vinculos"]]
        .sum()
    )
    summary["remuneracao_media_nom_ano"] = summary["remuneracao_soma"] / summary["qtd_vinculos"]
    return summary[
        [
            "ano_referencia",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
            "qtd_vinculos",
            "remuneracao_media_nom_ano",
        ]
    ].sort_values(["ano_referencia", "cnae_divisao_codigo"]).reset_index(drop=True)


def build_metadata(stats_rows: list[dict[str, int | str]], industrial_only: bool) -> pd.DataFrame:
    metadata_rows: list[dict[str, object]] = [
        {"chave": "filtro_uf", "valor": "Santa Catarina (municipios com codigo iniciado por 42)"},
        {"chave": "filtro_vinculo_ativo_3112", "valor": "sim"},
        {"chave": "filtro_ind_vinculo_abandonado", "valor": "nao"},
        {"chave": "filtro_vl_rem_media_nom_positivo", "valor": "sim"},
        {"chave": "filtro_industrial_05_43", "valor": "sim" if industrial_only else "nao"},
        {"chave": "metrica_remuneracao", "valor": "media de Vl Rem Media Nom ponderada por qtd_vinculos"},
    ]

    for row in stats_rows:
        year = row["ano_referencia"]
        for key, value in row.items():
            if key == "ano_referencia":
                continue
            metadata_rows.append({"chave": f"{key}_{year}", "valor": value})

    return pd.DataFrame(metadata_rows)


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
        applied_format = None
        if column in {"remuneracao_media_nom_ano"}:
            applied_format = money_format
        elif column in {"qtd_vinculos", "ano_referencia", "municipio_codigo"}:
            applied_format = integer_format
        worksheet.set_column(index, index, min(max_length + 2, 40), applied_format)


def build_default_output_path(output_dir: Path, years: list[str], industrial_only: bool) -> Path:
    suffix = years[0] if len(years) == 1 else f"{years[0]}_{years[-1]}"
    basename = "remuneracao_media_ano_vinculos_sc_cnae_divisao"
    if industrial_only:
        basename = f"{basename}_industrial"
    return output_dir / f"{basename}_{suffix}.xlsx"


def export_workbook(
    detail: pd.DataFrame,
    summary: pd.DataFrame,
    metadata: pd.DataFrame,
    output_path: Path,
) -> Path:
    export_detail = detail.drop(columns=["remuneracao_soma"])
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        export_detail.to_excel(writer, sheet_name="Municipio_Divisao", index=False)
        summary.to_excel(writer, sheet_name="Resumo_Divisao_SC", index=False)
        metadata.to_excel(writer, sheet_name="Metadados", index=False)
        auto_fit_columns(writer, "Municipio_Divisao", export_detail)
        auto_fit_columns(writer, "Resumo_Divisao_SC", summary)
        auto_fit_columns(writer, "Metadados", metadata)
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera remuneracao media anual nominal da RAIS vinculo para SC por CNAE divisao."
    )
    parser.add_argument(
        "--years",
        nargs="+",
        default=DEFAULT_YEARS,
        help="Lista de anos para processar. Exemplo: --years 2022 2023 2024 2025",
    )
    parser.add_argument(
        "--project-root",
        default=str(PROJECT_ROOT),
        help="Diretorio raiz do projeto.",
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
        "--output-dir",
        default=str(OUTPUT_DIR),
        help="Diretorio base das saidas.",
    )
    parser.add_argument(
        "--output-path",
        default=None,
        help="Caminho final do workbook. Se omitido, usa o nome padrao no diretorio de saida.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=CHUNK_SIZE,
        help="Quantidade de linhas por chunk durante a leitura.",
    )
    parser.add_argument(
        "--industrial-only",
        action="store_true",
        help="Restringe a divisao CNAE ao intervalo 05 a 43.",
    )
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()

    years = [str(year) for year in args.years]
    project_root = Path(args.project_root).resolve()
    cnae_dimension_path = Path(args.cnae_dimension_path).resolve()
    municipality_reference_path = Path(args.municipality_reference_path).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = (
        Path(args.output_path).resolve()
        if args.output_path
        else build_default_output_path(output_dir, years, args.industrial_only)
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)

    cnae_dimension = load_cnae_dimension(cnae_dimension_path)
    municipality_reference = load_municipality_reference(municipality_reference_path)

    yearly_frames: list[pd.DataFrame] = []
    stats_rows: list[dict[str, int | str]] = []

    for year in years:
        input_file = ensure_input_file(project_root, year)
        final_year, stats = process_year(
            year=year,
            input_file=input_file,
            cnae_dimension=cnae_dimension,
            municipality_reference=municipality_reference,
            chunk_size=args.chunk_size,
            industrial_only=args.industrial_only,
        )
        yearly_frames.append(final_year)
        stats_rows.append(stats)

    detail = pd.concat(yearly_frames, ignore_index=True).sort_values(
        ["ano_referencia", "municipio_nome", "cnae_divisao_codigo"]
    ).reset_index(drop=True)
    summary = build_summary(detail)
    metadata = build_metadata(stats_rows, args.industrial_only)

    export_workbook(detail, summary, metadata, output_path)
    LOGGER.info("Arquivo final gerado em %s", output_path)


if __name__ == "__main__":
    main()
