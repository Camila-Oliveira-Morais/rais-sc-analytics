from __future__ import annotations

import argparse
import json
import logging
import unicodedata
from pathlib import Path
from urllib.request import urlopen

import pandas as pd
import py7zr


LOGGER = logging.getLogger("remuneracao_3112_cbo_subclasse_sc")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
CBO_DICTIONARY_FILE = PROJECT_ROOT / "data" / "dict" / "dicionario_cbo.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
DEFAULT_YEARS = ["2022", "2023", "2024", "2025"]
CHUNK_SIZE = 500_000
INPC_TABLE_ID = 1736
INPC_INDEX_VARIABLE_ID = 2289


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

    return digits.zfill(width) if width is not None else digits


def parse_numeric_series(series: pd.Series) -> pd.Series:
    direct = pd.to_numeric(series, errors="coerce")
    if direct.notna().sum() >= max(1, int(series.notna().sum() * 0.5)):
        return direct

    text = series.astype(str).str.strip()
    converted = pd.to_numeric(
        text.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    return converted if converted.notna().sum() > direct.notna().sum() else direct


def detect_separator(file_path: Path) -> str:
    with file_path.open("r", encoding="latin-1", newline="") as file_handle:
        sample = file_handle.readline()
    return ";" if sample.count(";") > sample.count(",") else ","


def resolve_first_available_column(normalized_columns: dict[str, str], candidates: list[str]) -> str | None:
    for candidate in candidates:
        matched_column = normalized_columns.get(candidate)
        if matched_column is not None:
            return matched_column
    return None


def ensure_input_file(project_root: Path, year: str, folder_suffix: str, archive_name: str) -> Path:
    extracted_dir = project_root / "data" / "raw" / f"{year}{folder_suffix}" / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if extracted_files:
        return extracted_files[0]

    archive_path = project_root / "data" / "raw" / f"{year}{folder_suffix}" / archive_name
    if not archive_path.exists():
        raise FileNotFoundError(f"Arquivo bruto nao encontrado para o ano {year}: {archive_path}")

    LOGGER.info("Extraindo %s", archive_path.name)
    with py7zr.SevenZipFile(archive_path, mode="r") as archive:
        archive.extractall(path=extracted_dir)

    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if not extracted_files:
        raise FileNotFoundError(f"Nenhum arquivo foi extraido para o ano {year}.")
    return extracted_files[0]


def load_cnae_dimension(path: Path) -> pd.DataFrame:
    dimension = pd.read_excel(path, sheet_name="CNAE-SUB")
    normalized_columns = {normalize_text(column): column for column in dimension.columns}
    dimension = dimension.rename(
        columns={
            normalized_columns["cod_subcsp"]: "cd_subclasse",
            normalized_columns["subclasse"]: "nm_subclasse",
            normalized_columns["sc competitiva"]: "nm_sc_competitiva",
            normalized_columns["gr_setor"]: "nm_gr_setor",
        }
    )

    sector_labels = {
        "agropecuaria": "Agropecuaria",
        "industria": "Industria",
        "servicos": "Servico",
    }

    dimension["cd_subclasse"] = dimension["cd_subclasse"].apply(lambda value: clean_digits(value, width=7))
    dimension["nm_subclasse"] = dimension["nm_subclasse"].fillna("Subclasse nao identificada").astype(str).str.strip()
    dimension["nm_sc_competitiva"] = (
        dimension["nm_sc_competitiva"].fillna("Nao classificado").astype(str).str.strip()
    )
    dimension["nm_gr_setor"] = (
        dimension["nm_gr_setor"]
        .fillna("Nao classificado")
        .apply(lambda value: sector_labels.get(normalize_text(value), str(value).strip()))
    )

    return (
        dimension[["cd_subclasse", "nm_subclasse", "nm_sc_competitiva", "nm_gr_setor"]]
        .dropna(subset=["cd_subclasse"])
        .drop_duplicates(subset=["cd_subclasse"])
        .reset_index(drop=True)
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
            "cd_cbo_ocupacao": parsed[0].apply(lambda value: clean_digits(value, width=6)),
            "nm_cbo_ocupacao": parsed[1].fillna("").astype(str).str.strip(),
        }
    )

    return (
        dictionary.loc[dictionary["cd_cbo_ocupacao"].notna()]
        .drop_duplicates(subset=["cd_cbo_ocupacao"])
        .reset_index(drop=True)
    )


def resolve_estab_columns(file_path: Path, separator: str) -> dict[str, str]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    required_columns = {
        "uf_codigo": ["uf - codigo", "uf"],
        "cd_subclasse": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
        "ind_rais_negativa": ["ind rais negativa - codigo", "ind rais negativa"],
    }

    resolved: dict[str, str] = {}
    for alias, candidates in required_columns.items():
        matched_column = resolve_first_available_column(normalized_columns, candidates)
        if matched_column is None:
            raise ValueError(f"Coluna obrigatoria nao encontrada: {candidates[0]}")
        resolved[alias] = matched_column
    return resolved


def resolve_vinc_columns(file_path: Path, separator: str) -> tuple[dict[str, str], bool]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    required_columns = {
        "municipio_codigo": ["municipio - codigo", "municipio"],
        "cd_subclasse": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
        "cd_cbo_ocupacao": ["cbo 2002 ocupacao - codigo", "cbo ocupacao 2002", "cbo 2002 ocupacao"],
        "vl_rem_dezembro_nom": ["vl rem dezembro nom", "vl remun dezembro nom"],
        "vl_rem_media_nom": ["vl rem media nom", "vl remun media nom"],
        "ind_vinculo_ativo_3112": ["ind vinculo ativo 31/12 - codigo", "vinculo ativo 31/12"],
    }

    resolved: dict[str, str] = {}
    for alias, candidates in required_columns.items():
        matched_column = resolve_first_available_column(normalized_columns, candidates)
        if matched_column is None:
            raise ValueError(f"Coluna obrigatoria nao encontrada: {candidates[0]}")
        resolved[alias] = matched_column

    abandonment_column = resolve_first_available_column(
        normalized_columns,
        ["ind vinculo abandonado - codigo", "ind vinculo abandonado"],
    )
    has_abandonment_filter = abandonment_column is not None
    if abandonment_column is not None:
        resolved["ind_vinculo_abandonado"] = abandonment_column

    return resolved, has_abandonment_filter


def aggregate_estab_chunk(chunk: pd.DataFrame, cnae_dimension: pd.DataFrame) -> pd.DataFrame:
    chunk["uf_codigo"] = pd.to_numeric(chunk["uf_codigo"], errors="coerce")
    chunk["ind_rais_negativa"] = pd.to_numeric(chunk["ind_rais_negativa"], errors="coerce")
    chunk["cd_subclasse"] = chunk["cd_subclasse"].apply(lambda value: clean_digits(value, width=7))

    filtered = chunk.loc[chunk["uf_codigo"].eq(42) & chunk["ind_rais_negativa"].eq(0)].copy()
    if filtered.empty:
        return filtered

    filtered["numero_estabelecimentos"] = 1
    enriched = filtered.merge(cnae_dimension, on="cd_subclasse", how="left")
    enriched["nm_subclasse"] = enriched["nm_subclasse"].fillna("Subclasse nao identificada")
    enriched["nm_sc_competitiva"] = enriched["nm_sc_competitiva"].fillna("Nao classificado")
    enriched["nm_gr_setor"] = enriched["nm_gr_setor"].fillna("Nao classificado")

    return (
        enriched.groupby(
            ["cd_subclasse", "nm_subclasse", "nm_sc_competitiva", "nm_gr_setor"],
            as_index=False,
        )
        .agg(numero_estabelecimentos=("numero_estabelecimentos", "sum"))
    )


def aggregate_vinc_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
    has_abandonment_filter: bool,
) -> pd.DataFrame:
    chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce").astype("Int64")
    chunk["vl_rem_dezembro_nom"] = parse_numeric_series(chunk["vl_rem_dezembro_nom"])
    chunk["vl_rem_media_nom"] = parse_numeric_series(chunk["vl_rem_media_nom"])
    chunk["ind_vinculo_ativo_3112"] = pd.to_numeric(chunk["ind_vinculo_ativo_3112"], errors="coerce")
    chunk["cd_subclasse"] = chunk["cd_subclasse"].apply(lambda value: clean_digits(value, width=7))
    chunk["cd_cbo_ocupacao"] = chunk["cd_cbo_ocupacao"].apply(lambda value: clean_digits(value, width=6))

    filtered = chunk.loc[
        chunk["municipio_codigo"].notna()
        & chunk["municipio_codigo"].astype(str).str.startswith("42")
        & chunk["ind_vinculo_ativo_3112"].eq(1)
        & chunk["vl_rem_dezembro_nom"].gt(0)
        & chunk["vl_rem_media_nom"].gt(0)
    ].copy()

    if has_abandonment_filter:
        filtered["ind_vinculo_abandonado"] = pd.to_numeric(filtered["ind_vinculo_abandonado"], errors="coerce")
        filtered = filtered.loc[filtered["ind_vinculo_abandonado"].eq(0)].copy()

    if filtered.empty:
        return filtered

    filtered["qt_vinculos_ativos_total"] = 1
    filtered["remuneracao_soma_nom_3112"] = filtered["vl_rem_dezembro_nom"]
    enriched = filtered.merge(cnae_dimension, on="cd_subclasse", how="left")
    enriched["nm_subclasse"] = enriched["nm_subclasse"].fillna("Subclasse nao identificada")
    enriched["nm_sc_competitiva"] = enriched["nm_sc_competitiva"].fillna("Nao classificado")
    enriched["nm_gr_setor"] = enriched["nm_gr_setor"].fillna("Nao classificado")

    return (
        enriched.groupby(
            ["cd_subclasse", "nm_subclasse", "nm_sc_competitiva", "nm_gr_setor", "cd_cbo_ocupacao"],
            as_index=False,
        )
        .agg(
            qt_vinculos_ativos_total=("qt_vinculos_ativos_total", "sum"),
            remuneracao_soma_nom_3112=("remuneracao_soma_nom_3112", "sum"),
        )
    )


def fetch_inpc_december_indexes(years: list[str]) -> pd.DataFrame:
    periods = ",".join(f"{year}12" for year in years)
    url = (
        f"https://apisidra.ibge.gov.br/values/t/{INPC_TABLE_ID}/n1/1/"
        f"v/{INPC_INDEX_VARIABLE_ID}/p/{periods}"
    )
    with urlopen(url, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))

    rows: list[dict[str, object]] = []
    for item in payload[1:]:
        rows.append(
            {
                "nr_ano_competencia": int(item["D3C"][:4]),
                "periodo_inpc": item["D3C"],
                "descricao_periodo": item["D3N"],
                "inpc_numero_indice": float(item["V"]),
            }
        )

    return pd.DataFrame(rows).sort_values("nr_ano_competencia").reset_index(drop=True)


def process_estab_year(
    year: str,
    project_root: Path,
    cnae_dimension: pd.DataFrame,
    chunk_size: int,
) -> tuple[pd.DataFrame, dict[str, int]]:
    file_path = ensure_input_file(project_root, year, "", "RAIS_ESTAB_PUB.7z")
    separator = detect_separator(file_path)
    selected_columns = resolve_estab_columns(file_path, separator)

    reader = pd.read_csv(
        file_path,
        encoding="latin-1",
        sep=separator,
        usecols=list(selected_columns.values()),
        chunksize=chunk_size,
        low_memory=False,
    )

    stats = {"linhas_lidas_estab_total": 0, "linhas_estab_validas": 0}
    grouped_frames: list[pd.DataFrame] = []

    for chunk_number, chunk in enumerate(reader, start=1):
        stats["linhas_lidas_estab_total"] += len(chunk)
        renamed = chunk.rename(columns={source: target for target, source in selected_columns.items()})
        aggregated = aggregate_estab_chunk(renamed, cnae_dimension)
        if not aggregated.empty:
            stats["linhas_estab_validas"] += int(aggregated["numero_estabelecimentos"].sum())
            grouped_frames.append(aggregated)
        LOGGER.info("Ano %s | chunk estabelecimentos %s processado", year, chunk_number)

    if not grouped_frames:
        raise ValueError(f"Nenhum estabelecimento valido encontrado para {year}.")

    final = (
        pd.concat(grouped_frames, ignore_index=True)
        .groupby(["cd_subclasse", "nm_subclasse", "nm_sc_competitiva", "nm_gr_setor"], as_index=False)[
            ["numero_estabelecimentos"]
        ]
        .sum()
    )
    final["nr_ano_competencia"] = int(year)
    return final, stats


def process_vinc_year(
    year: str,
    project_root: Path,
    cnae_dimension: pd.DataFrame,
    chunk_size: int,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    file_path = ensure_input_file(project_root, year, "_vinc_sul", "RAIS_VINC_PUB_SUL.7z")
    separator = detect_separator(file_path)
    selected_columns, has_abandonment_filter = resolve_vinc_columns(file_path, separator)

    reader = pd.read_csv(
        file_path,
        encoding="latin-1",
        sep=separator,
        usecols=list(selected_columns.values()),
        chunksize=chunk_size,
        low_memory=False,
    )

    stats = {
        "linhas_lidas_vinc_total": 0,
        "linhas_vinc_utilizadas": 0,
        "filtro_vinculo_abandonado_aplicado": "sim" if has_abandonment_filter else "nao",
    }
    grouped_frames: list[pd.DataFrame] = []

    for chunk_number, chunk in enumerate(reader, start=1):
        stats["linhas_lidas_vinc_total"] += len(chunk)
        renamed = chunk.rename(columns={source: target for target, source in selected_columns.items()})
        aggregated = aggregate_vinc_chunk(renamed, cnae_dimension, has_abandonment_filter=has_abandonment_filter)
        if not aggregated.empty:
            stats["linhas_vinc_utilizadas"] += int(aggregated["qt_vinculos_ativos_total"].sum())
            grouped_frames.append(aggregated)
        LOGGER.info("Ano %s | chunk vinculos %s processado", year, chunk_number)

    if not grouped_frames:
        raise ValueError(f"Nenhum vinculo valido encontrado para {year}.")

    final = (
        pd.concat(grouped_frames, ignore_index=True)
        .groupby(
            ["cd_subclasse", "nm_subclasse", "nm_sc_competitiva", "nm_gr_setor", "cd_cbo_ocupacao"],
            as_index=False,
        )[["qt_vinculos_ativos_total", "remuneracao_soma_nom_3112"]]
        .sum()
    )
    final["nr_ano_competencia"] = int(year)
    return final, stats


def build_detail(
    estab_frames: list[pd.DataFrame],
    vinc_frames: list[pd.DataFrame],
    cbo_dictionary: pd.DataFrame,
    inpc_table: pd.DataFrame,
) -> pd.DataFrame:
    estabelecimentos = pd.concat(estab_frames, ignore_index=True)
    vinculos = pd.concat(vinc_frames, ignore_index=True)

    merged = vinculos.merge(
        estabelecimentos,
        on=["nr_ano_competencia", "cd_subclasse", "nm_subclasse", "nm_sc_competitiva", "nm_gr_setor"],
        how="left",
    )
    merged = merged.merge(cbo_dictionary, on="cd_cbo_ocupacao", how="left")
    merged = merged.merge(inpc_table[["nr_ano_competencia", "inpc_numero_indice"]], on="nr_ano_competencia", how="left")

    base_year = int(inpc_table["nr_ano_competencia"].max())
    base_index = float(
        inpc_table.loc[inpc_table["nr_ano_competencia"].eq(base_year), "inpc_numero_indice"].iloc[0]
    )

    merged["vl_salario_medio_3112_nom"] = merged["remuneracao_soma_nom_3112"] / merged["qt_vinculos_ativos_total"]
    merged["vl_salario_medio_3112_real"] = merged["vl_salario_medio_3112_nom"] * (
        base_index / merged["inpc_numero_indice"]
    )
    merged["cd_unidade_federativa"] = 42
    merged["nm_unidade_federativa"] = "Santa Catarina"
    merged["nm_cbo_ocupacao"] = merged["nm_cbo_ocupacao"].fillna("CBO nao identificado")
    merged["numero_estabelecimentos"] = merged["numero_estabelecimentos"].fillna(0).astype("int64")

    return merged[
        [
            "cd_unidade_federativa",
            "nm_unidade_federativa",
            "nm_gr_setor",
            "cd_subclasse",
            "nm_subclasse",
            "nr_ano_competencia",
            "nm_sc_competitiva",
            "cd_cbo_ocupacao",
            "nm_cbo_ocupacao",
            "qt_vinculos_ativos_total",
            "numero_estabelecimentos",
            "vl_salario_medio_3112_nom",
            "vl_salario_medio_3112_real",
        ]
    ].sort_values(["nr_ano_competencia", "cd_subclasse", "cd_cbo_ocupacao"]).reset_index(drop=True)


def build_summary(detail: pd.DataFrame) -> pd.DataFrame:
    summary = (
        detail.groupby(
            ["nr_ano_competencia", "nm_gr_setor", "cd_subclasse", "nm_subclasse", "nm_sc_competitiva"],
            as_index=False,
        )
        .agg(
            qt_vinculos_ativos_total=("qt_vinculos_ativos_total", "sum"),
            numero_estabelecimentos=("numero_estabelecimentos", "max"),
        )
    )

    weighted_nom = (
        detail.assign(soma_nom=detail["vl_salario_medio_3112_nom"] * detail["qt_vinculos_ativos_total"])
        .groupby(["nr_ano_competencia", "nm_gr_setor", "cd_subclasse", "nm_subclasse", "nm_sc_competitiva"], as_index=False)[
            ["soma_nom"]
        ]
        .sum()
    )
    weighted_real = (
        detail.assign(soma_real=detail["vl_salario_medio_3112_real"] * detail["qt_vinculos_ativos_total"])
        .groupby(["nr_ano_competencia", "nm_gr_setor", "cd_subclasse", "nm_subclasse", "nm_sc_competitiva"], as_index=False)[
            ["soma_real"]
        ]
        .sum()
    )

    summary = summary.merge(
        weighted_nom,
        on=["nr_ano_competencia", "nm_gr_setor", "cd_subclasse", "nm_subclasse", "nm_sc_competitiva"],
        how="left",
    ).merge(
        weighted_real,
        on=["nr_ano_competencia", "nm_gr_setor", "cd_subclasse", "nm_subclasse", "nm_sc_competitiva"],
        how="left",
    )
    summary["vl_salario_medio_3112_nom"] = summary["soma_nom"] / summary["qt_vinculos_ativos_total"]
    summary["vl_salario_medio_3112_real"] = summary["soma_real"] / summary["qt_vinculos_ativos_total"]

    return summary[
        [
            "nr_ano_competencia",
            "nm_gr_setor",
            "cd_subclasse",
            "nm_subclasse",
            "nm_sc_competitiva",
            "qt_vinculos_ativos_total",
            "numero_estabelecimentos",
            "vl_salario_medio_3112_nom",
            "vl_salario_medio_3112_real",
        ]
    ].sort_values(["nr_ano_competencia", "cd_subclasse"]).reset_index(drop=True)


def build_metadata(
    years: list[str],
    estab_stats: list[dict[str, int]],
    vinc_stats: list[dict[str, int | str]],
    inpc_table: pd.DataFrame,
) -> pd.DataFrame:
    base_year = int(inpc_table["nr_ano_competencia"].max())
    metadata_rows: list[dict[str, object]] = [
        {"chave": "escopo_territorial", "valor": "Santa Catarina (UF 42)"},
        {"chave": "anos_processados", "valor": ", ".join(years)},
        {"chave": "remuneracao_origem", "valor": "Vl Rem Dezembro Nom"},
        {"chave": "filtro_vinculo_ativo_3112", "valor": "sim"},
        {"chave": "filtro_remuneracao_positiva", "valor": "sim"},
        {"chave": "filtro_vl_rem_media_nom_positivo", "valor": "sim"},
        {"chave": "deflator", "valor": "INPC geral do IBGE"},
        {"chave": "tabela_sidra", "valor": INPC_TABLE_ID},
        {"chave": "variavel_sidra", "valor": INPC_INDEX_VARIABLE_ID},
        {"chave": "base_valor_real", "valor": f"dezembro/{base_year}"},
        {
            "chave": "observacao_metodologica",
            "valor": (
                "A remuneracao real media usa a remuneracao de dezembro dos vinculos ativos em 31/12, "
                "mantem apenas registros com Vl Rem Media Nom maior que zero, "
                "desconsidera valores nulos e deflaciona pelo numero-indice do INPC de dezembro."
            ),
        },
        {
            "chave": "observacao_numero_estabelecimentos",
            "valor": "Numero de estabelecimentos validos da RAIS Estabelecimentos por subclasse/ano, repetido nas linhas de CBO.",
        },
        {
            "chave": "observacao_porte",
            "valor": "A coluna de porte do arquivo externo nao foi reproduzida porque essa classificacao nao esta identificada nos microdados locais do projeto.",
        },
    ]

    for year, stats in zip(years, estab_stats, strict=True):
        for key, value in stats.items():
            metadata_rows.append({"chave": f"estab_{key}_{year}", "valor": value})
    for year, stats in zip(years, vinc_stats, strict=True):
        for key, value in stats.items():
            metadata_rows.append({"chave": f"vinc_{key}_{year}", "valor": value})

    return pd.DataFrame(metadata_rows)


def auto_fit_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))

    workbook = writer.book
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    integer_format = workbook.add_format({"num_format": "#,##0"})
    decimal_format = workbook.add_format({"num_format": "#,##0.0000"})

    for index, column in enumerate(dataframe.columns):
        max_length = max(
            len(str(column)),
            dataframe[column].astype(str).map(len).max() if not dataframe.empty else 0,
        )
        applied_format = None
        if column in {"vl_salario_medio_3112_nom", "vl_salario_medio_3112_real"}:
            applied_format = money_format
        elif column in {"cd_unidade_federativa", "nr_ano_competencia", "qt_vinculos_ativos_total", "numero_estabelecimentos"}:
            applied_format = integer_format
        elif column == "inpc_numero_indice":
            applied_format = decimal_format
        worksheet.set_column(index, index, min(max_length + 2, 42), applied_format)


def build_default_output_path(output_dir: Path, years: list[str]) -> Path:
    suffix = years[0] if len(years) == 1 else f"{years[0]}_{years[-1]}"
    return output_dir / f"remuneracao_3112_real_sc_cbo_subclasse_{suffix}.xlsx"


def export_workbook(
    detail: pd.DataFrame,
    summary: pd.DataFrame,
    metadata: pd.DataFrame,
    inpc_table: pd.DataFrame,
    output_path: Path,
) -> Path:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        detail.to_excel(writer, sheet_name="SC_CBO_Subclasse", index=False)
        summary.to_excel(writer, sheet_name="Resumo_SC", index=False)
        metadata.to_excel(writer, sheet_name="Metadados", index=False)
        inpc_table.to_excel(writer, sheet_name="INPC_IBGE", index=False)
        auto_fit_columns(writer, "SC_CBO_Subclasse", detail)
        auto_fit_columns(writer, "Resumo_SC", summary)
        auto_fit_columns(writer, "Metadados", metadata)
        auto_fit_columns(writer, "INPC_IBGE", inpc_table)
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera um workbook com remuneracao media de dezembro nominal e real por CBO e subclasse CNAE para SC."
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
        "--cbo-dictionary-path",
        default=str(CBO_DICTIONARY_FILE),
        help="Caminho do dicionario de CBO.",
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
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()

    years = [str(year) for year in args.years]
    project_root = Path(args.project_root).resolve()
    cnae_dimension_path = Path(args.cnae_dimension_path).resolve()
    cbo_dictionary_path = Path(args.cbo_dictionary_path).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = (
        Path(args.output_path).resolve() if args.output_path else build_default_output_path(output_dir, years)
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)

    cnae_dimension = load_cnae_dimension(cnae_dimension_path)
    cbo_dictionary = load_cbo_dictionary(cbo_dictionary_path)
    inpc_table = fetch_inpc_december_indexes(years)

    estab_frames: list[pd.DataFrame] = []
    vinc_frames: list[pd.DataFrame] = []
    estab_stats: list[dict[str, int]] = []
    vinc_stats: list[dict[str, int | str]] = []

    for year in years:
        estab_year, estab_year_stats = process_estab_year(
            year=year,
            project_root=project_root,
            cnae_dimension=cnae_dimension,
            chunk_size=args.chunk_size,
        )
        vinc_year, vinc_year_stats = process_vinc_year(
            year=year,
            project_root=project_root,
            cnae_dimension=cnae_dimension,
            chunk_size=args.chunk_size,
        )
        estab_frames.append(estab_year)
        vinc_frames.append(vinc_year)
        estab_stats.append(estab_year_stats)
        vinc_stats.append(vinc_year_stats)

    detail = build_detail(estab_frames, vinc_frames, cbo_dictionary, inpc_table)
    summary = build_summary(detail)
    metadata = build_metadata(years, estab_stats, vinc_stats, inpc_table)
    export_workbook(detail, summary, metadata, inpc_table, output_path)
    LOGGER.info("Arquivo final gerado em %s", output_path)


if __name__ == "__main__":
    main()
