from __future__ import annotations

import argparse
import csv
import gzip
import json
import logging
import sqlite3
from dataclasses import dataclass
from ftplib import FTP
from pathlib import Path
from typing import Iterable
from urllib.request import Request, urlopen

import pandas as pd
import py7zr
from tqdm import tqdm


LOGGER = logging.getLogger("rais_sc_pipeline")

FTP_HOST = "ftp.mtps.gov.br"
FTP_ROOT = "/pdet/microdados/RAIS"
IBGE_MESO_URL = "http://servicodados.ibge.gov.br/api/v1/localidades/estados/42/mesorregioes"
IBGE_MUNICIPIOS_URL = "http://servicodados.ibge.gov.br/api/v1/localidades/mesorregioes/{mesorregiao_id}/municipios"

RAIS_COLUMN_CANDIDATES = {
    "cnae_classe_codigo": ["CNAE 2.0 Classe - Código", "CNAE 2.0 Classe"],
    "cnae_subclasse_codigo": ["CNAE 2.0 Subclasse - Código", "CNAE 2.0 Subclasse"],
    "ind_atividade_ano": ["Ind Atividade Ano - Código", "Ind Atividade Ano"],
    "ind_rais_negativa": ["Ind RAIS Negativa - Código", "Ind Rais Negativa"],
    "municipio_codigo": ["Município - Código", "Município"],
    "qtd_vinculos_ativos": ["Qtd Vínculos Ativos"],
    "qtd_vinculos_clt": ["Qtd Vínculos CLT"],
    "qtd_vinculos_estatutarios": ["Qtd Vínculos Estatutários"],
    "uf_codigo": ["UF - Código", "UF"],
}

REFERENCE_COLUMNS = [
    "municipio_codigo",
    "municipio_nome",
    "mesorregiao_id",
    "mesorregiao_nome",
]


@dataclass
class PipelineConfig:
    years: list[str]
    project_root: Path
    cnae_dimension_path: Path
    ftp_host: str = FTP_HOST
    ftp_root: str = FTP_ROOT
    chunk_size: int = 200_000
    force_download: bool = False
    force_extract: bool = False

    @property
    def raw_dir(self) -> Path:
        if len(self.years) != 1:
            raise ValueError("raw_dir is only available when processing a single year.")
        return self.project_root / "data" / "raw" / self.years[0]

    @property
    def raw_archive_path(self) -> Path:
        return self.raw_dir / "RAIS_ESTAB_PUB.7z"

    @property
    def extracted_dir(self) -> Path:
        return self.raw_dir / "extracted"

    @property
    def output_dir(self) -> Path:
        return self.project_root / "data" / "output"

    @property
    def reference_dir(self) -> Path:
        return self.project_root / "data" / "reference"

    @property
    def municipality_reference_path(self) -> Path:
        return self.reference_dir / "municipios_sc_mesorregioes.csv"

    @property
    def output_workbook_path(self) -> Path:
        if len(self.years) == 1:
            suffix = self.years[0]
        else:
            suffix = f"{self.years[0]}_{self.years[-1]}"
        return self.output_dir / f"rais_estabelecimentos_sc_{suffix}.xlsx"

    @property
    def output_database_path(self) -> Path:
        if len(self.years) == 1:
            suffix = self.years[0]
        else:
            suffix = f"{self.years[0]}_{self.years[-1]}"
        return self.output_dir / f"rais_estabelecimentos_sc_{suffix}.sqlite"

    def raw_dir_for_year(self, year: str) -> Path:
        return self.project_root / "data" / "raw" / year

    def raw_archive_path_for_year(self, year: str) -> Path:
        return self.raw_dir_for_year(year) / "RAIS_ESTAB_PUB.7z"

    def extracted_dir_for_year(self, year: str) -> Path:
        return self.raw_dir_for_year(year) / "extracted"

    def remote_directory_for_year(self, year: str) -> str:
        return f"{self.ftp_root}/{year}"


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def ensure_directories(config: PipelineConfig) -> None:
    paths = [config.output_dir, config.reference_dir]
    for year in config.years:
        paths.extend(
            [
                config.raw_dir_for_year(year),
                config.extracted_dir_for_year(year),
            ]
        )

    for path in paths:
        path.mkdir(parents=True, exist_ok=True)


def download_rais_archive(config: PipelineConfig, year: str) -> Path:
    destination = config.raw_archive_path_for_year(year)
    if destination.exists() and not config.force_download:
        LOGGER.info("Arquivo bruto ja existe em %s", destination)
        return destination

    LOGGER.info("Baixando RAIS de estabelecimentos %s do FTP oficial", year)
    with FTP(config.ftp_host, timeout=120) as ftp:
        ftp.encoding = "latin-1"
        ftp.login()
        ftp.cwd(config.remote_directory_for_year(year))
        total_size = ftp.size(destination.name)
        with destination.open("wb") as output_file:
            with tqdm(
                total=total_size,
                unit="B",
                unit_scale=True,
                desc=destination.name,
            ) as progress:

                def write_chunk(chunk: bytes) -> None:
                    output_file.write(chunk)
                    progress.update(len(chunk))

                ftp.retrbinary(
                    f"RETR {destination.name}",
                    write_chunk,
                    blocksize=1024 * 256,
                )

    LOGGER.info("Download concluido: %s", destination)
    return destination


def extract_7z_archive(config: PipelineConfig, year: str) -> Path:
    extracted_dir = config.extracted_dir_for_year(year)
    extracted_files = list(extracted_dir.glob("*"))
    if extracted_files and not config.force_extract:
        LOGGER.info("Arquivo extraido ja encontrado em %s", extracted_dir)
        return extracted_files[0]

    LOGGER.info("Extraindo %s", config.raw_archive_path_for_year(year).name)
    for file_path in extracted_files:
        if file_path.is_file():
            file_path.unlink()

    with py7zr.SevenZipFile(config.raw_archive_path_for_year(year), mode="r") as archive:
        archive.extractall(path=extracted_dir)

    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if not extracted_files:
        raise FileNotFoundError("Nenhum arquivo foi extraido do arquivo .7z da RAIS.")

    LOGGER.info("Extracao concluida: %s", extracted_files[0])
    return extracted_files[0]


def read_ibge_json(url: str) -> list[dict]:
    request = Request(url, headers={"Accept-Encoding": "gzip"})
    with urlopen(request, timeout=60) as response:
        payload = response.read()

    if payload[:2] == b"\x1f\x8b":
        payload = gzip.decompress(payload)

    return json.loads(payload.decode("utf-8"))


def build_or_load_municipality_reference(config: PipelineConfig) -> pd.DataFrame:
    cache_path = config.municipality_reference_path
    if cache_path.exists():
        LOGGER.info("Usando cache territorial em %s", cache_path)
        reference = pd.read_csv(cache_path)
        if reference["municipio_codigo"].max() > 999999:
            reference["municipio_codigo"] = (
                pd.to_numeric(reference["municipio_codigo"], errors="coerce")
                .fillna(0)
                .astype("int64")
                .floordiv(10)
            )
            reference = reference.drop_duplicates(subset=["municipio_codigo"]).reset_index(drop=True)
            reference.to_csv(cache_path, index=False, quoting=csv.QUOTE_MINIMAL)
        return reference[REFERENCE_COLUMNS]

    LOGGER.info("Consultando API do IBGE para montar a dimensao de mesorregioes de SC")
    mesorregioes = read_ibge_json(IBGE_MESO_URL)
    records: list[dict[str, object]] = []

    for mesorregiao in mesorregioes:
        municipios = read_ibge_json(
            IBGE_MUNICIPIOS_URL.format(mesorregiao_id=mesorregiao["id"])
        )
        for municipio in municipios:
            records.append(
                {
                    "municipio_codigo": int(municipio["id"]) // 10,
                    "municipio_nome": municipio["nome"],
                    "mesorregiao_id": int(mesorregiao["id"]),
                    "mesorregiao_nome": mesorregiao["nome"],
                }
            )

    reference = pd.DataFrame(records).sort_values("municipio_nome").reset_index(drop=True)
    reference.to_csv(cache_path, index=False, quoting=csv.QUOTE_MINIMAL)
    LOGGER.info("Cache territorial salvo em %s", cache_path)
    return reference[REFERENCE_COLUMNS]


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


def detect_csv_separator(file_path: Path) -> str:
    with file_path.open("r", encoding="latin-1", newline="") as file_handle:
        sample = file_handle.readline()

    if sample.count(";") > sample.count(","):
        return ";"
    return ","


def detect_rais_columns(file_path: Path, separator: str) -> dict[str, str]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    available_columns = set(header.columns.tolist())
    resolved_columns: dict[str, str] = {}

    for normalized_name, candidates in RAIS_COLUMN_CANDIDATES.items():
        matched_column = next((column for column in candidates if column in available_columns), None)
        if matched_column is None:
            raise ValueError(
                f"Coluna obrigatoria nao encontrada para '{normalized_name}'. "
                f"Candidatas testadas: {candidates}"
            )
        resolved_columns[normalized_name] = matched_column

    return resolved_columns


def load_cnae_dimension(path: Path) -> pd.DataFrame:
    LOGGER.info("Lendo dimensao de CNAE em %s", path)
    dimension = pd.read_excel(path, sheet_name="CNAE-SUB")
    dimension = dimension.rename(
        columns={
            "COD_DIV": "cnae_divisao_codigo",
            "Divisão": "cnae_divisao_nome",
            "COD_SUBCsp": "cnae_subclasse_codigo",
            "Subclasse": "cnae_subclasse_nome",
            "SC Competitiva": "sc_competitiva",
        }
    )
    dimension["cnae_subclasse_codigo"] = dimension["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )
    dimension["cnae_divisao_codigo"] = dimension["cnae_divisao_codigo"].apply(
        lambda value: clean_digits(value, width=2)
    )
    dimension["sc_competitiva"] = (
        dimension["sc_competitiva"].fillna("Nao classificado").astype(str).str.strip()
    )

    dimension = (
        dimension[
            [
                "cnae_subclasse_codigo",
                "cnae_subclasse_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
                "sc_competitiva",
            ]
        ]
        .dropna(subset=["cnae_subclasse_codigo", "cnae_divisao_codigo"])
        .drop_duplicates(subset=["cnae_subclasse_codigo"])
        .reset_index(drop=True)
    )
    return dimension


def aggregate_chunk(
    chunk: pd.DataFrame,
    selected_columns: dict[str, str],
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
) -> pd.DataFrame:
    prepared = chunk.rename(columns={source: target for target, source in selected_columns.items()})

    numeric_columns = [
        "municipio_codigo",
        "uf_codigo",
        "ind_atividade_ano",
        "ind_rais_negativa",
        "qtd_vinculos_ativos",
        "qtd_vinculos_clt",
        "qtd_vinculos_estatutarios",
    ]
    for column in numeric_columns:
        prepared[column] = pd.to_numeric(prepared[column], errors="coerce").fillna(0)

    filtered = prepared.loc[prepared["uf_codigo"].eq(42)].copy()
    if filtered.empty:
        return filtered

    filtered["municipio_codigo"] = filtered["municipio_codigo"].astype("int64")
    filtered["cnae_subclasse_codigo"] = filtered["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )
    filtered["estabelecimentos"] = 1
    filtered["estabelecimentos_ativos_ano"] = filtered["ind_atividade_ano"].eq(1).astype("int64")
    filtered["estabelecimentos_rais_negativa"] = filtered["ind_rais_negativa"].ne(0).astype("int64")

    enriched = filtered.merge(cnae_dimension, on="cnae_subclasse_codigo", how="left")
    enriched = enriched.merge(municipality_reference, on="municipio_codigo", how="left")

    enriched["cnae_divisao_codigo"] = enriched["cnae_divisao_codigo"].fillna("NA")
    enriched["cnae_divisao_nome"] = enriched["cnae_divisao_nome"].fillna("Divisao nao identificada")
    enriched["sc_competitiva"] = enriched["sc_competitiva"].fillna("Nao classificado")
    enriched["municipio_nome"] = enriched["municipio_nome"].fillna("Municipio nao identificado")
    enriched["mesorregiao_nome"] = enriched["mesorregiao_nome"].fillna("Mesorregiao nao identificada")

    return enriched


AGG_SPECS = {
    "estabelecimentos": "sum",
    "estabelecimentos_ativos_ano": "sum",
    "estabelecimentos_rais_negativa": "sum",
    "qtd_vinculos_ativos": "sum",
    "qtd_vinculos_clt": "sum",
    "qtd_vinculos_estatutarios": "sum",
}

BASE_METRIC_COLUMNS = list(AGG_SPECS.keys())


def combine_grouped_frames(frames: Iterable[pd.DataFrame], group_keys: list[str]) -> pd.DataFrame:
    valid_frames = [frame for frame in frames if not frame.empty]
    if not valid_frames:
        return pd.DataFrame(columns=group_keys + BASE_METRIC_COLUMNS)

    combined = pd.concat(valid_frames, ignore_index=True)
    metrics = [column for column in combined.columns if column not in group_keys]
    return (
        combined.groupby(group_keys, as_index=False)[metrics]
        .sum()
        .sort_values(group_keys)
        .reset_index(drop=True)
    )


def aggregate_outputs(
    frames: dict[str, list[pd.DataFrame]],
    include_negative_code_summary: bool = False,
) -> dict[str, pd.DataFrame]:
    grouped_data = {
        "municipio_divisao": combine_grouped_frames(
            frames["municipio_divisao"],
            [
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
        ),
        "mesorregiao_divisao": combine_grouped_frames(
            frames["mesorregiao_divisao"],
            ["mesorregiao_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
        ),
        "sc_competitiva_municipio": combine_grouped_frames(
            frames["sc_competitiva_municipio"],
            ["municipio_codigo", "municipio_nome", "mesorregiao_nome", "sc_competitiva"],
        ),
        "sc_competitiva_mesorregiao": combine_grouped_frames(
            frames["sc_competitiva_mesorregiao"],
            ["mesorregiao_nome", "sc_competitiva"],
        ),
        "resumo_divisao": combine_grouped_frames(
            frames["resumo_divisao"],
            ["cnae_divisao_codigo", "cnae_divisao_nome"],
        ).sort_values("qtd_vinculos_ativos", ascending=False),
        "resumo_municipio": combine_grouped_frames(
            frames["resumo_municipio"],
            ["municipio_codigo", "municipio_nome", "mesorregiao_nome"],
        ).sort_values("qtd_vinculos_ativos", ascending=False),
        "resumo_sc_competitiva": combine_grouped_frames(
            frames["resumo_sc_competitiva"],
            ["sc_competitiva"],
        ).sort_values("qtd_vinculos_ativos", ascending=False),
    }
    grouped_data["resumo_mesorregiao"] = (
        grouped_data["mesorregiao_divisao"]
        .groupby("mesorregiao_nome", as_index=False)[BASE_METRIC_COLUMNS]
        .sum()
        .sort_values("qtd_vinculos_ativos", ascending=False)
        .reset_index(drop=True)
    )

    if include_negative_code_summary:
        grouped_data["resumo_rais_negativa_codigo"] = combine_grouped_frames(
            frames["resumo_rais_negativa_codigo"],
            ["ind_rais_negativa"],
        ).sort_values(["ind_rais_negativa"], ascending=True)

    return grouped_data


def process_rais_file(
    config: PipelineConfig,
    year: str,
    extracted_file_path: Path,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame], dict[str, int]]:
    LOGGER.info("Processando %s em chunks de %s linhas", extracted_file_path.name, config.chunk_size)
    separator = detect_csv_separator(extracted_file_path)
    selected_columns = detect_rais_columns(extracted_file_path, separator)
    LOGGER.info("Layout detectado para %s: separador '%s'", year, separator)
    grouped_frames = {
        "municipio_divisao": [],
        "mesorregiao_divisao": [],
        "sc_competitiva_municipio": [],
        "sc_competitiva_mesorregiao": [],
        "resumo_divisao": [],
        "resumo_municipio": [],
        "resumo_sc_competitiva": [],
    }
    negative_frames = {
        "municipio_divisao": [],
        "mesorregiao_divisao": [],
        "sc_competitiva_municipio": [],
        "sc_competitiva_mesorregiao": [],
        "resumo_divisao": [],
        "resumo_municipio": [],
        "resumo_sc_competitiva": [],
        "resumo_rais_negativa_codigo": [],
    }

    stats = {
        "ano_referencia": year,
        "linhas_lidas_total": 0,
        "linhas_filtradas_sc": 0,
        "linhas_rais_negativa": 0,
        "linhas_validas_sem_rais_negativa": 0,
    }

    reader = pd.read_csv(
        extracted_file_path,
        encoding="latin-1",
        sep=separator,
        usecols=list(selected_columns.values()),
        chunksize=config.chunk_size,
        low_memory=False,
    )

    for chunk_number, chunk in enumerate(reader, start=1):
        stats["linhas_lidas_total"] += len(chunk)
        enriched = aggregate_chunk(chunk, selected_columns, cnae_dimension, municipality_reference)
        if enriched.empty:
            LOGGER.info("Chunk %s sem registros de SC", chunk_number)
            continue

        stats["linhas_filtradas_sc"] += len(enriched)
        negative = enriched.loc[enriched["ind_rais_negativa"].ne(0)].copy()
        valid = enriched.loc[enriched["ind_rais_negativa"].eq(0)].copy()

        stats["linhas_rais_negativa"] += len(negative)
        stats["linhas_validas_sem_rais_negativa"] += len(valid)
        LOGGER.info(
            "Chunk %s com %s registros de SC (%s validos e %s RAIS negativa)",
            chunk_number,
            len(enriched),
            len(valid),
            len(negative),
        )

        if not valid.empty:
            grouped_frames["municipio_divisao"].append(
                valid.groupby(
                    ["municipio_codigo", "municipio_nome", "mesorregiao_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
                    as_index=False,
                ).agg(AGG_SPECS)
            )
            grouped_frames["mesorregiao_divisao"].append(
                valid.groupby(
                    ["mesorregiao_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
                    as_index=False,
                ).agg(AGG_SPECS)
            )
            grouped_frames["sc_competitiva_municipio"].append(
                valid.groupby(
                    ["municipio_codigo", "municipio_nome", "mesorregiao_nome", "sc_competitiva"],
                    as_index=False,
                ).agg(AGG_SPECS)
            )
            grouped_frames["sc_competitiva_mesorregiao"].append(
                valid.groupby(["mesorregiao_nome", "sc_competitiva"], as_index=False).agg(AGG_SPECS)
            )
            grouped_frames["resumo_divisao"].append(
                valid.groupby(["cnae_divisao_codigo", "cnae_divisao_nome"], as_index=False).agg(AGG_SPECS)
            )
            grouped_frames["resumo_municipio"].append(
                valid.groupby(["municipio_codigo", "municipio_nome", "mesorregiao_nome"], as_index=False).agg(AGG_SPECS)
            )
            grouped_frames["resumo_sc_competitiva"].append(
                valid.groupby(["sc_competitiva"], as_index=False).agg(AGG_SPECS)
            )

        if not negative.empty:
            negative_frames["municipio_divisao"].append(
                negative.groupby(
                    ["municipio_codigo", "municipio_nome", "mesorregiao_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
                    as_index=False,
                ).agg(AGG_SPECS)
            )
            negative_frames["mesorregiao_divisao"].append(
                negative.groupby(
                    ["mesorregiao_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
                    as_index=False,
                ).agg(AGG_SPECS)
            )
            negative_frames["sc_competitiva_municipio"].append(
                negative.groupby(
                    ["municipio_codigo", "municipio_nome", "mesorregiao_nome", "sc_competitiva"],
                    as_index=False,
                ).agg(AGG_SPECS)
            )
            negative_frames["sc_competitiva_mesorregiao"].append(
                negative.groupby(["mesorregiao_nome", "sc_competitiva"], as_index=False).agg(AGG_SPECS)
            )
            negative_frames["resumo_divisao"].append(
                negative.groupby(["cnae_divisao_codigo", "cnae_divisao_nome"], as_index=False).agg(AGG_SPECS)
            )
            negative_frames["resumo_municipio"].append(
                negative.groupby(["municipio_codigo", "municipio_nome", "mesorregiao_nome"], as_index=False).agg(AGG_SPECS)
            )
            negative_frames["resumo_sc_competitiva"].append(
                negative.groupby(["sc_competitiva"], as_index=False).agg(AGG_SPECS)
            )
            negative_frames["resumo_rais_negativa_codigo"].append(
                negative.groupby(["ind_rais_negativa"], as_index=False).agg(AGG_SPECS)
            )

    grouped_data = aggregate_outputs(grouped_frames)
    negative_grouped_data = aggregate_outputs(negative_frames, include_negative_code_summary=True)
    return grouped_data, negative_grouped_data, stats


def write_table(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))

    for index, column in enumerate(dataframe.columns):
        max_length = max(
            len(str(column)),
            dataframe[column].astype(str).map(len).max() if not dataframe.empty else 0,
        )
        worksheet.set_column(index, index, min(max_length + 2, 28))


def write_summary_sheet(
    writer: pd.ExcelWriter,
    grouped_data: dict[str, pd.DataFrame],
    negative_grouped_data: dict[str, pd.DataFrame],
    stats: pd.DataFrame,
    config: PipelineConfig,
) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Resumo")
    writer.sheets["Resumo"] = worksheet

    title_format = workbook.add_format({"bold": True, "font_size": 14})
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9E2F3", "border": 1})
    number_format = workbook.add_format({"num_format": "#,##0", "border": 1})
    text_format = workbook.add_format({"border": 1})

    years_label = ", ".join(config.years)
    worksheet.write("A1", f"RAIS Estabelecimentos Santa Catarina {years_label}", title_format)
    summary_rows = [
        ("Anos de referencia", years_label),
        ("Workbook gerado em", str(config.output_workbook_path)),
        ("Observacao", "Abas principais excluem RAIS negativa; ver aba Resumo_RAIS_Negativa."),
    ]

    worksheet.write_row("A3", ["Indicador", "Valor"], header_format)
    for row_number, row in enumerate(summary_rows, start=3):
        worksheet.write(row_number, 0, row[0], text_format)
        worksheet.write(row_number, 1, row[1], text_format)

    stats_table = stats.copy()
    stats_table["arquivo_origem"] = stats_table["ano_referencia"].apply(
        lambda year: f"ftp://{config.ftp_host}{config.remote_directory_for_year(str(year))}/RAIS_ESTAB_PUB.7z"
    )
    stats_table = stats_table[
        [
            "ano_referencia",
            "linhas_lidas_total",
            "linhas_filtradas_sc",
            "linhas_validas_sem_rais_negativa",
            "linhas_rais_negativa",
            "arquivo_origem",
        ]
    ]
    worksheet.write(7, 0, "Processamento por ano", title_format)
    worksheet.write_row(8, 0, stats_table.columns.tolist(), header_format)
    for row_offset, (_, row) in enumerate(stats_table.iterrows(), start=0):
        for col_offset, value in enumerate(row.tolist()):
            cell_format = number_format if isinstance(value, int) else text_format
            worksheet.write(9 + row_offset, col_offset, value, cell_format)

    latest_year = max(config.years)
    top_divisoes = (
        grouped_data["resumo_divisao"]
        .loc[grouped_data["resumo_divisao"]["ano_referencia"].astype(str).eq(latest_year)]
        .head(10)
        .copy()
    )
    top_municipios = (
        grouped_data["resumo_municipio"]
        .loc[grouped_data["resumo_municipio"]["ano_referencia"].astype(str).eq(latest_year)]
        .head(10)
        .copy()
    )
    top_mesorregioes = (
        grouped_data["resumo_mesorregiao"]
        .loc[grouped_data["resumo_mesorregiao"]["ano_referencia"].astype(str).eq(latest_year)]
        .head(10)
        .copy()
    )
    top_sc_comp = (
        grouped_data["resumo_sc_competitiva"]
        .loc[grouped_data["resumo_sc_competitiva"]["ano_referencia"].astype(str).eq(latest_year)]
        .head(10)
        .copy()
    )

    tables = [
        ("Top divisões por vínculos ativos", top_divisoes[["cnae_divisao_nome", "qtd_vinculos_ativos"]], 14, "A"),
        ("Top municípios por vínculos ativos", top_municipios[["municipio_nome", "qtd_vinculos_ativos"]], 14, "E"),
        ("Mesorregiões por vínculos ativos", top_mesorregioes[["mesorregiao_nome", "qtd_vinculos_ativos"]], 14, "I"),
        ("SC Competitiva por vínculos ativos", top_sc_comp[["sc_competitiva", "qtd_vinculos_ativos"]], 14, "M"),
    ]

    for title, table, start_row, start_col_letter in tables:
        start_col = ord(start_col_letter) - ord("A")
        worksheet.write(start_row, start_col, title, title_format)
        worksheet.write_row(start_row + 1, start_col, table.columns.tolist(), header_format)
        for offset, (_, row) in enumerate(table.iterrows(), start=0):
            worksheet.write(start_row + 2 + offset, start_col, row.iloc[0], text_format)
            worksheet.write(start_row + 2 + offset, start_col + 1, int(row.iloc[1]), number_format)
        worksheet.set_column(start_col, start_col, 28)
        worksheet.set_column(start_col + 1, start_col + 1, 16)

    chart_specs = [
        ("Divisões", 14, 0, 4),
        ("Municípios", 14, 4, 8),
        ("Mesorregiões", 14, 8, 12),
        ("SC Competitiva", 14, 12, 16),
    ]

    for title, start_row, start_col, chart_col in chart_specs:
        chart = workbook.add_chart({"type": "column"})
        categories_col = start_col
        values_col = start_col + 1
        chart.add_series(
            {
                "name": f"{title} - vínculos ativos",
                "categories": ["Resumo", start_row + 2, categories_col, start_row + 11, categories_col],
                "values": ["Resumo", start_row + 2, values_col, start_row + 11, values_col],
            }
        )
        chart.set_title({"name": title})
        chart.set_y_axis({"name": "Vínculos ativos"})
        chart.set_legend({"none": True})
        worksheet.insert_chart(start_row + 14, chart_col, chart, {"x_scale": 1.2, "y_scale": 1.2})

    negative_summary = negative_grouped_data["resumo_rais_negativa_codigo"].copy()
    negative_summary = negative_summary.rename(columns={"ind_rais_negativa": "codigo_rais_negativa"})
    worksheet.write(32, 0, f"Resumo RAIS negativa ({latest_year})", title_format)
    latest_negative = negative_summary.loc[
        negative_summary["ano_referencia"].astype(str).eq(latest_year)
    ]
    worksheet.write_row(33, 0, latest_negative.columns.tolist(), header_format)
    for row_offset, (_, row) in enumerate(latest_negative.iterrows(), start=0):
        for col_offset, value in enumerate(row.tolist()):
            cell_format = number_format if isinstance(value, int) else text_format
            worksheet.write(34 + row_offset, col_offset, value, cell_format)


def export_to_excel(
    grouped_data: dict[str, pd.DataFrame],
    negative_grouped_data: dict[str, pd.DataFrame],
    stats: pd.DataFrame,
    config: PipelineConfig,
) -> Path:
    LOGGER.info("Gerando workbook analitico em %s", config.output_workbook_path)
    with pd.ExcelWriter(config.output_workbook_path, engine="xlsxwriter") as writer:
        write_summary_sheet(writer, grouped_data, negative_grouped_data, stats, config)
        write_table(writer, "Municipio_Divisao", grouped_data["municipio_divisao"])
        write_table(writer, "Mesorregiao_Divisao", grouped_data["mesorregiao_divisao"])
        write_table(writer, "Municipio_SC_Compet", grouped_data["sc_competitiva_municipio"])
        write_table(writer, "Mesorregiao_SC_Compet", grouped_data["sc_competitiva_mesorregiao"])
        write_table(writer, "Resumo_Divisao", grouped_data["resumo_divisao"])
        write_table(writer, "Resumo_Municipio", grouped_data["resumo_municipio"])
        write_table(writer, "Resumo_Mesorregiao", grouped_data["resumo_mesorregiao"])
        write_table(writer, "Resumo_SC_Compet", grouped_data["resumo_sc_competitiva"])
        write_table(writer, "Resumo_RAIS_Negativa", negative_grouped_data["resumo_rais_negativa_codigo"])
        write_table(writer, "RAIS_Negativa_Divisao", negative_grouped_data["resumo_divisao"])
        write_table(writer, "RAIS_Negativa_Municipio", negative_grouped_data["resumo_municipio"])
        write_table(writer, "RAIS_Negativa_Mesorregiao", negative_grouped_data["resumo_mesorregiao"])
        write_table(writer, "RAIS_Negativa_SC_Compet", negative_grouped_data["resumo_sc_competitiva"])
        metadata_rows: list[dict[str, object]] = [
            {"chave": "anos_referencia", "valor": ", ".join(config.years)},
            {"chave": "dimensao_cnae", "valor": str(config.cnae_dimension_path)},
            {"chave": "referencia_mesorregioes", "valor": str(config.municipality_reference_path)},
        ]
        for year in config.years:
            metadata_rows.extend(
                [
                    {"chave": f"arquivo_origem_{year}", "valor": str(config.raw_archive_path_for_year(year))},
                    {"chave": f"arquivo_extraido_{year}", "valor": str(config.extracted_dir_for_year(year))},
                ]
            )
        for _, row in stats.iterrows():
            year = str(row["ano_referencia"])
            metadata_rows.extend(
                [
                    {"chave": f"linhas_lidas_total_{year}", "valor": int(row["linhas_lidas_total"])},
                    {"chave": f"linhas_filtradas_sc_{year}", "valor": int(row["linhas_filtradas_sc"])},
                    {
                        "chave": f"linhas_validas_sem_rais_negativa_{year}",
                        "valor": int(row["linhas_validas_sem_rais_negativa"]),
                    },
                    {"chave": f"linhas_rais_negativa_{year}", "valor": int(row["linhas_rais_negativa"])},
                ]
            )
        metadata = pd.DataFrame(metadata_rows)
        write_table(writer, "Metadados", metadata)
    return config.output_workbook_path


def export_to_sqlite(
    grouped_data: dict[str, pd.DataFrame],
    negative_grouped_data: dict[str, pd.DataFrame],
    stats: pd.DataFrame,
    config: PipelineConfig,
) -> Path:
    LOGGER.info("Gerando base analitica em %s", config.output_database_path)
    with sqlite3.connect(config.output_database_path) as connection:
        for table_name, dataframe in grouped_data.items():
            dataframe.to_sql(table_name, connection, if_exists="replace", index=False)
        for table_name, dataframe in negative_grouped_data.items():
            dataframe.to_sql(f"neg_{table_name}", connection, if_exists="replace", index=False)
        stats.to_sql("processamento_anos", connection, if_exists="replace", index=False)
    return config.output_database_path


def run_pipeline(config: PipelineConfig) -> Path:
    ensure_directories(config)
    cnae_dimension = load_cnae_dimension(config.cnae_dimension_path)
    municipality_reference = build_or_load_municipality_reference(config)
    grouped_by_year: dict[str, list[pd.DataFrame]] = {
        "municipio_divisao": [],
        "mesorregiao_divisao": [],
        "sc_competitiva_municipio": [],
        "sc_competitiva_mesorregiao": [],
        "resumo_divisao": [],
        "resumo_municipio": [],
        "resumo_mesorregiao": [],
        "resumo_sc_competitiva": [],
    }
    negative_by_year: dict[str, list[pd.DataFrame]] = {
        "municipio_divisao": [],
        "mesorregiao_divisao": [],
        "sc_competitiva_municipio": [],
        "sc_competitiva_mesorregiao": [],
        "resumo_divisao": [],
        "resumo_municipio": [],
        "resumo_mesorregiao": [],
        "resumo_sc_competitiva": [],
        "resumo_rais_negativa_codigo": [],
    }
    stats_rows: list[dict[str, int | str]] = []

    for year in config.years:
        download_rais_archive(config, year)
        extracted_file_path = extract_7z_archive(config, year)
        grouped_data, negative_grouped_data, stats = process_rais_file(
            config=config,
            year=year,
            extracted_file_path=extracted_file_path,
            cnae_dimension=cnae_dimension,
            municipality_reference=municipality_reference,
        )
        stats_rows.append(stats)

        for key, dataframe in grouped_data.items():
            grouped_by_year[key].append(dataframe.assign(ano_referencia=year))
        for key, dataframe in negative_grouped_data.items():
            negative_by_year[key].append(dataframe.assign(ano_referencia=year))

    grouped_data_final = {
        key: pd.concat(value, ignore_index=True) if value else pd.DataFrame()
        for key, value in grouped_by_year.items()
    }
    negative_grouped_data_final = {
        key: pd.concat(value, ignore_index=True) if value else pd.DataFrame()
        for key, value in negative_by_year.items()
    }

    ordered_columns = {
        "municipio_divisao": [
            "ano_referencia",
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
        ]
        + BASE_METRIC_COLUMNS,
        "mesorregiao_divisao": [
            "ano_referencia",
            "mesorregiao_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
        ]
        + BASE_METRIC_COLUMNS,
        "sc_competitiva_municipio": [
            "ano_referencia",
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "sc_competitiva",
        ]
        + BASE_METRIC_COLUMNS,
        "sc_competitiva_mesorregiao": [
            "ano_referencia",
            "mesorregiao_nome",
            "sc_competitiva",
        ]
        + BASE_METRIC_COLUMNS,
        "resumo_divisao": ["ano_referencia", "cnae_divisao_codigo", "cnae_divisao_nome"] + BASE_METRIC_COLUMNS,
        "resumo_municipio": ["ano_referencia", "municipio_codigo", "municipio_nome", "mesorregiao_nome"]
        + BASE_METRIC_COLUMNS,
        "resumo_mesorregiao": ["ano_referencia", "mesorregiao_nome"] + BASE_METRIC_COLUMNS,
        "resumo_sc_competitiva": ["ano_referencia", "sc_competitiva"] + BASE_METRIC_COLUMNS,
        "resumo_rais_negativa_codigo": ["ano_referencia", "ind_rais_negativa"] + BASE_METRIC_COLUMNS,
    }

    for collection in [grouped_data_final, negative_grouped_data_final]:
        for key, columns in ordered_columns.items():
            if key in collection and not collection[key].empty:
                collection[key] = collection[key][columns].sort_values(columns[: len(columns) - len(BASE_METRIC_COLUMNS)]).reset_index(drop=True)

    stats_df = pd.DataFrame(stats_rows).sort_values("ano_referencia").reset_index(drop=True)
    export_to_sqlite(grouped_data_final, negative_grouped_data_final, stats_df, config)
    return export_to_excel(grouped_data_final, negative_grouped_data_final, stats_df, config)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Coleta e analisa os microdados da RAIS de estabelecimentos para Santa Catarina."
    )
    parser.add_argument(
        "--year",
        default=None,
        help="Ano de referencia da RAIS no FTP oficial. Mantido por compatibilidade para execucao de um unico ano.",
    )
    parser.add_argument(
        "--years",
        nargs="+",
        default=None,
        help="Lista de anos para consolidar no mesmo workbook. Exemplo: --years 2022 2023 2024",
    )
    parser.add_argument(
        "--project-root",
        default=str(Path(__file__).resolve().parents[1]),
        help="Diretorio raiz do projeto.",
    )
    parser.add_argument(
        "--cnae-dimension-path",
        default=None,
        help="Caminho para a planilha com a dimensao de CNAE e agrupamento SC Competitiva.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=200_000,
        help="Quantidade de linhas por chunk durante o processamento.",
    )
    parser.add_argument(
        "--force-download",
        action="store_true",
        help="Refaz o download mesmo se o arquivo bruto ja existir.",
    )
    parser.add_argument(
        "--force-extract",
        action="store_true",
        help="Refaz a extracao do arquivo 7z mesmo se ja houver arquivo extraido.",
    )
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()
    project_root = Path(args.project_root).resolve()
    years = [str(year) for year in (args.years or ([args.year] if args.year else ["2022", "2023", "2024"]))]
    cnae_dimension_path = (
        Path(args.cnae_dimension_path).resolve()
        if args.cnae_dimension_path
        else project_root / "data" / "cnae_dimensao.xlsx"
    )

    config = PipelineConfig(
        years=years,
        project_root=project_root,
        cnae_dimension_path=cnae_dimension_path,
        chunk_size=args.chunk_size,
        force_download=args.force_download,
        force_extract=args.force_extract,
    )

    workbook_path = run_pipeline(config)
    LOGGER.info("Pipeline concluido. Workbook final: %s", workbook_path)


if __name__ == "__main__":
    main()
