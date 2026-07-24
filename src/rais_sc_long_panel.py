from __future__ import annotations

import argparse
import logging
import unicodedata
from dataclasses import dataclass
from ftplib import FTP
from pathlib import Path

import pandas as pd
import py7zr
from tqdm import tqdm

from rais_sc_pipeline import (
    FTP_HOST,
    FTP_ROOT,
    PipelineConfig,
    build_or_load_municipality_reference,
    clean_digits,
    combine_grouped_frames,
    detect_csv_separator,
    download_rais_archive,
    ensure_directories,
    extract_7z_archive,
    load_cnae_dimension,
    process_rais_file,
    write_table,
)


LOGGER = logging.getLogger("rais_sc_long_panel")
VINC_ARCHIVE_NAME = "RAIS_VINC_PUB_SUL.7z"


@dataclass
class LongPanelConfig:
    years: list[str]
    project_root: Path
    cnae_dimension_path: Path
    chunk_size_estab: int = 200_000
    chunk_size_vinc: int = 500_000
    ftp_host: str = FTP_HOST
    ftp_root: str = FTP_ROOT
    force_download: bool = False
    force_extract: bool = False

    @property
    def output_dir(self) -> Path:
        return self.project_root / "data" / "output"

    @property
    def output_workbook_path(self) -> Path:
        if len(self.years) == 1:
            suffix = self.years[0]
        else:
            suffix = f"{self.years[0]}_{self.years[-1]}"
        return self.output_dir / f"rais_sc_municipio_divisao_long_{suffix}.xlsx"

    def vinc_raw_dir_for_year(self, year: str) -> Path:
        return self.project_root / "data" / "raw" / f"{year}_vinc_sul"

    def vinc_archive_path_for_year(self, year: str) -> Path:
        return self.vinc_raw_dir_for_year(year) / VINC_ARCHIVE_NAME

    def vinc_extracted_dir_for_year(self, year: str) -> Path:
        return self.vinc_raw_dir_for_year(year) / "extracted"

    def remote_directory_for_year(self, year: str) -> str:
        return f"{self.ftp_root}/{year}"


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def normalize_text(value: object) -> str:
    normalized = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return normalized.lower().strip()


def ensure_long_panel_directories(config: LongPanelConfig) -> None:
    ensure_directories(
        PipelineConfig(
            years=config.years,
            project_root=config.project_root,
            cnae_dimension_path=config.cnae_dimension_path,
            chunk_size=config.chunk_size_estab,
            force_download=config.force_download,
            force_extract=config.force_extract,
        )
    )
    for year in config.years:
        config.vinc_raw_dir_for_year(year).mkdir(parents=True, exist_ok=True)
        config.vinc_extracted_dir_for_year(year).mkdir(parents=True, exist_ok=True)
    config.output_dir.mkdir(parents=True, exist_ok=True)


def download_vinc_archive(config: LongPanelConfig, year: str) -> Path:
    destination = config.vinc_archive_path_for_year(year)
    if destination.exists() and not config.force_download:
        LOGGER.info("Arquivo bruto de vinculos ja existe em %s", destination)
        return destination

    LOGGER.info("Baixando RAIS vinculo SUL %s do FTP oficial", year)
    with FTP(config.ftp_host, timeout=120) as ftp:
        ftp.encoding = "latin-1"
        ftp.login()
        ftp.cwd(config.remote_directory_for_year(year))
        total_size = ftp.size(VINC_ARCHIVE_NAME)
        with destination.open("wb") as output_file:
            with tqdm(
                total=total_size,
                unit="B",
                unit_scale=True,
                desc=f"{year}_{VINC_ARCHIVE_NAME}",
            ) as progress:

                def write_chunk(chunk: bytes) -> None:
                    output_file.write(chunk)
                    progress.update(len(chunk))

                ftp.retrbinary(
                    f"RETR {VINC_ARCHIVE_NAME}",
                    write_chunk,
                    blocksize=1024 * 256,
                )

    LOGGER.info("Download concluido: %s", destination)
    return destination


def extract_vinc_archive(config: LongPanelConfig, year: str) -> Path:
    extracted_dir = config.vinc_extracted_dir_for_year(year)
    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if extracted_files and not config.force_extract:
        LOGGER.info("Arquivo de vinculos extraido ja encontrado em %s", extracted_dir)
        return extracted_files[0]

    LOGGER.info("Extraindo %s", config.vinc_archive_path_for_year(year).name)
    for file_path in extracted_files:
        file_path.unlink()

    with py7zr.SevenZipFile(config.vinc_archive_path_for_year(year), mode="r") as archive:
        archive.extractall(path=extracted_dir)

    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if not extracted_files:
        raise FileNotFoundError("Nenhum arquivo foi extraido do arquivo .7z da RAIS vinculo.")

    LOGGER.info("Extracao concluida: %s", extracted_files[0])
    return extracted_files[0]


def resolve_first_available_column(
    normalized_columns: dict[str, str],
    candidates: list[str],
) -> str | None:
    for candidate in candidates:
        matched_column = normalized_columns.get(candidate)
        if matched_column is not None:
            return matched_column
    return None


def detect_vinc_columns(file_path: Path, separator: str) -> tuple[dict[str, str], bool, bool]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    required_columns = {
        "municipio_codigo": ["municipio - codigo", "municipio"],
        "cnae_subclasse_codigo": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
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

    active_column = resolve_first_available_column(
        normalized_columns,
        ["ind vinculo ativo 31/12 - codigo", "vinculo ativo 31/12"],
    )
    has_active_3112_filter = active_column is not None
    if active_column is not None:
        resolved["ind_vinculo_ativo_3112"] = active_column

    return resolved, has_abandonment_filter, has_active_3112_filter


def aggregate_vinc_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
    has_abandonment_filter: bool,
    has_active_3112_filter: bool,
) -> pd.DataFrame:
    chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce")
    chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )

    filtered = chunk.loc[
        chunk["municipio_codigo"].notna()
        & chunk["municipio_codigo"].astype("Int64").astype(str).str.startswith("42")
    ].copy()
    if has_abandonment_filter:
        filtered["ind_vinculo_abandonado"] = pd.to_numeric(filtered["ind_vinculo_abandonado"], errors="coerce")
        filtered = filtered.loc[filtered["ind_vinculo_abandonado"].eq(0)].copy()
    if has_active_3112_filter:
        filtered["ind_vinculo_ativo_3112"] = pd.to_numeric(filtered["ind_vinculo_ativo_3112"], errors="coerce")
        filtered = filtered.loc[filtered["ind_vinculo_ativo_3112"].eq(1)].copy()
    if filtered.empty:
        return filtered

    filtered["municipio_codigo"] = filtered["municipio_codigo"].astype("int64")
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
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )
        .agg(qtd_vinculos=("qtd_vinculos", "sum"))
    )


def process_vinc_file(
    file_path: Path,
    cnae_dimension: pd.DataFrame,
    municipality_reference: pd.DataFrame,
    chunk_size: int,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    separator = detect_csv_separator(file_path)
    selected_columns, has_abandonment_filter, has_active_3112_filter = detect_vinc_columns(file_path, separator)
    LOGGER.info("Layout de vinculos detectado: separador '%s'", separator)

    stats = {
        "linhas_lidas_vinc_total": 0,
        "linhas_vinc_sc_utilizadas": 0,
        "filtro_vinculo_abandonado_aplicado": "sim" if has_abandonment_filter else "nao",
        "filtro_vinculo_ativo_3112_aplicado": "sim" if has_active_3112_filter else "nao",
    }
    grouped_frames: list[pd.DataFrame] = []

    reader = pd.read_csv(
        file_path,
        encoding="latin-1",
        sep=separator,
        usecols=list(selected_columns.values()),
        chunksize=chunk_size,
        low_memory=False,
    )

    for chunk_number, chunk in enumerate(reader, start=1):
        stats["linhas_lidas_vinc_total"] += len(chunk)
        renamed_chunk = chunk.rename(columns={source: target for target, source in selected_columns.items()})
        aggregated = aggregate_vinc_chunk(
            renamed_chunk,
            cnae_dimension,
            municipality_reference,
            has_abandonment_filter=has_abandonment_filter,
            has_active_3112_filter=has_active_3112_filter,
        )
        if aggregated.empty:
            LOGGER.info("Chunk vinculo %s sem registros validos de SC", chunk_number)
            continue

        stats["linhas_vinc_sc_utilizadas"] += int(aggregated["qtd_vinculos"].sum())
        grouped_frames.append(aggregated)
        LOGGER.info("Chunk vinculo %s processado", chunk_number)

    combined = combine_grouped_frames(
        grouped_frames,
        [
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
        ],
    )
    return combined, stats


def build_long_panel(
    estabelecimentos: pd.DataFrame,
    vinculos: pd.DataFrame,
) -> pd.DataFrame:
    estab_long = (
        estabelecimentos[
            [
                "ano_referencia",
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
                "estabelecimentos",
            ]
        ]
        .rename(columns={"estabelecimentos": "quantidade"})
        .assign(indicador="estabelecimentos")
    )
    vinc_long = (
        vinculos[
            [
                "ano_referencia",
                "municipio_codigo",
                "municipio_nome",
                "mesorregiao_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
                "qtd_vinculos",
            ]
        ]
        .rename(columns={"qtd_vinculos": "quantidade"})
        .assign(indicador="vinculos")
    )

    final = pd.concat([estab_long, vinc_long], ignore_index=True)
    final["cnae_divisao_codigo"] = final["cnae_divisao_codigo"].astype(str).str.zfill(2).where(
        final["cnae_divisao_codigo"].astype(str).ne("NA"),
        final["cnae_divisao_codigo"],
    )
    return final[
        [
            "ano_referencia",
            "municipio_codigo",
            "municipio_nome",
            "mesorregiao_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
            "indicador",
            "quantidade",
        ]
    ].sort_values(
        [
            "ano_referencia",
            "municipio_nome",
            "cnae_divisao_codigo",
            "indicador",
        ]
    ).reset_index(drop=True)


def build_metadata(stats_rows: list[dict[str, int | str]]) -> pd.DataFrame:
    metadata_rows: list[dict[str, object]] = [
        {
            "chave": "observacao_vinculos",
            "valor": "Vinculos contam registros da RAIS vinculo SUL com filtro municipio de SC. Quando as variaveis existirem no layout do ano, os filtros aplicados sao Ind Vinculo Abandonado = 0 e Vinculo Ativo 31/12 = 1.",
        },
        {
            "chave": "observacao_estabelecimentos",
            "valor": "Estabelecimentos contam registros validos da RAIS estabelecimentos, excluindo RAIS negativa.",
        },
    ]

    for row in stats_rows:
        year = str(row["ano_referencia"])
        for key, value in row.items():
            metadata_rows.append({"chave": f"{key}_{year}" if key != "ano_referencia" else key, "valor": value})

    return pd.DataFrame(metadata_rows)


def export_workbook(
    final: pd.DataFrame,
    metadata: pd.DataFrame,
    output_path: Path,
) -> Path:
    LOGGER.info("Gerando workbook final em %s", output_path)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        write_table(writer, "Municipio_Divisao_Long", final)
        write_table(writer, "Metadados", metadata)
    return output_path


def run_long_panel(config: LongPanelConfig) -> Path:
    ensure_long_panel_directories(config)
    cnae_dimension = load_cnae_dimension(config.cnae_dimension_path)
    municipality_reference = build_or_load_municipality_reference(
        PipelineConfig(
            years=config.years,
            project_root=config.project_root,
            cnae_dimension_path=config.cnae_dimension_path,
            chunk_size=config.chunk_size_estab,
            force_download=config.force_download,
            force_extract=config.force_extract,
        )
    )

    estab_frames: list[pd.DataFrame] = []
    vinc_frames: list[pd.DataFrame] = []
    stats_rows: list[dict[str, int | str]] = []

    for year in config.years:
        estab_pipeline_config = PipelineConfig(
            years=[year],
            project_root=config.project_root,
            cnae_dimension_path=config.cnae_dimension_path,
            chunk_size=config.chunk_size_estab,
            force_download=config.force_download,
            force_extract=config.force_extract,
        )
        download_rais_archive(estab_pipeline_config, year)
        estab_file = extract_7z_archive(estab_pipeline_config, year)
        grouped_data, _, estab_stats = process_rais_file(
            config=estab_pipeline_config,
            year=year,
            extracted_file_path=estab_file,
            cnae_dimension=cnae_dimension,
            municipality_reference=municipality_reference,
        )
        estab_frames.append(grouped_data["municipio_divisao"].assign(ano_referencia=year))

        download_vinc_archive(config, year)
        vinc_file = extract_vinc_archive(config, year)
        vinc_grouped, vinc_stats = process_vinc_file(
            file_path=vinc_file,
            cnae_dimension=cnae_dimension,
            municipality_reference=municipality_reference,
            chunk_size=config.chunk_size_vinc,
        )
        vinc_frames.append(vinc_grouped.assign(ano_referencia=year))

        stats_rows.append(
            {
                "ano_referencia": year,
                "linhas_lidas_estab_total": estab_stats["linhas_lidas_total"],
                "linhas_estab_sc_validas": estab_stats["linhas_validas_sem_rais_negativa"],
                "linhas_estab_sc_rais_negativa": estab_stats["linhas_rais_negativa"],
                "linhas_lidas_vinc_total": vinc_stats["linhas_lidas_vinc_total"],
                "linhas_vinc_sc_utilizadas": vinc_stats["linhas_vinc_sc_utilizadas"],
                "filtro_vinculo_abandonado_aplicado": vinc_stats["filtro_vinculo_abandonado_aplicado"],
                "filtro_vinculo_ativo_3112_aplicado": vinc_stats["filtro_vinculo_ativo_3112_aplicado"],
            }
        )

    estabelecimentos = pd.concat(estab_frames, ignore_index=True)
    vinculos = pd.concat(vinc_frames, ignore_index=True)
    final = build_long_panel(estabelecimentos, vinculos)
    metadata = build_metadata(stats_rows)
    return export_workbook(final, metadata, config.output_workbook_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera painel long por municipio e divisao CNAE para SC com estabelecimentos e vinculos nao abandonados."
    )
    parser.add_argument(
        "--years",
        nargs="+",
        default=["2022", "2023", "2024", "2025"],
        help="Lista de anos para processar. Exemplo: --years 2022 2023 2024 2025",
    )
    parser.add_argument(
        "--project-root",
        default=str(Path(__file__).resolve().parents[1]),
        help="Diretorio raiz do projeto.",
    )
    parser.add_argument(
        "--cnae-dimension-path",
        default=None,
        help="Caminho para a planilha de dimensao CNAE.",
    )
    parser.add_argument(
        "--chunk-size-estab",
        type=int,
        default=200_000,
        help="Quantidade de linhas por chunk na RAIS estabelecimentos.",
    )
    parser.add_argument(
        "--chunk-size-vinc",
        type=int,
        default=500_000,
        help="Quantidade de linhas por chunk na RAIS vinculo.",
    )
    parser.add_argument(
        "--force-download",
        action="store_true",
        help="Refaz o download dos arquivos brutos mesmo se ja existirem.",
    )
    parser.add_argument(
        "--force-extract",
        action="store_true",
        help="Refaz a extracao dos arquivos 7z mesmo se ja existirem.",
    )
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()
    project_root = Path(args.project_root).resolve()
    cnae_dimension_path = (
        Path(args.cnae_dimension_path).resolve()
        if args.cnae_dimension_path
        else project_root / "data" / "cnae_dimensao.xlsx"
    )
    config = LongPanelConfig(
        years=[str(year) for year in args.years],
        project_root=project_root,
        cnae_dimension_path=cnae_dimension_path,
        chunk_size_estab=args.chunk_size_estab,
        chunk_size_vinc=args.chunk_size_vinc,
        force_download=args.force_download,
        force_extract=args.force_extract,
    )
    output_path = run_long_panel(config)
    LOGGER.info("Painel concluido. Arquivo final: %s", output_path)


if __name__ == "__main__":
    main()
