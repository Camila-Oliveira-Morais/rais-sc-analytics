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
    clean_digits,
    combine_grouped_frames,
    detect_csv_separator,
    download_rais_archive,
    ensure_directories,
    extract_7z_archive,
    load_cnae_dimension,
    write_table,
)


LOGGER = logging.getLogger("rais_brasil_uf_divisao_panel")

VINC_ARCHIVES = [
    ("centro_oeste", "RAIS_VINC_PUB_CENTRO_OESTE.7z"),
    ("mg_es_rj", "RAIS_VINC_PUB_MG_ES_RJ.7z"),
    ("ni", "RAIS_VINC_PUB_NI.7z"),
    ("nordeste", "RAIS_VINC_PUB_NORDESTE.7z"),
    ("norte", "RAIS_VINC_PUB_NORTE.7z"),
    ("sp", "RAIS_VINC_PUB_SP.7z"),
    ("sul", "RAIS_VINC_PUB_SUL.7z"),
]

UF_REFERENCE = {
    11: ("RO", "Rondonia"),
    12: ("AC", "Acre"),
    13: ("AM", "Amazonas"),
    14: ("RR", "Roraima"),
    15: ("PA", "Para"),
    16: ("AP", "Amapa"),
    17: ("TO", "Tocantins"),
    21: ("MA", "Maranhao"),
    22: ("PI", "Piaui"),
    23: ("CE", "Ceara"),
    24: ("RN", "Rio Grande do Norte"),
    25: ("PB", "Paraiba"),
    26: ("PE", "Pernambuco"),
    27: ("AL", "Alagoas"),
    28: ("SE", "Sergipe"),
    29: ("BA", "Bahia"),
    31: ("MG", "Minas Gerais"),
    32: ("ES", "Espirito Santo"),
    33: ("RJ", "Rio de Janeiro"),
    35: ("SP", "Sao Paulo"),
    41: ("PR", "Parana"),
    42: ("SC", "Santa Catarina"),
    43: ("RS", "Rio Grande do Sul"),
    50: ("MS", "Mato Grosso do Sul"),
    51: ("MT", "Mato Grosso"),
    52: ("GO", "Goias"),
    53: ("DF", "Distrito Federal"),
}


@dataclass
class BrasilUfDivisaoConfig:
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
        return self.output_dir / f"rais_brasil_uf_divisao_{suffix}.xlsx"

    def vinc_raw_dir_for_year_region(self, year: str, region_key: str) -> Path:
        return self.project_root / "data" / "raw" / f"{year}_vinc_{region_key}"

    def vinc_archive_path_for_year_region(self, year: str, region_key: str, archive_name: str) -> Path:
        return self.vinc_raw_dir_for_year_region(year, region_key) / archive_name

    def vinc_extracted_dir_for_year_region(self, year: str, region_key: str) -> Path:
        return self.vinc_raw_dir_for_year_region(year, region_key) / "extracted"

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


def ensure_brasil_directories(config: BrasilUfDivisaoConfig) -> None:
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
        for region_key, _ in VINC_ARCHIVES:
            config.vinc_raw_dir_for_year_region(year, region_key).mkdir(parents=True, exist_ok=True)
            config.vinc_extracted_dir_for_year_region(year, region_key).mkdir(parents=True, exist_ok=True)
    config.output_dir.mkdir(parents=True, exist_ok=True)


def resolve_first_available_column(
    normalized_columns: dict[str, str],
    candidates: list[str],
) -> str | None:
    for candidate in candidates:
        matched_column = normalized_columns.get(candidate)
        if matched_column is not None:
            return matched_column
    return None


def download_vinc_archive(
    config: BrasilUfDivisaoConfig,
    year: str,
    region_key: str,
    archive_name: str,
) -> Path:
    destination = config.vinc_archive_path_for_year_region(year, region_key, archive_name)
    if destination.exists() and not config.force_download:
        LOGGER.info("Arquivo bruto de vinculos ja existe em %s", destination)
        return destination

    LOGGER.info("Baixando %s de %s do FTP oficial", archive_name, year)
    with FTP(config.ftp_host, timeout=120) as ftp:
        ftp.encoding = "latin-1"
        ftp.login()
        ftp.cwd(config.remote_directory_for_year(year))
        total_size = ftp.size(archive_name)
        with destination.open("wb") as output_file:
            with tqdm(
                total=total_size,
                unit="B",
                unit_scale=True,
                desc=f"{year}_{archive_name}",
            ) as progress:

                def write_chunk(chunk: bytes) -> None:
                    output_file.write(chunk)
                    progress.update(len(chunk))

                ftp.retrbinary(
                    f"RETR {archive_name}",
                    write_chunk,
                    blocksize=1024 * 256,
                )

    LOGGER.info("Download concluido: %s", destination)
    return destination


def extract_vinc_archive(
    config: BrasilUfDivisaoConfig,
    year: str,
    region_key: str,
    archive_name: str,
) -> Path:
    extracted_dir = config.vinc_extracted_dir_for_year_region(year, region_key)
    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if extracted_files and not config.force_extract:
        LOGGER.info("Arquivo de vinculos extraido ja encontrado em %s", extracted_dir)
        return extracted_files[0]

    LOGGER.info("Extraindo %s", archive_name)
    for file_path in extracted_files:
        file_path.unlink()

    with py7zr.SevenZipFile(
        config.vinc_archive_path_for_year_region(year, region_key, archive_name),
        mode="r",
    ) as archive:
        archive.extractall(path=extracted_dir)

    extracted_files = [path for path in extracted_dir.iterdir() if path.is_file()]
    if not extracted_files:
        raise FileNotFoundError("Nenhum arquivo foi extraido do arquivo .7z da RAIS vinculo.")

    LOGGER.info("Extracao concluida: %s", extracted_files[0])
    return extracted_files[0]


def detect_estab_columns(file_path: Path, separator: str) -> dict[str, str]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}
    required_columns = {
        "uf_codigo": ["uf - codigo", "uf"],
        "cnae_subclasse_codigo": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
        "ind_rais_negativa": ["ind rais negativa - codigo", "ind rais negativa"],
    }

    resolved: dict[str, str] = {}
    for alias, candidates in required_columns.items():
        matched_column = resolve_first_available_column(normalized_columns, candidates)
        if matched_column is None:
            raise ValueError(f"Coluna obrigatoria nao encontrada: {candidates[0]}")
        resolved[alias] = matched_column

    return resolved


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


def enrich_uf_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    uf_ref = pd.DataFrame(
        [
            {"uf_codigo": uf_codigo, "uf_sigla": uf_sigla, "uf_nome": uf_nome}
            for uf_codigo, (uf_sigla, uf_nome) in UF_REFERENCE.items()
        ]
    )
    enriched = dataframe.merge(uf_ref, on="uf_codigo", how="left")
    enriched["uf_sigla"] = enriched["uf_sigla"].fillna("NA")
    enriched["uf_nome"] = enriched["uf_nome"].fillna("UF nao identificada")
    return enriched


def aggregate_estab_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
) -> pd.DataFrame:
    chunk["uf_codigo"] = pd.to_numeric(chunk["uf_codigo"], errors="coerce")
    chunk["ind_rais_negativa"] = pd.to_numeric(chunk["ind_rais_negativa"], errors="coerce")
    chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )

    filtered = chunk.loc[
        chunk["uf_codigo"].isin(list(UF_REFERENCE.keys()))
        & chunk["ind_rais_negativa"].eq(0)
    ].copy()
    if filtered.empty:
        return filtered

    filtered["uf_codigo"] = filtered["uf_codigo"].astype("int64")
    filtered["estabelecimentos"] = 1
    enriched = filtered.merge(cnae_dimension, on="cnae_subclasse_codigo", how="left")
    enriched = enrich_uf_columns(enriched)

    enriched["cnae_divisao_codigo"] = enriched["cnae_divisao_codigo"].fillna("NA")
    enriched["cnae_divisao_nome"] = enriched["cnae_divisao_nome"].fillna("Divisao nao identificada")

    return (
        enriched.groupby(
            [
                "uf_codigo",
                "uf_sigla",
                "uf_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )
        .agg(estabelecimentos=("estabelecimentos", "sum"))
    )


def aggregate_vinc_chunk(
    chunk: pd.DataFrame,
    cnae_dimension: pd.DataFrame,
    has_abandonment_filter: bool,
    has_active_3112_filter: bool,
) -> pd.DataFrame:
    chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce")
    chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].apply(
        lambda value: clean_digits(value, width=7)
    )

    filtered = chunk.loc[chunk["municipio_codigo"].notna()].copy()
    filtered["uf_codigo"] = filtered["municipio_codigo"].floordiv(10000).astype("Int64")
    filtered = filtered.loc[filtered["uf_codigo"].isin(list(UF_REFERENCE.keys()))].copy()

    if has_abandonment_filter:
        filtered["ind_vinculo_abandonado"] = pd.to_numeric(filtered["ind_vinculo_abandonado"], errors="coerce")
        filtered = filtered.loc[filtered["ind_vinculo_abandonado"].eq(0)].copy()
    if has_active_3112_filter:
        filtered["ind_vinculo_ativo_3112"] = pd.to_numeric(filtered["ind_vinculo_ativo_3112"], errors="coerce")
        filtered = filtered.loc[filtered["ind_vinculo_ativo_3112"].eq(1)].copy()
    if filtered.empty:
        return filtered

    filtered["uf_codigo"] = filtered["uf_codigo"].astype("int64")
    filtered["vinculos"] = 1
    enriched = filtered.merge(cnae_dimension, on="cnae_subclasse_codigo", how="left")
    enriched = enrich_uf_columns(enriched)

    enriched["cnae_divisao_codigo"] = enriched["cnae_divisao_codigo"].fillna("NA")
    enriched["cnae_divisao_nome"] = enriched["cnae_divisao_nome"].fillna("Divisao nao identificada")

    return (
        enriched.groupby(
            [
                "uf_codigo",
                "uf_sigla",
                "uf_nome",
                "cnae_divisao_codigo",
                "cnae_divisao_nome",
            ],
            as_index=False,
        )
        .agg(vinculos=("vinculos", "sum"))
    )


def process_estab_file(
    file_path: Path,
    cnae_dimension: pd.DataFrame,
    chunk_size: int,
) -> tuple[pd.DataFrame, dict[str, int]]:
    separator = detect_csv_separator(file_path)
    selected_columns = detect_estab_columns(file_path, separator)
    LOGGER.info("Layout de estabelecimentos detectado: separador '%s'", separator)

    stats = {
        "linhas_lidas_estab_total": 0,
        "linhas_estab_validas": 0,
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
        stats["linhas_lidas_estab_total"] += len(chunk)
        renamed_chunk = chunk.rename(columns={source: target for target, source in selected_columns.items()})
        aggregated = aggregate_estab_chunk(renamed_chunk, cnae_dimension)
        if aggregated.empty:
            LOGGER.info("Chunk estab %s sem registros validos", chunk_number)
            continue

        stats["linhas_estab_validas"] += int(aggregated["estabelecimentos"].sum())
        grouped_frames.append(aggregated)
        LOGGER.info("Chunk estab %s processado", chunk_number)

    combined = combine_grouped_frames(
        grouped_frames,
        ["uf_codigo", "uf_sigla", "uf_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
    )
    return combined, stats


def process_vinc_file(
    file_path: Path,
    cnae_dimension: pd.DataFrame,
    chunk_size: int,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    separator = detect_csv_separator(file_path)
    selected_columns, has_abandonment_filter, has_active_3112_filter = detect_vinc_columns(file_path, separator)
    LOGGER.info("Layout de vinculos detectado: separador '%s'", separator)

    stats = {
        "linhas_lidas_vinc_total": 0,
        "linhas_vinc_utilizadas": 0,
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
            has_abandonment_filter=has_abandonment_filter,
            has_active_3112_filter=has_active_3112_filter,
        )
        if aggregated.empty:
            LOGGER.info("Chunk vinculo %s sem registros validos", chunk_number)
            continue

        stats["linhas_vinc_utilizadas"] += int(aggregated["vinculos"].sum())
        grouped_frames.append(aggregated)
        LOGGER.info("Chunk vinculo %s processado", chunk_number)

    combined = combine_grouped_frames(
        grouped_frames,
        ["uf_codigo", "uf_sigla", "uf_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
    )
    return combined, stats


def build_wide_panel(estabelecimentos: pd.DataFrame, vinculos: pd.DataFrame) -> pd.DataFrame:
    merged = estabelecimentos.merge(
        vinculos,
        on=[
            "ano_referencia",
            "uf_codigo",
            "uf_sigla",
            "uf_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
        ],
        how="outer",
    )
    merged["estabelecimentos"] = merged["estabelecimentos"].fillna(0).astype("int64")
    merged["vinculos"] = merged["vinculos"].fillna(0).astype("int64")
    merged["cnae_divisao_codigo"] = merged["cnae_divisao_codigo"].astype(str).str.zfill(2).where(
        merged["cnae_divisao_codigo"].astype(str).ne("NA"),
        merged["cnae_divisao_codigo"],
    )
    return merged[
        [
            "ano_referencia",
            "uf_codigo",
            "uf_sigla",
            "uf_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
            "estabelecimentos",
            "vinculos",
        ]
    ].sort_values(
        ["ano_referencia", "uf_codigo", "cnae_divisao_codigo"]
    ).reset_index(drop=True)


def build_long_panel(wide_panel: pd.DataFrame) -> pd.DataFrame:
    long_panel = wide_panel.melt(
        id_vars=[
            "ano_referencia",
            "uf_codigo",
            "uf_sigla",
            "uf_nome",
            "cnae_divisao_codigo",
            "cnae_divisao_nome",
        ],
        value_vars=["estabelecimentos", "vinculos"],
        var_name="indicador",
        value_name="quantidade",
    )
    return long_panel.sort_values(
        ["ano_referencia", "uf_codigo", "cnae_divisao_codigo", "indicador"]
    ).reset_index(drop=True)


def build_metadata(stats_rows: list[dict[str, int | str]]) -> pd.DataFrame:
    metadata_rows: list[dict[str, object]] = [
        {
            "chave": "observacao_vinculos",
            "valor": "Vinculos contam registros da RAIS vinculo com agregacao por UF e divisao CNAE. Quando as variaveis existirem no layout do ano, os filtros aplicados sao Ind Vinculo Abandonado = 0 e Vinculo Ativo 31/12 = 1.",
        },
        {
            "chave": "observacao_estabelecimentos",
            "valor": "Estabelecimentos contam registros validos da RAIS estabelecimentos, excluindo RAIS negativa.",
        },
        {
            "chave": "recorte_geografico",
            "valor": "Todas as UFs do Brasil.",
        },
    ]

    for row in stats_rows:
        year = str(row["ano_referencia"])
        for key, value in row.items():
            metadata_rows.append({"chave": f"{key}_{year}" if key != "ano_referencia" else key, "valor": value})

    return pd.DataFrame(metadata_rows)


def export_workbook(
    wide_panel: pd.DataFrame,
    long_panel: pd.DataFrame,
    metadata: pd.DataFrame,
    output_path: Path,
) -> Path:
    LOGGER.info("Gerando workbook final em %s", output_path)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        write_table(writer, "UF_Divisao_Wide", wide_panel)
        write_table(writer, "UF_Divisao_Long", long_panel)
        write_table(writer, "Metadados", metadata)
    return output_path


def run_brasil_uf_divisao_panel(config: BrasilUfDivisaoConfig) -> Path:
    ensure_brasil_directories(config)
    cnae_dimension = load_cnae_dimension(config.cnae_dimension_path)

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
        estab_grouped, estab_stats = process_estab_file(
            file_path=estab_file,
            cnae_dimension=cnae_dimension,
            chunk_size=config.chunk_size_estab,
        )
        estab_frames.append(estab_grouped.assign(ano_referencia=year))

        vinc_year_frames: list[pd.DataFrame] = []
        vinc_year_stats = {
            "linhas_lidas_vinc_total": 0,
            "linhas_vinc_utilizadas": 0,
            "filtro_vinculo_abandonado_aplicado": "sim",
            "filtro_vinculo_ativo_3112_aplicado": "sim",
        }
        for region_key, archive_name in VINC_ARCHIVES:
            download_vinc_archive(config, year, region_key, archive_name)
            vinc_file = extract_vinc_archive(config, year, region_key, archive_name)
            vinc_grouped, vinc_stats = process_vinc_file(
                file_path=vinc_file,
                cnae_dimension=cnae_dimension,
                chunk_size=config.chunk_size_vinc,
            )
            if not vinc_grouped.empty:
                vinc_year_frames.append(vinc_grouped)
            vinc_year_stats["linhas_lidas_vinc_total"] += int(vinc_stats["linhas_lidas_vinc_total"])
            vinc_year_stats["linhas_vinc_utilizadas"] += int(vinc_stats["linhas_vinc_utilizadas"])
            if vinc_stats["filtro_vinculo_abandonado_aplicado"] == "nao":
                vinc_year_stats["filtro_vinculo_abandonado_aplicado"] = "nao"
            if vinc_stats["filtro_vinculo_ativo_3112_aplicado"] == "nao":
                vinc_year_stats["filtro_vinculo_ativo_3112_aplicado"] = "nao"

        vinc_combined = combine_grouped_frames(
            vinc_year_frames,
            ["uf_codigo", "uf_sigla", "uf_nome", "cnae_divisao_codigo", "cnae_divisao_nome"],
        )
        vinc_frames.append(vinc_combined.assign(ano_referencia=year))

        stats_rows.append(
            {
                "ano_referencia": year,
                "linhas_lidas_estab_total": estab_stats["linhas_lidas_estab_total"],
                "linhas_estab_validas": estab_stats["linhas_estab_validas"],
                "linhas_lidas_vinc_total": vinc_year_stats["linhas_lidas_vinc_total"],
                "linhas_vinc_utilizadas": vinc_year_stats["linhas_vinc_utilizadas"],
                "filtro_vinculo_abandonado_aplicado": vinc_year_stats["filtro_vinculo_abandonado_aplicado"],
                "filtro_vinculo_ativo_3112_aplicado": vinc_year_stats["filtro_vinculo_ativo_3112_aplicado"],
            }
        )

    estabelecimentos = pd.concat(estab_frames, ignore_index=True) if estab_frames else pd.DataFrame()
    vinculos = pd.concat(vinc_frames, ignore_index=True) if vinc_frames else pd.DataFrame()
    wide_panel = build_wide_panel(estabelecimentos, vinculos)
    long_panel = build_long_panel(wide_panel)
    metadata = build_metadata(stats_rows)
    return export_workbook(wide_panel, long_panel, metadata, config.output_workbook_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera painel RAIS Brasil por UF e divisao CNAE com vinculos e estabelecimentos."
    )
    parser.add_argument(
        "--years",
        nargs="+",
        default=["2023", "2024", "2025"],
        help="Lista de anos para processar. Padrao 2023 2024 2025 porque nesses layouts o filtro de abandono esta disponivel.",
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
    config = BrasilUfDivisaoConfig(
        years=[str(year) for year in args.years],
        project_root=project_root,
        cnae_dimension_path=cnae_dimension_path,
        chunk_size_estab=args.chunk_size_estab,
        chunk_size_vinc=args.chunk_size_vinc,
        force_download=args.force_download,
        force_extract=args.force_extract,
    )
    output_path = run_brasil_uf_divisao_panel(config)
    LOGGER.info("Painel concluido. Arquivo final: %s", output_path)


if __name__ == "__main__":
    main()
