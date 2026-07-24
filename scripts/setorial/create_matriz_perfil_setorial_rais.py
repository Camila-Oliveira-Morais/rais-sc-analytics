from __future__ import annotations

import argparse
import logging
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter


LOGGER = logging.getLogger("matriz_perfil_setorial_rais")

PROJECT_ROOT = Path(__file__).resolve().parents[2]
RAW_DIR = PROJECT_ROOT / "data" / "raw"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
CBO_DICTIONARY_FILE = PROJECT_ROOT / "data" / "dict" / "dicionario_cbo.xlsx"
TARGET_YEARS = [2022, 2025]
TARGET_CNAES = {
    "Alimentos e bebidas": ["10"],
    "Têxtil": ["13"],
    "Madeira": ["16"],
    "Metalurgia": ["24"],
    "Equipamentos elétricos": ["27"],
    "Veículos e metalmecânica": ["29"],
    "Energia": ["35"],
    "Tecnologia da informação": ["62"],
    "Serviços administrativos e apoio": ["82"],
}

BASE_ALIAS_MAP = {
    "ano": ["ano", "ano_referencia"],
    "uf": ["uf"],
    "sigla_uf": ["sigla_uf", "uf_sigla"],
    "municipio": ["municipio", "municipio_nome"],
    "cod_municipio": ["cod_municipio", "municipio_codigo", "municipio - codigo", "municipio"],
    "cnae_divisao": ["cnae_divisao", "cnae_2_divisao", "cnae_div", "cnae divisao", "cnae 2.0 divisao - codigo"],
    "cnae_subclasse": ["cnae_subclasse", "cnae_subclasse_codigo", "cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
    "cbo_codigo": ["cbo_codigo", "cbo", "cbo 2002 ocupacao - codigo", "cbo ocupacao 2002", "cbo 2002 ocupacao"],
    "cbo_descricao": ["cbo_descricao", "cbo_nome", "ocupacao", "cbo_ocupacao_nome"],
    "sexo": ["sexo", "sexo_codigo", "sexo - codigo", "sexo trabalhador"],
    "idade": ["idade"],
    "faixa_etaria": ["faixa_etaria", "faixa etaria", "faixa etaria - codigo", "faixa etária", "faixa etária - código"],
    "escolaridade": [
        "escolaridade",
        "grau_instrucao",
        "grau instrucao",
        "escolaridade apos 2005",
        "escolaridade apos 2005 - codigo",
    ],
    "remuneracao_media": [
        "remuneracao_media",
        "remun_media",
        "vl_rem_dezembro_nom",
        "vl rem dezembro nom",
        "vl remun dezembro nom",
        "vl_remun_media_nom",
        "vl rem media nom",
        "vl remun media nom",
    ],
    "vinculo_ativo_3112": [
        "vinculo_ativo_3112",
        "ind vinculo ativo 31/12 - codigo",
        "vinculo ativo 31/12",
    ],
    "ind_vinculo_abandonado": [
        "ind_vinculo_abandonado",
        "ind vinculo abandonado - codigo",
        "ind vinculo abandonado",
    ],
    "qtd_vinculos": ["qtd_vinculos", "quantidade_vinculos", "peso_vinculos"],
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

    if isinstance(value, (int, np.integer)):
        formatted = str(int(value))
    elif isinstance(value, (float, np.floating)):
        if float(value).is_integer():
            formatted = str(int(value))
        else:
            formatted = str(value)
    else:
        formatted = str(value)

    digits = "".join(char for char in formatted if char.isdigit())
    if not digits:
        return None
    return digits.zfill(width) if width is not None else digits


def parse_numeric_series(series: pd.Series) -> pd.Series:
    direto = pd.to_numeric(series, errors="coerce")
    if direto.notna().sum() >= max(1, int(series.notna().sum() * 0.5)):
        return direto

    texto = series.astype(str).str.strip()
    convertido = pd.to_numeric(
        texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    if convertido.notna().sum() > direto.notna().sum():
        return convertido
    return direto


def encontrar_coluna(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized_columns = {normalize_text(column): column for column in df.columns}
    for alias in aliases:
        matched = normalized_columns.get(normalize_text(alias))
        if matched is not None:
            return matched
    return None


def mapear_grupo_setorial(cnae_divisao: object) -> str | None:
    cnae_limpo = clean_digits(cnae_divisao, width=2)
    if cnae_limpo is None:
        return None
    for grupo, divisoes in TARGET_CNAES.items():
        if cnae_limpo in divisoes:
            return grupo
    return None


def padronizar_sexo(valor: object) -> str:
    if pd.isna(valor):
        return "Não informado"

    normalized = normalize_text(valor)
    digits = clean_digits(valor)

    if digits == "1" or normalized in {"m", "masculino", "homem"}:
        return "Masculino"
    if digits == "2" or normalized in {"f", "feminino", "mulher"}:
        return "Feminino"
    return "Não informado"


def criar_faixa_etaria(idade: object) -> str:
    idade_numerica = pd.to_numeric(pd.Series([idade]), errors="coerce").iloc[0]
    if pd.isna(idade_numerica):
        return "Não informado"
    if idade_numerica <= 24:
        return "Até 24 anos"
    if idade_numerica <= 29:
        return "25 a 29 anos"
    if idade_numerica <= 39:
        return "30 a 39 anos"
    if idade_numerica <= 49:
        return "40 a 49 anos"
    return "50 anos ou mais"


def padronizar_faixa_etaria_existente(valor: object) -> str:
    if pd.isna(valor):
        return "Não informado"

    normalized = normalize_text(valor)
    digits = clean_digits(valor)

    mapping_by_code = {
        "1": "Até 24 anos",
        "2": "Até 24 anos",
        "3": "25 a 29 anos",
        "4": "30 a 39 anos",
        "5": "40 a 49 anos",
        "6": "50 anos ou mais",
        "7": "50 anos ou mais",
        "8": "50 anos ou mais",
        "99": "Não informado",
    }
    if digits in mapping_by_code:
        return mapping_by_code[digits]

    if "24" in normalized or "jov" in normalized:
        return "Até 24 anos"
    if "25" in normalized and "29" in normalized:
        return "25 a 29 anos"
    if "30" in normalized and "39" in normalized:
        return "30 a 39 anos"
    if "40" in normalized and "49" in normalized:
        return "40 a 49 anos"
    if "50" in normalized or "mais" in normalized:
        return "50 anos ou mais"
    return "Não informado"


def padronizar_escolaridade(valor: object) -> str:
    if pd.isna(valor):
        return "Não informado"

    normalized = normalize_text(valor)
    digits = clean_digits(valor)

    mapping_by_code = {
        "1": "Fundamental incompleto",
        "2": "Fundamental incompleto",
        "3": "Fundamental incompleto",
        "4": "Fundamental incompleto",
        "5": "Fundamental completo",
        "6": "Médio incompleto",
        "7": "Médio completo",
        "8": "Superior incompleto",
        "9": "Superior completo ou mais",
        "10": "Superior completo ou mais",
        "11": "Superior completo ou mais",
    }
    if digits in mapping_by_code:
        return mapping_by_code[digits]

    if "analf" in normalized:
        return "Fundamental incompleto"
    if "fund" in normalized and "incomp" in normalized:
        return "Fundamental incompleto"
    if "fund" in normalized and "comp" in normalized:
        return "Fundamental completo"
    if "medio" in normalized and "incomp" in normalized:
        return "Médio incompleto"
    if "medio" in normalized and "comp" in normalized:
        return "Médio completo"
    if "super" in normalized and "incomp" in normalized:
        return "Superior incompleto"
    if any(token in normalized for token in {"super", "gradu", "mestr", "doutor", "pos"}) and "incomp" not in normalized:
        return "Superior completo ou mais"
    return "Não informado"


def categoria_predominante(
    df: pd.DataFrame,
    coluna_categoria: str,
    coluna_peso: str | None = None,
) -> tuple[str, float]:
    if df.empty or coluna_categoria not in df.columns:
        return "Sem dados", np.nan

    dados = df[[coluna_categoria] + ([coluna_peso] if coluna_peso else [])].copy()
    dados[coluna_categoria] = dados[coluna_categoria].fillna("Não informado")

    if coluna_peso and coluna_peso in dados.columns:
        dados[coluna_peso] = pd.to_numeric(dados[coluna_peso], errors="coerce").fillna(0)
        agrupado = dados.groupby(coluna_categoria, as_index=False)[coluna_peso].sum()
        total = float(agrupado[coluna_peso].sum())
        if total <= 0:
            return "Sem dados", np.nan
        vencedor = agrupado.sort_values([coluna_peso, coluna_categoria], ascending=[False, True]).iloc[0]
        return str(vencedor[coluna_categoria]), round(float(vencedor[coluna_peso]) / total * 100, 1)

    agrupado = dados[coluna_categoria].value_counts(dropna=False)
    total = int(agrupado.sum())
    if total <= 0:
        return "Sem dados", np.nan
    vencedor = agrupado.index[0]
    return str(vencedor), round(float(agrupado.iloc[0]) / total * 100, 1)


def calcular_top_ocupacoes(
    df: pd.DataFrame,
    coluna_cbo: str,
    coluna_cbo_nome: str,
    coluna_peso: str | None = None,
    n: int = 3,
) -> str:
    if df.empty or coluna_cbo not in df.columns:
        return "Sem dados"

    dados = df[[coluna_cbo] + ([coluna_cbo_nome] if coluna_cbo_nome in df.columns else []) + ([coluna_peso] if coluna_peso else [])].copy()
    dados[coluna_cbo] = dados[coluna_cbo].fillna("CBO não identificado").astype(str).str.strip()
    if coluna_cbo_nome in dados.columns:
        dados[coluna_cbo_nome] = dados[coluna_cbo_nome].fillna("").astype(str).str.strip()
    else:
        dados[coluna_cbo_nome] = ""

    if coluna_peso and coluna_peso in dados.columns:
        dados[coluna_peso] = pd.to_numeric(dados[coluna_peso], errors="coerce").fillna(0)
        agrupado = (
            dados.groupby([coluna_cbo, coluna_cbo_nome], as_index=False)[coluna_peso]
            .sum()
            .sort_values([coluna_peso, coluna_cbo], ascending=[False, True])
        )
    else:
        agrupado = (
            dados.groupby([coluna_cbo, coluna_cbo_nome], as_index=False)
            .size()
            .rename(columns={"size": "peso"})
            .sort_values(["peso", coluna_cbo], ascending=[False, True])
        )
        coluna_peso = "peso"

    if agrupado.empty:
        return "Sem dados"

    top_n = agrupado.head(n)
    ocupacoes: list[str] = []
    for _, row in top_n.iterrows():
        codigo = str(row[coluna_cbo]).strip()
        nome = str(row[coluna_cbo_nome]).strip()
        if not nome or normalize_text(nome) in {"", "nan", "cbo nao identificado"}:
            ocupacoes.append(codigo)
        else:
            ocupacoes.append(f"{codigo} - {nome}")
    return "; ".join(ocupacoes)


def build_group_dictionary() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"setor_grupo": grupo, "cnae_divisoes": ", ".join(divisoes)}
            for grupo, divisoes in TARGET_CNAES.items()
        ]
    )


def detect_separator(file_path: Path) -> str:
    with file_path.open("r", encoding="latin-1", newline="") as file_handle:
        sample = file_handle.readline()
    return ";" if sample.count(";") > sample.count(",") else ","


def resolve_first_available_column(normalized_columns: dict[str, str], aliases: list[str]) -> str | None:
    for alias in aliases:
        matched = normalized_columns.get(normalize_text(alias))
        if matched is not None:
            return matched
    return None


def localizar_base_tratada(raw_dir: Path) -> Path | None:
    candidates = [
        raw_dir / "rais_tratada_2022_2025.csv",
        raw_dir / "rais_tratada_2022_2025.parquet",
        raw_dir / "rais_tratada_2022_2025.xlsx",
        raw_dir / "rais_tratada_2022_2025.xls",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def load_cnae_dimension(path: Path) -> tuple[pd.DataFrame, dict[str, str]]:
    # A dimensao de CNAE ajuda a converter subclasse em divisao quando a base nao traz a divisao diretamente.
    dimension = pd.read_excel(path, sheet_name="CNAE-SUB")
    normalized_columns = {normalize_text(column): column for column in dimension.columns}
    divisao_col = normalized_columns["cod_div"]
    subclasse_col = normalized_columns["cod_subcsp"]

    division_map = (
        pd.DataFrame(
            {
                "cnae_subclasse": dimension[subclasse_col].apply(lambda value: clean_digits(value, width=7)),
                "cnae_divisao": dimension[divisao_col].apply(lambda value: clean_digits(value, width=2)),
            }
        )
        .dropna(subset=["cnae_subclasse", "cnae_divisao"])
        .drop_duplicates(subset=["cnae_subclasse"])
        .reset_index(drop=True)
    )
    division_name_map = {
        clean_digits(divisao, width=2): str(nome).strip()
        for divisao, nome in zip(
            dimension[divisao_col],
            dimension[next(column for key, column in normalized_columns.items() if key.startswith("divis"))],
        )
        if clean_digits(divisao, width=2) is not None
    }
    return division_map, division_name_map


def load_cbo_dictionary(path: Path) -> dict[str, str]:
    # O dicionario de CBO e usado para deixar o Top 3 de ocupacoes legivel no material de slide.
    if not path.exists():
        LOGGER.warning("Dicionario de CBO nao encontrado em %s. Os top 3 usarao apenas o codigo.", path)
        return {}

    workbook = pd.ExcelFile(path)
    normalized_sheets = {normalize_text(sheet_name): sheet_name for sheet_name in workbook.sheet_names}
    sheet_name = normalized_sheets.get("ocupacao", workbook.sheet_names[0])
    dictionary = pd.read_excel(path, sheet_name=sheet_name)
    value_column = dictionary.columns[0]
    parsed = dictionary[value_column].astype(str).str.split(":", n=1, expand=True)
    cbo_df = pd.DataFrame(
        {
            "cbo_codigo": parsed[0].apply(lambda value: clean_digits(value, width=6)),
            "cbo_descricao": parsed[1].fillna("").astype(str).str.strip(),
        }
    )
    cbo_df = cbo_df.loc[cbo_df["cbo_codigo"].notna()].drop_duplicates(subset=["cbo_codigo"])
    return dict(zip(cbo_df["cbo_codigo"], cbo_df["cbo_descricao"]))


def localizar_arquivo_vinculo_extraido(year: int) -> Path:
    extracted_dir = RAW_DIR / f"{year}_vinc_sul" / "extracted"
    files = [path for path in extracted_dir.iterdir() if path.is_file()] if extracted_dir.exists() else []
    if not files:
        raise FileNotFoundError(
            f"Nenhum arquivo de vinculos extraido foi encontrado para {year} em {extracted_dir}."
        )
    return files[0]


def resolve_columns_raw(file_path: Path, separator: str) -> dict[str, str]:
    header = pd.read_csv(file_path, encoding="latin-1", sep=separator, nrows=0)
    normalized_columns = {normalize_text(column): column for column in header.columns}

    required = {
        "municipio_codigo": ["municipio - codigo", "municipio"],
        "cnae_subclasse": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
        "cbo_codigo": ["cbo 2002 ocupacao - codigo", "cbo ocupacao 2002", "cbo 2002 ocupacao"],
        "sexo": ["sexo - codigo", "sexo trabalhador"],
        "idade": ["idade"],
        "escolaridade": ["escolaridade apos 2005 - codigo", "escolaridade apos 2005"],
        "remuneracao_media": ["vl rem dezembro nom", "vl remun dezembro nom", "vl rem media nom", "vl remun media nom"],
    }
    optional = {
        "vinculo_ativo_3112": ["ind vinculo ativo 31/12 - codigo", "vinculo ativo 31/12"],
        "ind_vinculo_abandonado": ["ind vinculo abandonado - codigo", "ind vinculo abandonado"],
        "qtd_vinculos": ["qtd vinculos", "qtd_vinculos"],
    }

    resolved: dict[str, str] = {}
    for alias, aliases in required.items():
        matched = resolve_first_available_column(normalized_columns, aliases)
        if matched is None:
            raise ValueError(f"Coluna obrigatoria nao encontrada no microdado bruto: {aliases[0]}")
        resolved[alias] = matched

    for alias, aliases in optional.items():
        matched = resolve_first_available_column(normalized_columns, aliases)
        if matched is not None:
            resolved[alias] = matched

    return resolved


def normalize_active_flag(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    result = numeric.eq(1)
    if result.notna().any():
        return result.fillna(False)
    normalized = series.astype(str).map(normalize_text)
    return normalized.isin({"1", "s", "sim", "ativo", "true"})


def aggregate_standardized_frame(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(
            columns=[
                "ano",
                "setor_grupo",
                "cnae_divisao",
                "sexo",
                "faixa_etaria",
                "escolaridade",
                "cbo_codigo",
                "cbo_descricao",
                "peso_vinculos",
                "remuneracao_soma",
                "remuneracao_peso",
            ]
        )

    grouped = (
        frame.groupby(
            [
                "ano",
                "setor_grupo",
                "cnae_divisao",
                "sexo",
                "faixa_etaria",
                "escolaridade",
                "cbo_codigo",
                "cbo_descricao",
            ],
            as_index=False,
        )[["peso_vinculos", "remuneracao_soma", "remuneracao_peso"]]
        .sum()
    )
    grouped["ano"] = grouped["ano"].astype(int)
    return grouped


def padronizar_chunk_raw(
    chunk: pd.DataFrame,
    year: int,
    resolved_columns: dict[str, str],
    cnae_dimension: pd.DataFrame,
    cbo_dictionary: dict[str, str],
) -> tuple[pd.DataFrame, dict[str, int]]:
    # Nesta etapa os microdados brutos sao convertidos para um formato padrao e ja agregados.
    renamed = chunk.rename(columns={source: target for target, source in resolved_columns.items()})
    stats = {"linhas_lidas": len(renamed), "linhas_filtradas": 0}

    renamed["municipio_codigo"] = pd.to_numeric(renamed["municipio_codigo"], errors="coerce").astype("Int64")
    renamed = renamed.loc[renamed["municipio_codigo"].astype(str).str.startswith("42")].copy()

    if "vinculo_ativo_3112" in renamed.columns:
        active_mask = normalize_active_flag(renamed["vinculo_ativo_3112"])
        renamed = renamed.loc[active_mask].copy()
    if "ind_vinculo_abandonado" in renamed.columns:
        non_abandon_mask = parse_numeric_series(renamed["ind_vinculo_abandonado"]).eq(0)
        renamed = renamed.loc[non_abandon_mask].copy()

    renamed["cnae_subclasse"] = renamed["cnae_subclasse"].apply(lambda value: clean_digits(value, width=7))
    renamed["cbo_codigo"] = renamed["cbo_codigo"].apply(lambda value: clean_digits(value, width=6))
    renamed["idade"] = parse_numeric_series(renamed["idade"])
    renamed["remuneracao_media"] = parse_numeric_series(renamed["remuneracao_media"])

    renamed = renamed.merge(cnae_dimension, on="cnae_subclasse", how="left")
    renamed["setor_grupo"] = renamed["cnae_divisao"].map(mapear_grupo_setorial)
    renamed = renamed.loc[renamed["setor_grupo"].notna()].copy()
    stats["linhas_filtradas"] = len(renamed)

    if renamed.empty:
        return aggregate_standardized_frame(pd.DataFrame()), stats

    if "qtd_vinculos" in renamed.columns:
        peso = parse_numeric_series(renamed["qtd_vinculos"]).fillna(0)
    else:
        peso = pd.Series(1, index=renamed.index, dtype="float64")

    renamed["peso_vinculos"] = peso
    renamed["sexo"] = renamed["sexo"].map(padronizar_sexo)
    renamed["faixa_etaria"] = renamed["idade"].map(criar_faixa_etaria)
    renamed["escolaridade"] = renamed["escolaridade"].map(padronizar_escolaridade)
    renamed["cbo_descricao"] = renamed["cbo_codigo"].map(cbo_dictionary).fillna("CBO não identificado")
    renamed["ano"] = year
    remuneracao_valida = renamed["remuneracao_media"].gt(0)
    renamed["remuneracao_peso"] = np.where(remuneracao_valida, renamed["peso_vinculos"], 0.0)
    renamed["remuneracao_soma"] = np.where(remuneracao_valida, renamed["remuneracao_media"], 0.0) * renamed[
        "peso_vinculos"
    ]

    return aggregate_standardized_frame(
        renamed[
            [
                "ano",
                "setor_grupo",
                "cnae_divisao",
                "sexo",
                "faixa_etaria",
                "escolaridade",
                "cbo_codigo",
                "cbo_descricao",
                "peso_vinculos",
                "remuneracao_soma",
                "remuneracao_peso",
            ]
        ]
    ), stats


def carregar_base_bruta(
    years: list[int],
    cnae_dimension: pd.DataFrame,
    cbo_dictionary: dict[str, str],
    chunk_size: int,
) -> pd.DataFrame:
    # O fallback bruto usa apenas 2022 e 2025, lendo os arquivos de vinculos ja extraidos em chunks.
    frames: list[pd.DataFrame] = []

    for year in years:
        file_path = localizar_arquivo_vinculo_extraido(year)
        separator = detect_separator(file_path)
        resolved_columns = resolve_columns_raw(file_path, separator)

        LOGGER.info("Lendo microdado bruto de %s em %s", year, file_path.name)
        LOGGER.info("Ano %s | coluna de remuneracao usada: %s", year, resolved_columns["remuneracao_media"])
        reader = pd.read_csv(
            file_path,
            encoding="latin-1",
            sep=separator,
            usecols=list(resolved_columns.values()),
            chunksize=chunk_size,
            low_memory=False,
        )

        year_frames: list[pd.DataFrame] = []
        linhas_lidas_total = 0
        linhas_filtradas_total = 0
        for chunk_number, chunk in enumerate(reader, start=1):
            aggregated, stats = padronizar_chunk_raw(
                chunk=chunk,
                year=year,
                resolved_columns=resolved_columns,
                cnae_dimension=cnae_dimension,
                cbo_dictionary=cbo_dictionary,
            )
            linhas_lidas_total += stats["linhas_lidas"]
            linhas_filtradas_total += stats["linhas_filtradas"]
            if not aggregated.empty:
                year_frames.append(aggregated)
            LOGGER.info("Ano %s | chunk %s processado", year, chunk_number)

        LOGGER.info("Ano %s | registros antes dos filtros: %s", year, linhas_lidas_total)
        LOGGER.info("Ano %s | registros apos filtros principais: %s", year, linhas_filtradas_total)
        if year_frames:
            frames.append(aggregate_standardized_frame(pd.concat(year_frames, ignore_index=True)))

    if not frames:
        raise ValueError("Nenhum dado foi carregado a partir dos microdados brutos.")
    return aggregate_standardized_frame(pd.concat(frames, ignore_index=True))


def padronizar_chunk_tratado(
    chunk: pd.DataFrame,
    year_filter: list[int],
    cnae_dimension: pd.DataFrame,
    cbo_dictionary: dict[str, str],
) -> tuple[pd.DataFrame, dict[str, int]]:
    stats = {"linhas_lidas": len(chunk), "linhas_filtradas": 0}

    ano_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["ano"])
    if ano_col is None:
        raise ValueError("A base tratada nao possui uma coluna de ano reconhecida.")
    chunk["__ano"] = parse_numeric_series(chunk[ano_col]).astype("Int64")
    chunk = chunk.loc[chunk["__ano"].isin(year_filter)].copy()

    if chunk.empty:
        return aggregate_standardized_frame(pd.DataFrame()), stats

    sigla_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["sigla_uf"])
    uf_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["uf"])
    cod_municipio_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["cod_municipio"])

    if sigla_col is not None:
        chunk = chunk.loc[chunk[sigla_col].astype(str).str.upper().eq("SC")].copy()
    elif uf_col is not None:
        chunk = chunk.loc[parse_numeric_series(chunk[uf_col]).eq(42)].copy()
    elif cod_municipio_col is not None:
        chunk[cod_municipio_col] = chunk[cod_municipio_col].apply(lambda value: clean_digits(value))
        chunk = chunk.loc[chunk[cod_municipio_col].astype(str).str.startswith("42")].copy()

    active_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["vinculo_ativo_3112"])
    abandonment_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["ind_vinculo_abandonado"])
    if active_col is not None:
        chunk = chunk.loc[normalize_active_flag(chunk[active_col])].copy()
    if abandonment_col is not None:
        chunk = chunk.loc[parse_numeric_series(chunk[abandonment_col]).eq(0)].copy()

    cnae_divisao_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["cnae_divisao"])
    cnae_subclasse_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["cnae_subclasse"])
    if cnae_divisao_col is not None:
        chunk["__cnae_divisao"] = chunk[cnae_divisao_col].apply(lambda value: clean_digits(value, width=2))
    elif cnae_subclasse_col is not None:
        chunk["__cnae_subclasse"] = chunk[cnae_subclasse_col].apply(lambda value: clean_digits(value, width=7))
        chunk = chunk.merge(cnae_dimension, left_on="__cnae_subclasse", right_on="cnae_subclasse", how="left")
        chunk["__cnae_divisao"] = chunk["cnae_divisao"]
    else:
        raise ValueError("A base tratada nao possui coluna de CNAE divisao ou subclasse reconhecida.")

    chunk["setor_grupo"] = chunk["__cnae_divisao"].map(mapear_grupo_setorial)
    chunk = chunk.loc[chunk["setor_grupo"].notna()].copy()

    if chunk.empty:
        return aggregate_standardized_frame(pd.DataFrame()), stats

    qtd_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["qtd_vinculos"])
    if qtd_col is not None:
        chunk["peso_vinculos"] = parse_numeric_series(chunk[qtd_col]).fillna(0)
    else:
        chunk["peso_vinculos"] = 1.0

    sexo_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["sexo"])
    idade_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["idade"])
    faixa_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["faixa_etaria"])
    escolaridade_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["escolaridade"])
    cbo_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["cbo_codigo"])
    cbo_nome_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["cbo_descricao"])
    remuneracao_col = encontrar_coluna(chunk, BASE_ALIAS_MAP["remuneracao_media"])

    chunk["sexo"] = chunk[sexo_col].map(padronizar_sexo) if sexo_col is not None else "Não informado"
    if idade_col is not None:
        chunk["faixa_etaria"] = chunk[idade_col].map(criar_faixa_etaria)
    elif faixa_col is not None:
        chunk["faixa_etaria"] = chunk[faixa_col].map(padronizar_faixa_etaria_existente)
    else:
        chunk["faixa_etaria"] = "Não informado"

    chunk["escolaridade"] = (
        chunk[escolaridade_col].map(padronizar_escolaridade) if escolaridade_col is not None else "Não informado"
    )
    chunk["cbo_codigo"] = chunk[cbo_col].apply(lambda value: clean_digits(value, width=6)) if cbo_col is not None else "CBO não identificado"
    if cbo_nome_col is not None:
        chunk["cbo_descricao"] = chunk[cbo_nome_col].fillna("").astype(str).str.strip()
    else:
        chunk["cbo_descricao"] = chunk["cbo_codigo"].map(cbo_dictionary).fillna("CBO não identificado")

    if remuneracao_col is not None:
        chunk["remuneracao_media"] = parse_numeric_series(chunk[remuneracao_col])
    else:
        chunk["remuneracao_media"] = np.nan

    remuneracao_valida = chunk["remuneracao_media"].gt(0)
    chunk["remuneracao_peso"] = np.where(remuneracao_valida, chunk["peso_vinculos"], 0.0)
    chunk["remuneracao_soma"] = np.where(remuneracao_valida, chunk["remuneracao_media"], 0.0) * chunk[
        "peso_vinculos"
    ]
    chunk["ano"] = chunk["__ano"].astype(int)
    chunk["cnae_divisao"] = chunk["__cnae_divisao"]
    stats["linhas_filtradas"] = len(chunk)

    return aggregate_standardized_frame(
        chunk[
            [
                "ano",
                "setor_grupo",
                "cnae_divisao",
                "sexo",
                "faixa_etaria",
                "escolaridade",
                "cbo_codigo",
                "cbo_descricao",
                "peso_vinculos",
                "remuneracao_soma",
                "remuneracao_peso",
            ]
        ]
    ), stats


def carregar_base_tratada(
    file_path: Path,
    years: list[int],
    cnae_dimension: pd.DataFrame,
    cbo_dictionary: dict[str, str],
    chunk_size: int,
) -> pd.DataFrame:
    LOGGER.info("Usando base tratada em %s", file_path)

    frames: list[pd.DataFrame] = []
    linhas_lidas_total = 0
    linhas_filtradas_total = 0

    if file_path.suffix.lower() == ".csv":
        reader = pd.read_csv(file_path, chunksize=chunk_size, low_memory=False)
        for chunk_number, chunk in enumerate(reader, start=1):
            aggregated, stats = padronizar_chunk_tratado(chunk, years, cnae_dimension, cbo_dictionary)
            linhas_lidas_total += stats["linhas_lidas"]
            linhas_filtradas_total += stats["linhas_filtradas"]
            if not aggregated.empty:
                frames.append(aggregated)
            LOGGER.info("Base tratada | chunk %s processado", chunk_number)
    elif file_path.suffix.lower() == ".parquet":
        chunk = pd.read_parquet(file_path)
        aggregated, stats = padronizar_chunk_tratado(chunk, years, cnae_dimension, cbo_dictionary)
        linhas_lidas_total = stats["linhas_lidas"]
        linhas_filtradas_total = stats["linhas_filtradas"]
        if not aggregated.empty:
            frames.append(aggregated)
    elif file_path.suffix.lower() in {".xlsx", ".xls"}:
        chunk = pd.read_excel(file_path)
        aggregated, stats = padronizar_chunk_tratado(chunk, years, cnae_dimension, cbo_dictionary)
        linhas_lidas_total = stats["linhas_lidas"]
        linhas_filtradas_total = stats["linhas_filtradas"]
        if not aggregated.empty:
            frames.append(aggregated)
    else:
        raise ValueError(f"Formato de base tratada nao suportado: {file_path.suffix}")

    LOGGER.info("Base tratada | registros antes dos filtros: %s", linhas_lidas_total)
    LOGGER.info("Base tratada | registros apos filtros principais: %s", linhas_filtradas_total)

    if not frames:
        raise ValueError("Nenhum dado valido foi obtido da base tratada.")
    return aggregate_standardized_frame(pd.concat(frames, ignore_index=True))


def carregar_base_padronizada(
    project_root: Path,
    years: list[int],
    cnae_dimension_path: Path,
    cbo_dictionary_path: Path,
    chunk_size: int,
) -> pd.DataFrame:
    cnae_dimension, _ = load_cnae_dimension(cnae_dimension_path)
    cbo_dictionary = load_cbo_dictionary(cbo_dictionary_path)

    treated_path = localizar_base_tratada(project_root / "data" / "raw")
    if treated_path is not None:
        return carregar_base_tratada(
            file_path=treated_path,
            years=years,
            cnae_dimension=cnae_dimension,
            cbo_dictionary=cbo_dictionary,
            chunk_size=chunk_size,
        )

    LOGGER.info("Nenhuma base tratada encontrada. O script usara os microdados brutos de vinculos extraidos.")
    return carregar_base_bruta(
        years=years,
        cnae_dimension=cnae_dimension,
        cbo_dictionary=cbo_dictionary,
        chunk_size=chunk_size,
    )


def formatar_cnae_divisoes(setor_grupo: str) -> str:
    return ", ".join(TARGET_CNAES[setor_grupo])


def formatar_cnae_divisao_utilizada(setor_grupo: str) -> str:
    return TARGET_CNAES[setor_grupo][0]


def gerar_matriz_perfil(df: pd.DataFrame) -> pd.DataFrame:
    # A matriz principal resume cada setor por ano com estoque, perfil predominante, remuneracao e top ocupacoes.
    rows: list[dict[str, object]] = []
    anos_da_base = sorted({int(value) for value in df["ano"].dropna().tolist()}) if not df.empty else []

    for ano in anos_da_base:
        for setor_grupo in TARGET_CNAES:
            subset = df.loc[(df["ano"].eq(ano)) & (df["setor_grupo"].eq(setor_grupo))].copy()
            total_vinculos = int(subset["peso_vinculos"].sum()) if not subset.empty else 0

            if subset.empty or total_vinculos == 0:
                rows.append(
                    {
                        "ano": ano,
                        "setor_grupo": setor_grupo,
                        "cnae_divisao_utilizada": formatar_cnae_divisao_utilizada(setor_grupo),
                        "cnae_divisoes": formatar_cnae_divisoes(setor_grupo),
                        "vinculos_formais": 0,
                        "sexo_predominante": "Sem dados",
                        "perc_sexo_predominante": np.nan,
                        "faixa_etaria_predominante": "Sem dados",
                        "perc_faixa_etaria": np.nan,
                        "escolaridade_predominante": "Sem dados",
                        "perc_escolaridade": np.nan,
                        "remuneracao_media": np.nan,
                        "top_ocupacoes": "Sem dados",
                    }
                )
                continue

            sexo_pred, sexo_pct = categoria_predominante(subset, "sexo", "peso_vinculos")
            faixa_pred, faixa_pct = categoria_predominante(subset, "faixa_etaria", "peso_vinculos")
            escola_pred, escola_pct = categoria_predominante(subset, "escolaridade", "peso_vinculos")
            top_ocupacoes = calcular_top_ocupacoes(subset, "cbo_codigo", "cbo_descricao", "peso_vinculos", n=3)

            remuneracao_peso = float(subset["remuneracao_peso"].sum())
            remuneracao_media = np.nan
            if remuneracao_peso > 0:
                remuneracao_media = float(subset["remuneracao_soma"].sum()) / remuneracao_peso

            rows.append(
                {
                    "ano": ano,
                    "setor_grupo": setor_grupo,
                    "cnae_divisao_utilizada": formatar_cnae_divisao_utilizada(setor_grupo),
                    "cnae_divisoes": formatar_cnae_divisoes(setor_grupo),
                    "vinculos_formais": total_vinculos,
                    "sexo_predominante": sexo_pred,
                    "perc_sexo_predominante": round(sexo_pct, 1) if pd.notna(sexo_pct) else np.nan,
                    "faixa_etaria_predominante": faixa_pred,
                    "perc_faixa_etaria": round(faixa_pct, 1) if pd.notna(faixa_pct) else np.nan,
                    "escolaridade_predominante": escola_pred,
                    "perc_escolaridade": round(escola_pct, 1) if pd.notna(escola_pct) else np.nan,
                    "remuneracao_media": round(remuneracao_media, 2) if pd.notna(remuneracao_media) else np.nan,
                    "top_ocupacoes": top_ocupacoes,
                }
            )

    return pd.DataFrame(rows)


def montar_perfil_texto(row: pd.Series, suffix: str) -> str:
    sexo = row[f"sexo_pred_{suffix}"]
    faixa = row[f"faixa_etaria_{suffix}"]
    escolaridade = row[f"escolaridade_{suffix}"]
    if "Sem dados" in {sexo, faixa, escolaridade}:
        return "Sem dados"
    return f"{sexo}, {faixa}, {escolaridade}"


def gerar_matriz_comparativa(df_2022: pd.DataFrame, df_2025: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # O comparativo coloca lado a lado 2022 e 2025 para facilitar leitura executiva e uso em apresentacoes.
    matriz_2022 = gerar_matriz_perfil(df_2022).rename(
        columns={
            "vinculos_formais": "vinculos_2022",
            "sexo_predominante": "sexo_pred_2022",
            "faixa_etaria_predominante": "faixa_etaria_2022",
            "escolaridade_predominante": "escolaridade_2022",
            "remuneracao_media": "remuneracao_2022",
            "top_ocupacoes": "top_ocupacoes_2022",
        }
    )
    matriz_2025 = gerar_matriz_perfil(df_2025).rename(
        columns={
            "vinculos_formais": "vinculos_2025",
            "sexo_predominante": "sexo_pred_2025",
            "faixa_etaria_predominante": "faixa_etaria_2025",
            "escolaridade_predominante": "escolaridade_2025",
            "remuneracao_media": "remuneracao_2025",
            "top_ocupacoes": "top_ocupacoes_2025",
        }
    )

    keep_2022 = [
        "setor_grupo",
        "cnae_divisao_utilizada",
        "cnae_divisoes",
        "vinculos_2022",
        "sexo_pred_2022",
        "faixa_etaria_2022",
        "escolaridade_2022",
        "remuneracao_2022",
        "top_ocupacoes_2022",
    ]
    keep_2025 = [
        "setor_grupo",
        "cnae_divisao_utilizada",
        "cnae_divisoes",
        "vinculos_2025",
        "sexo_pred_2025",
        "faixa_etaria_2025",
        "escolaridade_2025",
        "remuneracao_2025",
        "top_ocupacoes_2025",
    ]

    comparativo = matriz_2022[keep_2022].merge(
        matriz_2025[keep_2025],
        on=["setor_grupo", "cnae_divisao_utilizada", "cnae_divisoes"],
        how="outer",
    )
    comparativo["variacao_abs_vinculos"] = comparativo["vinculos_2025"].fillna(0) - comparativo["vinculos_2022"].fillna(0)
    comparativo["variacao_perc_vinculos"] = np.where(
        comparativo["vinculos_2022"].fillna(0) > 0,
        (comparativo["vinculos_2025"].fillna(0) - comparativo["vinculos_2022"].fillna(0))
        / comparativo["vinculos_2022"].replace(0, np.nan)
        * 100,
        np.nan,
    )
    comparativo["variacao_perc_remuneracao"] = np.where(
        comparativo["remuneracao_2022"].fillna(0) > 0,
        (comparativo["remuneracao_2025"] - comparativo["remuneracao_2022"])
        / comparativo["remuneracao_2022"].replace(0, np.nan)
        * 100,
        np.nan,
    )
    comparativo["variacao_perc_vinculos"] = comparativo["variacao_perc_vinculos"].round(1)
    comparativo["variacao_perc_remuneracao"] = comparativo["variacao_perc_remuneracao"].round(1)
    comparativo["remuneracao_2022"] = comparativo["remuneracao_2022"].round(2)
    comparativo["remuneracao_2025"] = comparativo["remuneracao_2025"].round(2)

    comparativo = comparativo[
        [
            "setor_grupo",
            "cnae_divisao_utilizada",
            "cnae_divisoes",
            "vinculos_2022",
            "vinculos_2025",
            "variacao_abs_vinculos",
            "variacao_perc_vinculos",
            "sexo_pred_2022",
            "sexo_pred_2025",
            "faixa_etaria_2022",
            "faixa_etaria_2025",
            "escolaridade_2022",
            "escolaridade_2025",
            "remuneracao_2022",
            "remuneracao_2025",
            "variacao_perc_remuneracao",
            "top_ocupacoes_2022",
            "top_ocupacoes_2025",
        ]
    ]

    slide_friendly = comparativo[
        [
            "setor_grupo",
            "cnae_divisao_utilizada",
            "vinculos_2022",
            "vinculos_2025",
            "variacao_perc_vinculos",
            "sexo_pred_2022",
            "faixa_etaria_2022",
            "escolaridade_2022",
            "sexo_pred_2025",
            "faixa_etaria_2025",
            "escolaridade_2025",
            "remuneracao_2022",
            "remuneracao_2025",
            "top_ocupacoes_2025",
        ]
    ].copy()
    slide_friendly["perfil_2022"] = slide_friendly.apply(lambda row: montar_perfil_texto(row, "2022"), axis=1)
    slide_friendly["perfil_2025"] = slide_friendly.apply(lambda row: montar_perfil_texto(row, "2025"), axis=1)
    slide_friendly = slide_friendly[
        [
            "setor_grupo",
            "cnae_divisao_utilizada",
            "vinculos_2022",
            "vinculos_2025",
            "variacao_perc_vinculos",
            "perfil_2022",
            "perfil_2025",
            "remuneracao_2022",
            "remuneracao_2025",
            "top_ocupacoes_2025",
        ]
    ]

    return {"comparativo": comparativo, "slide_friendly": slide_friendly}


def auto_fit_columns(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for index, column in enumerate(dataframe.columns):
        max_length = max(
            len(str(column)),
            dataframe[column].apply(lambda value: len(str(value))).max() if not dataframe.empty else 0,
        )
        worksheet.column_dimensions[get_column_letter(index + 1)].width = min(max_length + 2, 42)


def exportar_resultados(matriz: pd.DataFrame, comparativo: dict[str, pd.DataFrame]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_file = OUTPUT_DIR / "matriz_perfil_setorial_rais_2022_2025.xlsx"
    matriz_2022 = matriz.loc[matriz["ano"].eq(2022)].reset_index(drop=True)
    matriz_2025 = matriz.loc[matriz["ano"].eq(2025)].reset_index(drop=True)
    comparativo_df = comparativo["comparativo"]
    slide_friendly_df = comparativo["slide_friendly"]
    dicionario_grupos = build_group_dictionary()

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        matriz.to_excel(writer, sheet_name="matriz_completa", index=False)
        matriz_2022.to_excel(writer, sheet_name="matriz_2022", index=False)
        matriz_2025.to_excel(writer, sheet_name="matriz_2025", index=False)
        comparativo_df.to_excel(writer, sheet_name="comparativo_2022_2025", index=False)
        slide_friendly_df.to_excel(writer, sheet_name="slide_friendly", index=False)
        dicionario_grupos.to_excel(writer, sheet_name="dicionario_grupos", index=False)

        for sheet_name, dataframe in {
            "matriz_completa": matriz,
            "matriz_2022": matriz_2022,
            "matriz_2025": matriz_2025,
            "comparativo_2022_2025": comparativo_df,
            "slide_friendly": slide_friendly_df,
            "dicionario_grupos": dicionario_grupos,
        }.items():
            auto_fit_columns(writer, sheet_name, dataframe)

    matriz_2022.to_csv(OUTPUT_DIR / "matriz_2022.csv", index=False, encoding="utf-8-sig")
    matriz_2025.to_csv(OUTPUT_DIR / "matriz_2025.csv", index=False, encoding="utf-8-sig")
    comparativo_df.to_csv(OUTPUT_DIR / "comparativo_2022_2025.csv", index=False, encoding="utf-8-sig")
    slide_friendly_df.to_csv(OUTPUT_DIR / "slide_friendly.csv", index=False, encoding="utf-8-sig")

    LOGGER.info("Arquivo Excel gerado em %s", output_file)


def validar_base(df: pd.DataFrame) -> None:
    anos_disponiveis = sorted(df["ano"].dropna().astype(int).unique().tolist()) if not df.empty else []
    LOGGER.info("Anos encontrados na base padronizada: %s", anos_disponiveis)

    missing_years = [year for year in TARGET_YEARS if year not in anos_disponiveis]
    if missing_years:
        raise ValueError(f"Os anos esperados nao foram encontrados na base: {missing_years}")

    cnaes_encontrados = set(df["cnae_divisao"].dropna().astype(str).unique().tolist())
    cnaes_esperados = {cnae for divisoes in TARGET_CNAES.values() for cnae in divisoes}
    cnaes_ausentes = sorted(cnaes_esperados - cnaes_encontrados)
    if cnaes_ausentes:
        LOGGER.warning("As seguintes divisões CNAE esperadas nao apareceram na base filtrada: %s", cnaes_ausentes)
    else:
        LOGGER.info("Todos os CNAEs esperados apareceram na base filtrada.")

    total_peso = float(df["peso_vinculos"].sum())
    peso_remunerado = float(df["remuneracao_peso"].sum())
    if total_peso > 0:
        share_sem_remuneracao = (total_peso - peso_remunerado) / total_peso * 100
        LOGGER.info("Percentual de vinculos sem remuneracao valida: %.1f%%", share_sem_remuneracao)
        if share_sem_remuneracao > 25:
            LOGGER.warning("A remuneracao media possui muitos valores nulos ou invalidos: %.1f%%", share_sem_remuneracao)


def validar_matriz(matriz: pd.DataFrame) -> None:
    sem_dados = matriz.loc[matriz["vinculos_formais"].fillna(0).eq(0), ["ano", "setor_grupo"]]
    for _, row in sem_dados.iterrows():
        LOGGER.warning("Grupo setorial sem dados: %s em %s", row["setor_grupo"], int(row["ano"]))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera matriz-resumo do perfil setorial da RAIS para 2022 e 2025 em Santa Catarina."
    )
    parser.add_argument(
        "--project-root",
        default=str(PROJECT_ROOT),
        help="Diretorio raiz do projeto.",
    )
    parser.add_argument(
        "--cnae-dimension-path",
        default=str(CNAE_DIMENSION_FILE),
        help="Caminho do arquivo cnae_dimensao.xlsx.",
    )
    parser.add_argument(
        "--cbo-dictionary-path",
        default=str(CBO_DICTIONARY_FILE),
        help="Caminho do dicionario de CBO.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=500_000,
        help="Quantidade de linhas por chunk ao ler CSVs grandes.",
    )
    return parser


def main() -> None:
    setup_logging()
    args = build_parser().parse_args()

    project_root = Path(args.project_root).resolve()
    cnae_dimension_path = Path(args.cnae_dimension_path).resolve()
    cbo_dictionary_path = Path(args.cbo_dictionary_path).resolve()

    global PROJECT_ROOT, RAW_DIR, OUTPUT_DIR
    PROJECT_ROOT = project_root
    RAW_DIR = PROJECT_ROOT / "data" / "raw"
    OUTPUT_DIR = PROJECT_ROOT / "data" / "output"

    base_padronizada = carregar_base_padronizada(
        project_root=PROJECT_ROOT,
        years=TARGET_YEARS,
        cnae_dimension_path=cnae_dimension_path,
        cbo_dictionary_path=cbo_dictionary_path,
        chunk_size=args.chunk_size,
    )
    validar_base(base_padronizada)

    matriz = gerar_matriz_perfil(base_padronizada)
    validar_matriz(matriz)

    df_2022 = base_padronizada.loc[base_padronizada["ano"].eq(2022)].copy()
    df_2025 = base_padronizada.loc[base_padronizada["ano"].eq(2025)].copy()
    comparativo = gerar_matriz_comparativa(df_2022, df_2025)

    exportar_resultados(matriz, comparativo)


if __name__ == "__main__":
    main()
