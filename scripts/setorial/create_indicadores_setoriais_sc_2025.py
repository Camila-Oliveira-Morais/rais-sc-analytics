from __future__ import annotations

import logging
import sys
import unicodedata
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
YEARS = [2022, 2023, 2024, 2025]
CNAE = ROOT / "data/cnae_dimensao.xlsx"
MUNICIPIOS = ROOT / "data/reference/municipios_sc_mesorregioes.csv"
CBO = ROOT / "data/dict/dicionario_cbo.xlsx"
SAIDA = ROOT / "data/output/indicadores_setoriais_sc_cnae_2022_2025.xlsx"
CHUNK_SIZE = 500_000


def norm(value: object) -> str:
    return unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode().lower().strip()


def digits(value: object, width: int) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).split(".", 1)[0]
    result = "".join(c for c in text if c.isdigit())
    return result.zfill(width) if result else None


def load_dimensions() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cnae = pd.read_excel(CNAE, sheet_name="CNAE-SUB")
    nc = {norm(c): c for c in cnae.columns}
    cnae = cnae.rename(columns={
        nc["cod_subcsp"]: "cnae_subclasse_codigo",
        nc["cod_div"]: "cnae_divisao_codigo",
        next(v for k, v in nc.items() if k.startswith("divis")): "cnae_divisao_nome",
    })
    cnae["cnae_subclasse_codigo"] = cnae["cnae_subclasse_codigo"].map(lambda x: digits(x, 7))
    cnae["cnae_divisao_codigo"] = cnae["cnae_divisao_codigo"].map(lambda x: digits(x, 2))
    cnae = cnae[["cnae_subclasse_codigo", "cnae_divisao_codigo", "cnae_divisao_nome"]].drop_duplicates("cnae_subclasse_codigo")

    municipios = pd.read_csv(MUNICIPIOS)
    municipios["municipio_codigo"] = pd.to_numeric(municipios["municipio_codigo"], errors="coerce").astype("Int64")
    municipios = municipios[["municipio_codigo", "municipio_nome", "mesorregiao_nome"]].drop_duplicates("municipio_codigo")

    book = pd.ExcelFile(CBO)
    sheet = next((s for s in book.sheet_names if norm(s) == "ocupacao"), book.sheet_names[0])
    raw = pd.read_excel(CBO, sheet_name=sheet).iloc[:, 0].astype(str).str.split(":", n=1, expand=True)
    cbo = pd.DataFrame({
        "cbo_codigo": raw[0].map(lambda x: digits(x, 6)),
        "cbo_nome": raw[1].fillna("").str.strip(),
    }).dropna(subset=["cbo_codigo"]).drop_duplicates("cbo_codigo")
    return cnae, municipios, cbo


def faixa(divisao: pd.Series) -> pd.Series:
    number = pd.to_numeric(divisao, errors="coerce")
    result = pd.Series(pd.NA, index=divisao.index, dtype="string")
    result.loc[number.between(1, 3)] = "CNAE 01 a 03"
    result.loc[number.between(5, 43)] = "CNAE 05 a 43"
    result.loc[number.between(45, 47)] = "CNAE 45 a 47"
    result.loc[number.between(49, 99)] = "CNAE 49 a 99"
    return result


def vinculos_path(year: int) -> Path:
    files = list((ROOT / f"data/raw/{year}_vinc_sul/extracted").glob("*"))
    if not files:
        raise FileNotFoundError(f"Base de vínculos não encontrada para {year}")
    return files[0]


def aggregate_vinculos(year: int, cnae: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, int | str]]:
    path = vinculos_path(year)
    with path.open("r", encoding="latin-1") as source:
        first_line = source.readline()
    separator = ";" if first_line.count(";") > first_line.count(",") else ","
    header = pd.read_csv(path, encoding="latin-1", sep=separator, nrows=0)
    nh = {norm(c): c for c in header.columns}
    candidates = {
        "municipio_codigo": ["municipio - codigo", "municipio"],
        "cbo_codigo": ["cbo 2002 ocupacao - codigo", "cbo ocupacao 2002"],
        "cnae_subclasse_codigo": ["cnae 2.0 subclasse - codigo", "cnae 2.0 subclasse"],
        "ativo": ["ind vinculo ativo 31/12 - codigo", "vinculo ativo 31/12"],
        "abandonado": ["ind vinculo abandonado - codigo", "ind vinculo abandonado"],
        "remuneracao": ["vl rem media nom", "vl remun media nom"],
    }
    selected = {}
    for alias, options in candidates.items():
        match = next((nh[name] for name in options if name in nh), None)
        if match is None and alias != "abandonado":
            raise ValueError(f"Ano {year}: campo obrigatório ausente: {options[0]}")
        if match is not None:
            selected[alias] = match
    has_abandonment = "abandonado" in selected
    detail_parts, total_parts = [], []
    stats: dict[str, int | str] = {"ano": year, "filtro_abandonado_aplicado": "sim" if has_abandonment else "nao_campo_ausente", "linhas_lidas": 0, "vinculos_filtrados": 0, "remuneracoes_positivas": 0}
    reader = pd.read_csv(path, encoding="latin-1", sep=separator, usecols=list(selected.values()), chunksize=CHUNK_SIZE, low_memory=False)
    for number, chunk in enumerate(reader, 1):
        stats["linhas_lidas"] += len(chunk)
        chunk = chunk.rename(columns={v: k for k, v in selected.items()})
        chunk["municipio_codigo"] = pd.to_numeric(chunk["municipio_codigo"], errors="coerce").astype("Int64")
        chunk["ativo"] = pd.to_numeric(chunk["ativo"], errors="coerce")
        if has_abandonment:
            chunk["abandonado"] = pd.to_numeric(chunk["abandonado"], errors="coerce")
        remuneration_text = chunk["remuneracao"].astype(str).str.strip()
        if year == 2022:
            remuneration_text = remuneration_text.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        chunk["remuneracao"] = pd.to_numeric(remuneration_text, errors="coerce")
        mask = chunk["municipio_codigo"].astype(str).str.startswith("42") & chunk["ativo"].eq(1)
        if has_abandonment:
            mask &= chunk["abandonado"].eq(0)
        chunk = chunk.loc[mask].copy()
        chunk["cnae_subclasse_codigo"] = chunk["cnae_subclasse_codigo"].map(lambda x: digits(x, 7))
        chunk["cbo_codigo"] = chunk["cbo_codigo"].map(lambda x: digits(x, 6))
        chunk = chunk.merge(cnae, on="cnae_subclasse_codigo", how="inner")
        chunk["faixa_cnae"] = faixa(chunk["cnae_divisao_codigo"])
        chunk = chunk.loc[chunk["faixa_cnae"].notna()].copy()
        chunk["qtd_vinculos_ativos_3112"] = 1
        chunk["remuneracao_soma"] = chunk["remuneracao"].where(chunk["remuneracao"].gt(0), 0)
        chunk["qtd_remuneracoes_positivas"] = chunk["remuneracao"].gt(0).astype("int64")
        stats["vinculos_filtrados"] += len(chunk)
        stats["remuneracoes_positivas"] += int(chunk["qtd_remuneracoes_positivas"].sum())
        metrics = ["qtd_vinculos_ativos_3112", "remuneracao_soma", "qtd_remuneracoes_positivas"]
        d = chunk.loc[chunk["cnae_divisao_codigo"].between("41", "43")]
        if not d.empty:
            detail_parts.append(d.groupby(["municipio_codigo", "cnae_divisao_codigo", "cnae_divisao_nome", "cbo_codigo"], dropna=False, as_index=False)[metrics].sum())
        total_parts.append(chunk.groupby(["municipio_codigo", "faixa_cnae"], as_index=False)[metrics].sum())
        logging.info("Ano %s | chunk %s processado", year, number)

    def combine(parts: list[pd.DataFrame], keys: list[str]) -> pd.DataFrame:
        return pd.concat(parts, ignore_index=True).groupby(keys, dropna=False, as_index=False)[["qtd_vinculos_ativos_3112", "remuneracao_soma", "qtd_remuneracoes_positivas"]].sum()
    detail = combine(detail_parts, ["municipio_codigo", "cnae_divisao_codigo", "cnae_divisao_nome", "cbo_codigo"])
    total = combine(total_parts, ["municipio_codigo", "faixa_cnae"])
    detail["ano_referencia"] = year
    total["ano_referencia"] = year
    return detail, total, stats


def load_estabelecimentos() -> tuple[pd.DataFrame, pd.DataFrame]:
    historic = pd.read_excel(ROOT / "data/output/rais_estabelecimentos_sc_2022_2024.xlsx", sheet_name="Municipio_Divisao")
    latest = pd.read_excel(ROOT / "data/output/rais_estabelecimentos_sc_2025.xlsx", sheet_name="Municipio_Divisao")
    estab = pd.concat([historic.loc[historic["ano_referencia"].between(2022, 2024)], latest], ignore_index=True)
    estab["cnae_divisao_codigo"] = estab["cnae_divisao_codigo"].map(lambda x: digits(x, 2))
    # A planilha consolidada ja exclui declaracoes RAIS negativas.
    detail = estab.loc[estab["cnae_divisao_codigo"].between("41", "43"), ["ano_referencia", "municipio_codigo", "cnae_divisao_codigo", "estabelecimentos"]]
    detail = detail.rename(columns={"estabelecimentos": "qtd_estabelecimentos_municipio_divisao"})
    estab["faixa_cnae"] = faixa(estab["cnae_divisao_codigo"])
    total = estab.loc[estab["faixa_cnae"].notna()].groupby(["ano_referencia", "municipio_codigo", "faixa_cnae"], as_index=False)["estabelecimentos"].sum()
    return detail, total.rename(columns={"estabelecimentos": "qtd_estabelecimentos"})


def format_sheet(writer: pd.ExcelWriter, sheet: str, df: pd.DataFrame) -> None:
    ws, wb = writer.sheets[sheet], writer.book
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
    money = wb.add_format({"num_format": "R$ #,##0.00"})
    integer = wb.add_format({"num_format": "#,##0"})
    for i, col in enumerate(df.columns):
        width = min(max(len(col), int(df[col].astype(str).str.len().max()) if len(df) else 0) + 2, 45)
        fmt = money if col == "remuneracao_media_ano" else integer if col.startswith("qtd_") else None
        ws.set_column(i, i, width, fmt)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s", datefmt="%H:%M:%S")
    cnae, municipios, cbo = load_dimensions()
    detail_parts, total_parts, stats_rows = [], [], []
    for year in YEARS:
        year_detail, year_total, year_stats = aggregate_vinculos(year, cnae)
        detail_parts.append(year_detail)
        total_parts.append(year_total)
        stats_rows.append(year_stats)
    detail = pd.concat(detail_parts, ignore_index=True)
    total = pd.concat(total_parts, ignore_index=True)
    estab_detail, estab_total = load_estabelecimentos()
    detail = detail.merge(municipios, on="municipio_codigo", how="left").merge(cbo, on="cbo_codigo", how="left").merge(estab_detail, on=["ano_referencia", "municipio_codigo", "cnae_divisao_codigo"], how="left")
    total = total.merge(municipios, on="municipio_codigo", how="left").merge(estab_total, on=["ano_referencia", "municipio_codigo", "faixa_cnae"], how="left")
    detail["cbo_nome"] = detail["cbo_nome"].fillna("CBO não identificado no dicionário")
    for df in (detail, total):
        df["remuneracao_media_ano"] = df["remuneracao_soma"] / df["qtd_remuneracoes_positivas"]
    detail = detail[["ano_referencia", "municipio_codigo", "municipio_nome", "mesorregiao_nome", "cnae_divisao_codigo", "cnae_divisao_nome", "cbo_codigo", "cbo_nome", "qtd_estabelecimentos_municipio_divisao", "qtd_vinculos_ativos_3112", "qtd_remuneracoes_positivas", "remuneracao_media_ano"]].sort_values(["municipio_nome", "cnae_divisao_codigo", "cbo_codigo"])
    total = total[["ano_referencia", "municipio_codigo", "municipio_nome", "mesorregiao_nome", "faixa_cnae", "qtd_estabelecimentos", "qtd_vinculos_ativos_3112", "qtd_remuneracoes_positivas", "remuneracao_media_ano"]].sort_values(["municipio_nome", "faixa_cnae"])
    metadata = pd.DataFrame([
        ("anos_referencia", "2022 a 2025"), ("territorio", "Santa Catarina, por municipio"),
        ("filtro_vinculo_abandonado", "0 (não) em 2023-2025; campo inexistente no layout 2022"), ("filtro_vinculo_ativo_31_12", "1 (sim), todos os anos"),
        ("regra_remuneracao", "Media de Vl Rem Media Nom apenas entre valores > 0; zeros e nulos nao entram no denominador."),
        ("regra_vinculos", "A quantidade de vinculos inclui todos os ativos em 31/12 e nao abandonados, mesmo quando remuneracao e zero/nula."),
        ("regra_estabelecimentos", "Base RAIS Estabelecimento, declaracoes nao negativas. Na aba detalhada a metrica esta no nivel municipio-divisao e se repete por CBO; nao somar entre CBOs."),
        ("fonte_vinculos", "data/raw/{ano}_vinc_sul/extracted"),
        ("fonte_estabelecimentos", "rais_estabelecimentos_sc_2022_2024.xlsx e rais_estabelecimentos_sc_2025.xlsx"),
        *[(f"{k}_{row['ano']}", v) for row in stats_rows for k, v in row.items() if k != "ano"],
    ], columns=["chave", "valor"])
    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(SAIDA, engine="xlsxwriter") as writer:
        detail.to_excel(writer, sheet_name="CNAE_41_43_CBO_Municipio", index=False)
        total.to_excel(writer, sheet_name="Faixas_CNAE_Municipio", index=False)
        metadata.to_excel(writer, sheet_name="Metodologia", index=False)
        for sheet, df in [("CNAE_41_43_CBO_Municipio", detail), ("Faixas_CNAE_Municipio", total), ("Metodologia", metadata)]:
            format_sheet(writer, sheet, df)
    logging.info("Gerado: %s", SAIDA)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("Falha na geracao")
        sys.exit(1)
