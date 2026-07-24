from __future__ import annotations

import argparse
import gzip
import json
import subprocess
import unicodedata
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

try:
    import cairosvg
except Exception:
    cairosvg = None


PROJECT_ROOT = Path(__file__).resolve().parents[2]
INPUT_WORKBOOK = PROJECT_ROOT / "data" / "output" / "rais_estabelecimentos_sc_2025.xlsx"
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
REFERENCE_DIR = PROJECT_ROOT / "data" / "reference"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
MICROREGIONS_JSON_PATH = REFERENCE_DIR / "ibge_sc_microrregioes.json"
MUNICIPALITIES_JSON_PATH = REFERENCE_DIR / "ibge_sc_municipios.json"
MICROREGIONS_SVG_PATH = REFERENCE_DIR / "ibge_sc_microrregioes.svg"

IBGE_MICROREGIONS_URL = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/42/microrregioes"
IBGE_MUNICIPALITIES_URL = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/42/municipios"
IBGE_MICROREGIONS_SVG_URL = (
    "https://servicodados.ibge.gov.br/api/v3/malhas/estados/42?intrarregiao=microrregiao&formato=svg&qualidade=minima"
)

SECTOR_COLORS = {
    "AGROPECUÁRIA": "#52B86A",
    "INDÚSTRIA": "#4A63F0",
    "SERVIÇOS": "#F5A623",
}


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


def ensure_reference_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def download_if_missing(url: str, destination: Path) -> Path:
    if destination.exists():
        return destination

    ensure_reference_dir(destination.parent)
    request = urllib.request.Request(url, headers={"Accept-Encoding": "gzip"})
    with urllib.request.urlopen(request, timeout=120) as response:
        payload = response.read()

    destination.write_bytes(payload)
    return destination


def read_json_payload(path: Path) -> list[dict]:
    payload = path.read_bytes()
    if payload[:2] == b"\x1f\x8b":
        payload = gzip.decompress(payload)
    return json.loads(payload.decode("utf-8"))


def load_ibge_microregion_lookup() -> tuple[pd.DataFrame, pd.DataFrame, str]:
    microregions_path = download_if_missing(IBGE_MICROREGIONS_URL, MICROREGIONS_JSON_PATH)
    municipalities_path = download_if_missing(IBGE_MUNICIPALITIES_URL, MUNICIPALITIES_JSON_PATH)
    svg_path = download_if_missing(IBGE_MICROREGIONS_SVG_URL, MICROREGIONS_SVG_PATH)

    microregions_raw = read_json_payload(microregions_path)
    municipalities_raw = read_json_payload(municipalities_path)

    microregions = pd.DataFrame(
        {
            "microrregiao_id": [int(item["id"]) for item in microregions_raw],
            "microrregiao_nome": [str(item["nome"]).strip() for item in microregions_raw],
        }
    )

    municipalities = pd.DataFrame(
        {
            "municipio_codigo": [int(item["id"]) // 10 for item in municipalities_raw],
            "microrregiao_id": [int(item["microrregiao"]["id"]) for item in municipalities_raw],
            "microrregiao_nome": [str(item["microrregiao"]["nome"]).strip() for item in municipalities_raw],
        }
    ).drop_duplicates(subset=["municipio_codigo"])

    return microregions, municipalities, svg_path.read_text(encoding="utf-8")


def load_gr_setor_dimension(path: Path) -> pd.DataFrame:
    dimension = pd.read_excel(path, sheet_name="CNAE-SUB")
    normalized_columns = {normalize_text(column): column for column in dimension.columns}
    division_column = normalized_columns["cod_div"]
    sector_column = normalized_columns["gr_setor"]

    prepared = pd.DataFrame(
        {
            "cnae_divisao_codigo": dimension[division_column].apply(lambda value: clean_digits(value, width=2)),
            "gr_setor": dimension[sector_column].fillna("Nao classificado").astype(str).str.strip(),
        }
    )
    return prepared.dropna(subset=["cnae_divisao_codigo"]).drop_duplicates(subset=["cnae_divisao_codigo"])


def load_rais_municipio_divisao(path: Path) -> pd.DataFrame:
    dataframe = pd.read_excel(path, sheet_name="Municipio_Divisao")
    dataframe["municipio_codigo"] = pd.to_numeric(dataframe["municipio_codigo"], errors="coerce").astype("Int64")
    dataframe["cnae_divisao_codigo"] = dataframe["cnae_divisao_codigo"].apply(lambda value: clean_digits(value, width=2))
    dataframe["qtd_vinculos_ativos"] = pd.to_numeric(dataframe["qtd_vinculos_ativos"], errors="coerce").fillna(0)
    return dataframe.dropna(subset=["municipio_codigo", "cnae_divisao_codigo"]).copy()


def build_microregion_sector_stock(
    rais: pd.DataFrame,
    gr_setor_dimension: pd.DataFrame,
    municipality_lookup: pd.DataFrame,
    microregions: pd.DataFrame,
) -> pd.DataFrame:
    enriched = rais.merge(gr_setor_dimension, on="cnae_divisao_codigo", how="left")
    enriched = enriched.merge(municipality_lookup, on="municipio_codigo", how="left")
    enriched["gr_setor"] = enriched["gr_setor"].fillna("Nao classificado")
    enriched = enriched.loc[enriched["microrregiao_id"].notna()].copy()
    enriched["microrregiao_id"] = enriched["microrregiao_id"].astype(int)

    aggregated = (
        enriched.groupby(
            ["microrregiao_id", "microrregiao_nome", "gr_setor"],
            as_index=False,
        )["qtd_vinculos_ativos"]
        .sum()
        .sort_values(["microrregiao_id", "qtd_vinculos_ativos", "gr_setor"], ascending=[True, False, True])
        .reset_index(drop=True)
    )

    all_combinations = microregions.assign(key=1).merge(
        pd.DataFrame({"gr_setor": sorted(SECTOR_COLORS)}).assign(key=1),
        on="key",
        how="inner",
    ).drop(columns=["key"])

    complete = all_combinations.merge(
        aggregated,
        on=["microrregiao_id", "microrregiao_nome", "gr_setor"],
        how="left",
    )
    complete["qtd_vinculos_ativos"] = complete["qtd_vinculos_ativos"].fillna(0).astype(int)
    return complete


def build_predominance_table(stock: pd.DataFrame) -> pd.DataFrame:
    totals = stock.groupby("microrregiao_id", as_index=False)["qtd_vinculos_ativos"].sum().rename(
        columns={"qtd_vinculos_ativos": "estoque_total_microrregiao"}
    )
    ranked = stock.sort_values(
        ["microrregiao_id", "qtd_vinculos_ativos", "gr_setor"],
        ascending=[True, False, True],
    )
    predominance = ranked.drop_duplicates(subset=["microrregiao_id"], keep="first").rename(
        columns={
            "gr_setor": "gr_setor_predominante",
            "qtd_vinculos_ativos": "estoque_predominante",
        }
    )
    predominance = predominance.merge(totals, on="microrregiao_id", how="left")
    predominance["participacao_predominante"] = predominance["estoque_predominante"] / predominance[
        "estoque_total_microrregiao"
    ]
    predominance["cor_predominante"] = predominance["gr_setor_predominante"].map(SECTOR_COLORS).fillna("#B8C0D4")
    return predominance.sort_values("estoque_total_microrregiao", ascending=False).reset_index(drop=True)


def hex_to_rgb(color: str) -> tuple[int, int, int]:
    return tuple(int(color[index : index + 2], 16) for index in (1, 3, 5))


def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def mix_with_white(color: str, ratio: float) -> str:
    ratio = max(0.0, min(1.0, ratio))
    base = hex_to_rgb(color)
    mixed = tuple(round(255 - (255 - channel) * ratio) for channel in base)
    return rgb_to_hex(mixed)


def build_map_svg(predominance: pd.DataFrame, source_svg: str) -> str:
    namespace = {"svg": "http://www.w3.org/2000/svg"}
    ET.register_namespace("", namespace["svg"])
    source_root = ET.fromstring(source_svg)

    width = 3000
    height = 2200
    map_group_translate_x = 120
    map_group_translate_y = 260
    map_scale = 1.55

    root = ET.Element(
        "{http://www.w3.org/2000/svg}svg",
        attrib={
            "version": "1.1",
            "width": str(width),
            "height": str(height),
            "viewBox": f"0 0 {width} {height}",
        },
    )

    title = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": "120",
            "y": "92",
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "44",
            "font-weight": "700",
            "fill": "#16244E",
        },
    )
    title.text = "Predominancia do Estoque de Empregos por GR_SETOR"

    subtitle = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": "120",
            "y": "136",
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "22",
            "fill": "#47557E",
        },
    )
    subtitle.text = "RAIS 2025 | Santa Catarina por microrregiao | Variavel: qtd_vinculos_ativos"

    note = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": "120",
            "y": "170",
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "18",
            "fill": "#5E6B8E",
        },
    )
    note.text = "Observacao: no arquivo cnae_dimensao, GR_SETOR possui 3 grupos: AGROPECUARIA, INDUSTRIA e SERVICOS."

    map_group = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}g",
        attrib={"transform": f"translate({map_group_translate_x},{map_group_translate_y}) scale({map_scale})"},
    )

    lookup = predominance.set_index("microrregiao_id").to_dict(orient="index")
    source_group = source_root.find(".//svg:g", namespace)
    if source_group is None:
        raise ValueError("Grupo principal da malha nao encontrado no SVG do IBGE.")

    for path in source_group.findall("svg:path", namespace):
        microregion_id = int(path.attrib["id"])
        row = lookup.get(microregion_id)
        fill = "#D7DCE8"
        if row is not None:
            fill = mix_with_white(row["cor_predominante"], 0.45 + 0.55 * float(row["participacao_predominante"]))
        ET.SubElement(
            map_group,
            "{http://www.w3.org/2000/svg}path",
            attrib={
                "d": path.attrib["d"],
                "fill": fill,
                "stroke": "#FFFFFF",
                "stroke-width": "180",
                "vector-effect": "non-scaling-stroke",
            },
        )

    legend_x = 1920
    legend_y = 320
    legend_title = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": str(legend_x),
            "y": str(legend_y),
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "28",
            "font-weight": "700",
            "fill": "#16244E",
        },
    )
    legend_title.text = "Legenda"

    for index, sector in enumerate(["AGROPECUÁRIA", "INDÚSTRIA", "SERVIÇOS"]):
        y = legend_y + 54 + index * 58
        ET.SubElement(
            root,
            "{http://www.w3.org/2000/svg}rect",
            attrib={
                "x": str(legend_x),
                "y": str(y - 22),
                "width": "36",
                "height": "36",
                "rx": "8",
                "fill": SECTOR_COLORS[sector],
            },
        )
        text = ET.SubElement(
            root,
            "{http://www.w3.org/2000/svg}text",
            attrib={
                "x": str(legend_x + 52),
                "y": str(y + 5),
                "font-family": "Segoe UI, Arial, sans-serif",
                "font-size": "24",
                "fill": "#22325D",
            },
        )
        text.text = sector

    intensity_title = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": str(legend_x),
            "y": str(legend_y + 255),
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "24",
            "font-weight": "700",
            "fill": "#16244E",
        },
    )
    intensity_title.text = "Intensidade da cor"

    intensity_text = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": str(legend_x),
            "y": str(legend_y + 288),
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "18",
            "fill": "#5E6B8E",
        },
    )
    intensity_text.text = "Maior saturacao = maior participacao do setor predominante na microrregiao."

    ranking_title = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": str(legend_x),
            "y": str(legend_y + 365),
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "26",
            "font-weight": "700",
            "fill": "#16244E",
        },
    )
    ranking_title.text = "Maiores estoques por microrregiao"

    for index, (_, row) in enumerate(predominance.head(8).iterrows(), start=1):
        y = legend_y + 405 + index * 46
        line = ET.SubElement(
            root,
            "{http://www.w3.org/2000/svg}text",
            attrib={
                "x": str(legend_x),
                "y": str(y),
                "font-family": "Segoe UI, Arial, sans-serif",
                "font-size": "19",
                "fill": "#22325D",
            },
        )
        participation = f"{row['participacao_predominante'] * 100:.1f}%".replace(".", ",")
        total_jobs = f"{int(row['estoque_total_microrregiao']):,}".replace(",", ".")
        line.text = (
            f"{index}. {row['microrregiao_nome']} | {row['gr_setor_predominante']} | "
            f"{total_jobs} empregos | {participation} do total"
        )

    footer = ET.SubElement(
        root,
        "{http://www.w3.org/2000/svg}text",
        attrib={
            "x": "120",
            "y": str(height - 56),
            "font-family": "Segoe UI, Arial, sans-serif",
            "font-size": "18",
            "fill": "#6D7898",
        },
    )
    footer.text = "Fonte: RAIS 2025 (qtd_vinculos_ativos) e malha oficial do IBGE."

    return ET.tostring(root, encoding="unicode")


def export_support_workbook(stock: pd.DataFrame, predominance: pd.DataFrame, output_path: Path) -> Path:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        stock.to_excel(writer, sheet_name="Estoque_Microregiao_GR", index=False)
        predominance.to_excel(writer, sheet_name="Predominancia", index=False)

        workbook = writer.book
        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#243B7D",
            }
        )
        integer_format = workbook.add_format({"num_format": "#,##0"})
        percent_format = workbook.add_format({"num_format": "0.0%"})

        for sheet_name, dataframe in {"Estoque_Microregiao_GR": stock, "Predominancia": predominance}.items():
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, max(len(dataframe), 1), len(dataframe.columns) - 1)
            for column_index, column_name in enumerate(dataframe.columns):
                worksheet.write(0, column_index, column_name, header_format)
                width = 22
                if "nome" in column_name or "gr_setor" in column_name:
                    width = 28
                fmt = None
                if column_name.startswith("estoque") or column_name.endswith("id") or column_name == "qtd_vinculos_ativos":
                    fmt = integer_format
                if column_name == "participacao_predominante":
                    fmt = percent_format
                worksheet.set_column(column_index, column_index, width, fmt)
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera mapa de predominancia do estoque de empregos por GR_SETOR nas microrregioes de SC."
    )
    parser.add_argument("--input-workbook", default=str(INPUT_WORKBOOK), help="Workbook RAIS 2025 de entrada.")
    parser.add_argument("--cnae-dimension-path", default=str(CNAE_DIMENSION_FILE), help="Arquivo cnae_dimensao.xlsx.")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="Diretorio de saida.")
    parser.add_argument("--png-scale", type=float, default=2.2, help="Escala de exportacao do PNG a partir do SVG.")
    parser.add_argument(
        "--browser-path",
        default=None,
        help="Caminho opcional para Edge/Chrome para gerar PNG a partir do SVG em modo headless.",
    )
    return parser


def export_png_with_browser(svg_path: Path, png_path: Path, browser_path: Path) -> None:
    svg_uri = svg_path.resolve().as_uri()
    command = [
        str(browser_path),
        "--headless",
        "--disable-gpu",
        "--default-background-color=00000000",
        f"--screenshot={png_path}",
        "--window-size=3000,2200",
        svg_uri,
    ]
    subprocess.run(command, check=True)


def main() -> None:
    args = build_parser().parse_args()
    input_workbook = Path(args.input_workbook).resolve()
    cnae_dimension_path = Path(args.cnae_dimension_path).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    microregions, municipality_lookup, source_svg = load_ibge_microregion_lookup()
    rais = load_rais_municipio_divisao(input_workbook)
    gr_setor_dimension = load_gr_setor_dimension(cnae_dimension_path)
    stock = build_microregion_sector_stock(rais, gr_setor_dimension, municipality_lookup, microregions)
    predominance = build_predominance_table(stock)

    svg_content = build_map_svg(predominance, source_svg)

    svg_path = output_dir / "mapa_predominancia_gr_setor_microrregiao_sc_rais_2025.svg"
    png_path = output_dir / "mapa_predominancia_gr_setor_microrregiao_sc_rais_2025.png"
    workbook_path = output_dir / "mapa_predominancia_gr_setor_microrregiao_sc_rais_2025.xlsx"

    svg_path.write_text(svg_content, encoding="utf-8")
    if cairosvg is not None:
        cairosvg.svg2png(
            bytestring=svg_content.encode("utf-8"),
            write_to=str(png_path),
            scale=args.png_scale,
            background_color=None,
        )
    elif args.browser_path:
        export_png_with_browser(svg_path, png_path, Path(args.browser_path))
    export_support_workbook(stock, predominance, workbook_path)


if __name__ == "__main__":
    main()
