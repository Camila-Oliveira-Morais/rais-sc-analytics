from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]
INPUT_WORKBOOK = PROJECT_ROOT / "data" / "output" / "rais_estabelecimentos_sc_2025.xlsx"
CNAE_DIMENSION_FILE = PROJECT_ROOT / "data" / "cnae_dimensao.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
HEATMAP_COLOR_MIN = "#90BDFF"
HEATMAP_COLOR_MAX = "#3E50E8"


def normalize_text(value: object) -> str:
    return (
        str(value)
        .strip()
        .lower()
        .replace("á", "a")
        .replace("à", "a")
        .replace("â", "a")
        .replace("ã", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
    )


def load_heatmap_base(workbook_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    mesorregiao_sc_comp = pd.read_excel(workbook_path, sheet_name="Mesorregiao_SC_Compet")
    resumo_sc_comp = pd.read_excel(workbook_path, sheet_name="Resumo_SC_Compet")
    return mesorregiao_sc_comp, resumo_sc_comp


def load_sc_comp_categories(dimension_path: Path) -> set[str]:
    dimension = pd.read_excel(dimension_path, sheet_name="CNAE-SUB")
    column_name = next(column for column in dimension.columns if normalize_text(column) == "sc competitiva")
    return {
        str(value).strip()
        for value in dimension[column_name].dropna().unique().tolist()
        if str(value).strip()
    }


def build_heatmap_matrix(
    mesorregiao_sc_comp: pd.DataFrame,
    resumo_sc_comp: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    pivot = mesorregiao_sc_comp.pivot_table(
        index="sc_competitiva",
        columns="mesorregiao_nome",
        values="qtd_vinculos_ativos",
        aggfunc="sum",
        fill_value=0,
    )
    total_by_sector = (
        resumo_sc_comp[["sc_competitiva", "qtd_vinculos_ativos"]]
        .rename(columns={"qtd_vinculos_ativos": "sc_total_vinculos_ativos"})
        .copy()
    )
    column_order = (
        mesorregiao_sc_comp.groupby("mesorregiao_nome", as_index=False)["qtd_vinculos_ativos"]
        .sum()
        .sort_values("qtd_vinculos_ativos", ascending=False)["mesorregiao_nome"]
        .tolist()
    )
    pivot = pivot.reindex(columns=column_order)
    pivot["sc_total_vinculos_ativos"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("sc_total_vinculos_ativos", ascending=False)

    ranking = pivot[["sc_total_vinculos_ativos"]].reset_index()
    ranking["participacao_percentual_sc"] = (
        ranking["sc_total_vinculos_ativos"] / ranking["sc_total_vinculos_ativos"].sum() * 100
    )
    return pivot.reset_index(), ranking


def interpolate_color(color_min: str, color_max: str, value: float) -> str:
    value = max(0.0, min(1.0, value))
    min_rgb = tuple(int(color_min[index : index + 2], 16) for index in (1, 3, 5))
    max_rgb = tuple(int(color_max[index : index + 2], 16) for index in (1, 3, 5))
    rgb = tuple(round(min_rgb[i] + (max_rgb[i] - min_rgb[i]) * value) for i in range(3))
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def build_svg(
    heatmap_matrix: pd.DataFrame,
    output_path: Path,
    title: str,
    subtitle: str,
) -> Path:
    mesorregioes = [column for column in heatmap_matrix.columns if column not in {"sc_competitiva", "sc_total_vinculos_ativos"}]
    values = heatmap_matrix[mesorregioes]
    min_value = float(values.min().min())
    max_value = float(values.max().max())

    cell_width = 125
    cell_height = 28
    left_margin = 380
    top_margin = 140
    right_margin = 60
    bottom_margin = 80
    header_height = 60
    width = left_margin + len(mesorregioes) * cell_width + right_margin
    height = top_margin + len(heatmap_matrix) * cell_height + bottom_margin + header_height

    parts: list[str] = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
        '<rect width="100%" height="100%" fill="#F7FAFF"/>',
        f'<text x="{left_margin}" y="42" font-family="Segoe UI, Arial, sans-serif" font-size="26" font-weight="700" fill="#13225B">{title}</text>',
        f'<text x="{left_margin}" y="70" font-family="Segoe UI, Arial, sans-serif" font-size="14" fill="#42527A">{subtitle}</text>',
        f'<text x="{left_margin}" y="92" font-family="Segoe UI, Arial, sans-serif" font-size="12" fill="#5B6B94">Escala: {HEATMAP_COLOR_MIN} (menor estoque) a {HEATMAP_COLOR_MAX} (maior estoque)</text>',
    ]

    for column_index, mesorregiao in enumerate(mesorregioes):
        x = left_margin + column_index * cell_width + cell_width / 2
        parts.append(
            f'<text x="{x}" y="{top_margin - 18}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="600" fill="#23376E">{mesorregiao}</text>'
        )

    for row_index, (_, row) in enumerate(heatmap_matrix.iterrows()):
        y = top_margin + row_index * cell_height
        label_y = y + 19
        parts.append(
            f'<text x="{left_margin - 12}" y="{label_y}" text-anchor="end" font-family="Segoe UI, Arial, sans-serif" font-size="12" fill="#23376E">{row["sc_competitiva"]}</text>'
        )
        for column_index, mesorregiao in enumerate(mesorregioes):
            x = left_margin + column_index * cell_width
            cell_value = float(row[mesorregiao])
            ratio = 0.0 if max_value == min_value else (cell_value - min_value) / (max_value - min_value)
            fill = interpolate_color(HEATMAP_COLOR_MIN, HEATMAP_COLOR_MAX, ratio)
            text_fill = "#FFFFFF" if ratio > 0.58 else "#17306C"
            parts.append(
                f'<rect x="{x}" y="{y}" width="{cell_width - 4}" height="{cell_height - 4}" rx="4" ry="4" fill="{fill}" />'
            )
            parts.append(
                f'<text x="{x + (cell_width - 4) / 2}" y="{label_y}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="600" fill="{text_fill}">{int(cell_value):,}</text>'.replace(",", ".")
            )

    legend_x = left_margin
    legend_y = top_margin + len(heatmap_matrix) * cell_height + 30
    for step in range(0, 101):
        x = legend_x + step * 4
        color = interpolate_color(HEATMAP_COLOR_MIN, HEATMAP_COLOR_MAX, step / 100)
        parts.append(f'<rect x="{x}" y="{legend_y}" width="4" height="14" fill="{color}" />')

    parts.extend(
        [
            f'<text x="{legend_x}" y="{legend_y + 30}" font-family="Segoe UI, Arial, sans-serif" font-size="11" fill="#42527A">{int(min_value):,}</text>'.replace(",", "."),
            f'<text x="{legend_x + 400}" y="{legend_y + 30}" text-anchor="end" font-family="Segoe UI, Arial, sans-serif" font-size="11" fill="#42527A">{int(max_value):,}</text>'.replace(",", "."),
        ]
    )
    parts.append("</svg>")

    output_path.write_text("".join(parts), encoding="utf-8")
    return output_path


def export_workbook(
    heatmap_matrix: pd.DataFrame,
    ranking: pd.DataFrame,
    base_dataframe: pd.DataFrame,
    metadata: pd.DataFrame,
    output_path: Path,
) -> Path:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        heatmap_matrix.to_excel(writer, sheet_name="Heatmap_SC_Compet", index=False)
        ranking.to_excel(writer, sheet_name="Ranking_Setores", index=False)
        base_dataframe.to_excel(writer, sheet_name="Base_Mesorregiao", index=False)
        metadata.to_excel(writer, sheet_name="Metadados", index=False)

        workbook = writer.book
        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": HEATMAP_COLOR_MAX,
                "border": 0,
            }
        )
        integer_format = workbook.add_format({"num_format": "#,##0"})
        percent_format = workbook.add_format({"num_format": "0.0%"})

        heatmap_sheet = writer.sheets["Heatmap_SC_Compet"]
        ranking_sheet = writer.sheets["Ranking_Setores"]
        base_sheet = writer.sheets["Base_Mesorregiao"]
        metadata_sheet = writer.sheets["Metadados"]

        for sheet in [heatmap_sheet, ranking_sheet, base_sheet, metadata_sheet]:
            sheet.freeze_panes(1, 1)

        for column_index, column_name in enumerate(heatmap_matrix.columns):
            heatmap_sheet.write(0, column_index, column_name, header_format)
            width = 18 if column_name == "sc_competitiva" else 20
            if column_name == "sc_competitiva":
                width = 38
            heatmap_sheet.set_column(column_index, column_index, width, integer_format if column_name != "sc_competitiva" else None)

        mesorregioes = [
            column for column in heatmap_matrix.columns if column not in {"sc_competitiva", "sc_total_vinculos_ativos"}
        ]
        start_col = heatmap_matrix.columns.get_loc(mesorregioes[0])
        end_col = heatmap_matrix.columns.get_loc(mesorregioes[-1])
        heatmap_sheet.conditional_format(
            1,
            start_col,
            len(heatmap_matrix),
            end_col,
            {
                "type": "2_color_scale",
                "min_color": HEATMAP_COLOR_MIN,
                "max_color": HEATMAP_COLOR_MAX,
            },
        )
        heatmap_sheet.autofilter(0, 0, max(len(heatmap_matrix), 1), len(heatmap_matrix.columns) - 1)

        for column_index, column_name in enumerate(ranking.columns):
            ranking_sheet.write(0, column_index, column_name, header_format)
            ranking_sheet.set_column(column_index, column_index, 38 if column_name == "sc_competitiva" else 22)
        ranking_sheet.set_column(
            ranking.columns.get_loc("sc_total_vinculos_ativos"),
            ranking.columns.get_loc("sc_total_vinculos_ativos"),
            22,
            integer_format,
        )
        ranking_sheet.set_column(
            ranking.columns.get_loc("participacao_percentual_sc"),
            ranking.columns.get_loc("participacao_percentual_sc"),
            16,
            percent_format,
        )
        ranking_sheet.autofilter(0, 0, max(len(ranking), 1), len(ranking.columns) - 1)

        for column_index, column_name in enumerate(base_dataframe.columns):
            base_sheet.write(0, column_index, column_name, header_format)
            base_sheet.set_column(column_index, column_index, 22, integer_format if "qtd_" in column_name else None)
        base_sheet.set_column(base_dataframe.columns.get_loc("mesorregiao_nome"), base_dataframe.columns.get_loc("mesorregiao_nome"), 24)
        base_sheet.set_column(base_dataframe.columns.get_loc("sc_competitiva"), base_dataframe.columns.get_loc("sc_competitiva"), 38)
        base_sheet.autofilter(0, 0, max(len(base_dataframe), 1), len(base_dataframe.columns) - 1)

        for column_index, column_name in enumerate(metadata.columns):
            metadata_sheet.write(0, column_index, column_name, header_format)
            metadata_sheet.set_column(column_index, column_index, 48)
        metadata_sheet.autofilter(0, 0, max(len(metadata), 1), len(metadata.columns) - 1)

    return output_path


def build_metadata(
    categories_in_dimension: set[str],
    categories_in_output: set[str],
    ranking: pd.DataFrame,
) -> pd.DataFrame:
    missing_in_output = sorted(categories_in_dimension - categories_in_output)
    extra_in_output = sorted(categories_in_output - categories_in_dimension)
    total_jobs = int(ranking["sc_total_vinculos_ativos"].sum())
    top_sector = ranking.iloc[0]
    rows = [
        {"chave": "fonte_rais", "valor": str(INPUT_WORKBOOK)},
        {"chave": "ano_referencia", "valor": 2025},
        {"chave": "variavel_estoque_empregos", "valor": "qtd_vinculos_ativos"},
        {"chave": "classificacao_setorial", "valor": str(CNAE_DIMENSION_FILE)},
        {"chave": "escala_cor_min", "valor": HEATMAP_COLOR_MIN},
        {"chave": "escala_cor_max", "valor": HEATMAP_COLOR_MAX},
        {"chave": "total_empregos_sc", "valor": total_jobs},
        {"chave": "setor_lider_estoque", "valor": top_sector["sc_competitiva"]},
        {"chave": "setor_lider_vinculos_ativos", "valor": int(top_sector["sc_total_vinculos_ativos"])},
        {"chave": "categorias_sc_competitiva_dimensao", "valor": len(categories_in_dimension)},
        {"chave": "categorias_sc_competitiva_saida", "valor": len(categories_in_output)},
        {"chave": "categorias_sem_emprego_2025", "valor": ", ".join(missing_in_output) if missing_in_output else "nenhuma"},
        {"chave": "categorias_na_saida_nao_na_dimensao", "valor": ", ".join(extra_in_output) if extra_in_output else "nenhuma"},
    ]
    return pd.DataFrame(rows)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera heatmap de estoque de empregos RAIS 2025 por SC Competitiva e mesorregiao."
    )
    parser.add_argument("--input-workbook", default=str(INPUT_WORKBOOK), help="Workbook RAIS 2025 de entrada.")
    parser.add_argument("--cnae-dimension-path", default=str(CNAE_DIMENSION_FILE), help="Arquivo cnae_dimensao.xlsx.")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="Diretorio de saida.")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    input_workbook = Path(args.input_workbook).resolve()
    cnae_dimension_path = Path(args.cnae_dimension_path).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    mesorregiao_sc_comp, resumo_sc_comp = load_heatmap_base(input_workbook)
    categories_in_dimension = load_sc_comp_categories(cnae_dimension_path)
    categories_in_output = set(mesorregiao_sc_comp["sc_competitiva"].dropna().astype(str).str.strip().tolist())

    heatmap_matrix, ranking = build_heatmap_matrix(mesorregiao_sc_comp, resumo_sc_comp)
    metadata = build_metadata(categories_in_dimension, categories_in_output, ranking)

    workbook_path = output_dir / "heatmap_estoque_empregos_sc_competitiva_rais_2025.xlsx"
    svg_path = output_dir / "heatmap_estoque_empregos_sc_competitiva_rais_2025.svg"

    export_workbook(heatmap_matrix, ranking, mesorregiao_sc_comp, metadata, workbook_path)
    build_svg(
        heatmap_matrix,
        svg_path,
        title="Heatmap de Estoque de Empregos RAIS 2025",
        subtitle="SC Competitiva x Mesorregiao de SC | Intensidade medida por qtd_vinculos_ativos",
    )


if __name__ == "__main__":
    main()
