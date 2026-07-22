"""
pptx_generator.py

Turns the CSV files produced by stats_processor.py (a team-comparison
KPI table + optional per-player stat tables) into a formatted
PowerPoint (.pptx) report, with an optional average-position image
slide.

This module has no Streamlit dependency - it can be run standalone or
imported by streamlit_app.py.
"""

import math
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION


# ============================================================
# LAYOUT CONSTANTS
# ============================================================

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5

DARK_TEXT = RGBColor(0x1A, 0x1A, 0x1A)
LIGHT_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
MUTED_TEXT = RGBColor(0x66, 0x66, 0x66)
BODY_BG = RGBColor(0xFF, 0xFF, 0xFF)
ROW_ALT_BG = RGBColor(0xF2, 0xF2, 0xF2)

METRICS_PER_TABLE_SLIDE = 8
CHART_METRICS = [
    "Total Passes",
    "Passes into Attacking Third",
    "Passes into Penalty Area",
    "Touches in Attacking Third",
    "Open Play Crosses",
    "Possession Lost",
]
TOP_PLAYERS_N = 5
TOP_PLAYERS_SORT_COL = "touches_attacking_third"
TOP_PLAYERS_DISPLAY_COLS = [
    ("Player", "Player"),
    ("total_passes", "Passes"),
    ("forward_pass_percentage", "Fwd Pass %"),
    ("touches_attacking_third", "Att 3rd Touches"),
    ("passes_to_penalty_area", "Passes to Box"),
    ("possession_lost", "Poss. Lost"),
]


# ============================================================
# COLOR HELPERS
# ============================================================

def _hex_to_rgb(hex_color):
    """Accepts '#RRGGBB' or 'RRGGBB' and returns an RGBColor."""
    h = hex_color.strip().lstrip("#")
    if len(h) != 6:
        raise ValueError(f"'{hex_color}' is not a valid hex color (expected 6 hex digits)")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _readable_text_color(rgb_color):
    """Pick black or white text based on background luminance."""
    luminance = (0.299 * rgb_color[0] + 0.587 * rgb_color[1] + 0.114 * rgb_color[2])
    return DARK_TEXT if luminance > 150 else LIGHT_TEXT


# ============================================================
# LOW-LEVEL SLIDE HELPERS
# ============================================================

def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # fully blank layout


def _add_background(slide, prs, rgb_color):
    bg = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        0, 0, prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb_color
    bg.line.fill.background()
    bg.shadow.inherit = False
    # send to back
    spTree = slide.shapes._spTree
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)
    return bg


def _add_textbox(slide, left, top, width, height, text, size=18, bold=False,
                  color=DARK_TEXT, align=PP_ALIGN.LEFT, font_name="Calibri",
                  anchor=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return box


def _add_footer(slide, prs, text, accent_color):
    _add_textbox(
        slide,
        Inches(0.5), Inches(SLIDE_HEIGHT_IN - 0.45),
        Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.35),
        text, size=10, color=MUTED_TEXT
    )
    bar = slide.shapes.add_shape(1, Inches(0), Inches(SLIDE_HEIGHT_IN - 0.06),
                                  prs.slide_width, Inches(0.06))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent_color
    bar.line.fill.background()
    bar.shadow.inherit = False


# ============================================================
# SLIDE BUILDERS
# ============================================================

def add_title_slide(prs, match_title, team_name, team_color, opponent_name, opponent_color, subtitle=None):
    slide = _blank_slide(prs)
    _add_background(slide, prs, team_color)

    _add_textbox(
        slide, Inches(0.8), Inches(2.3), Inches(SLIDE_WIDTH_IN - 1.6), Inches(1.2),
        match_title, size=40, bold=True, color=_readable_text_color(team_color),
        align=PP_ALIGN.LEFT
    )

    vs_text = f"{team_name}  vs  {opponent_name}"
    _add_textbox(
        slide, Inches(0.8), Inches(3.5), Inches(SLIDE_WIDTH_IN - 1.6), Inches(0.7),
        vs_text, size=24, bold=False, color=_readable_text_color(team_color)
    )

    if subtitle:
        _add_textbox(
            slide, Inches(0.8), Inches(4.2), Inches(SLIDE_WIDTH_IN - 1.6), Inches(0.5),
            subtitle, size=14, color=_readable_text_color(team_color)
        )

    # small color chips for each team as a simple legend/motif
    chip_y = Inches(5.3)
    chip = slide.shapes.add_shape(1, Inches(0.8), chip_y, Inches(0.35), Inches(0.35))
    chip.fill.solid()
    chip.fill.fore_color.rgb = team_color
    chip.line.color.rgb = _readable_text_color(team_color)
    _add_textbox(slide, Inches(1.25), chip_y - Inches(0.05), Inches(3), Inches(0.4),
                 team_name, size=14, color=_readable_text_color(team_color))

    chip2 = slide.shapes.add_shape(1, Inches(4.0), chip_y, Inches(0.35), Inches(0.35))
    chip2.fill.solid()
    chip2.fill.fore_color.rgb = opponent_color
    chip2.line.color.rgb = _readable_text_color(team_color)
    _add_textbox(slide, Inches(4.45), chip_y - Inches(0.05), Inches(3), Inches(0.4),
                 opponent_name, size=14, color=_readable_text_color(team_color))

    return slide


def _style_table(table, header_color, header_text_color):
    for c, cell in enumerate(table.rows[0].cells):
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = header_text_color
                r.font.size = Pt(14)

    for row_idx in range(1, len(table.rows)):
        bg = ROW_ALT_BG if row_idx % 2 == 0 else BODY_BG
        for cell in table.rows[row_idx].cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(13)
                    r.font.color.rgb = DARK_TEXT


def add_kpi_table_slide(prs, title, comparison_df, team_name, opponent_name,
                         team_color, opponent_color, section_label=None):
    slide = _blank_slide(prs)
    _add_background(slide, prs, BODY_BG)

    _add_textbox(slide, Inches(0.5), Inches(0.35), Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.6),
                 title, size=28, bold=True, color=DARK_TEXT)
    if section_label:
        _add_textbox(slide, Inches(0.5), Inches(0.95), Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.4),
                     section_label, size=14, color=MUTED_TEXT)

    n_rows = len(comparison_df) + 1
    n_cols = 3
    table_top = Inches(1.55)
    table_left = Inches(0.7)
    table_width = Inches(SLIDE_WIDTH_IN - 1.4)
    table_height = Inches(min(5.3, 0.55 * n_rows))

    gshape = slide.shapes.add_table(n_rows, n_cols, table_left, table_top, table_width, table_height)
    table = gshape.table

    table.columns[0].width = Emu(int(table_width * 0.44))
    table.columns[1].width = Emu(int(table_width * 0.28))
    table.columns[2].width = Emu(int(table_width * 0.28))

    table.cell(0, 0).text = "KPI"
    table.cell(0, 1).text = team_name
    table.cell(0, 2).text = opponent_name

    for i, row in enumerate(comparison_df.itertuples(index=False), start=1):
        table.cell(i, 0).text = str(row.Metric)
        val_team = getattr(row, team_name) if hasattr(row, team_name) else row[1]
        val_opp = getattr(row, opponent_name) if hasattr(row, opponent_name) else row[2]
        table.cell(i, 1).text = "-" if pd.isna(val_team) else str(val_team)
        table.cell(i, 2).text = "-" if pd.isna(val_opp) else str(val_opp)
        for c in range(3):
            table.cell(i, c).text_frame.paragraphs[0].alignment = (
                PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            )

    table.cell(0, 0).text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    table.cell(0, 1).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    table.cell(0, 2).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    _style_table(table, team_color, _readable_text_color(team_color))

    _add_footer(slide, prs, f"{team_name} vs {opponent_name}", team_color)
    return slide


def add_kpi_chart_slide(prs, comparison_df, team_name, opponent_name, team_color, opponent_color):
    slide = _blank_slide(prs)
    _add_background(slide, prs, BODY_BG)

    _add_textbox(slide, Inches(0.5), Inches(0.35), Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.6),
                 "KPI Comparison", size=28, bold=True, color=DARK_TEXT)

    chart_rows = comparison_df[comparison_df["Metric"].isin(CHART_METRICS)]
    chart_rows = chart_rows.set_index("Metric").reindex(CHART_METRICS).reset_index()

    chart_data = CategoryChartData()
    chart_data.categories = chart_rows["Metric"].tolist()
    chart_data.add_series(team_name, chart_rows[team_name].fillna(0).tolist())
    chart_data.add_series(opponent_name, chart_rows[opponent_name].fillna(0).tolist())

    x, y, cx, cy = Inches(0.6), Inches(1.3), Inches(SLIDE_WIDTH_IN - 1.2), Inches(5.5)
    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    )
    chart = graphic_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    series = chart.plots[0].series
    series[0].format.fill.solid()
    series[0].format.fill.fore_color.rgb = team_color
    series[1].format.fill.solid()
    series[1].format.fill.fore_color.rgb = opponent_color

    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.font.size = Pt(9)

    category_axis = chart.category_axis
    category_axis.tick_labels.font.size = Pt(10)
    value_axis = chart.value_axis
    value_axis.has_major_gridlines = False
    value_axis.tick_labels.font.size = Pt(10)

    _add_footer(slide, prs, f"{team_name} vs {opponent_name}", team_color)
    return slide


def add_top_players_slide(prs, title, player_stats_df, team_name, team_color,
                           sort_col=TOP_PLAYERS_SORT_COL, n=TOP_PLAYERS_N):
    slide = _blank_slide(prs)
    _add_background(slide, prs, BODY_BG)

    _add_textbox(slide, Inches(0.5), Inches(0.35), Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.6),
                 title, size=28, bold=True, color=DARK_TEXT)

    df = player_stats_df.copy()
    if sort_col not in df.columns:
        sort_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    df = df.sort_values(sort_col, ascending=False).head(n)

    cols = [c for c in TOP_PLAYERS_DISPLAY_COLS if c[0] in df.columns or c[0] == "Player"]
    n_rows = len(df) + 1
    n_cols = len(cols)

    table_top = Inches(1.4)
    table_left = Inches(0.7)
    table_width = Inches(SLIDE_WIDTH_IN - 1.4)
    table_height = Inches(min(4.8, 0.55 * n_rows))

    gshape = slide.shapes.add_table(n_rows, n_cols, table_left, table_top, table_width, table_height)
    table = gshape.table

    for c, (src_col, label) in enumerate(cols):
        table.cell(0, c).text = label

    for r, (_, row) in enumerate(df.iterrows(), start=1):
        for c, (src_col, label) in enumerate(cols):
            if src_col == "Player":
                val = row["Player"] if "Player" in df.columns else row.name
            else:
                val = row.get(src_col, "")
                if isinstance(val, float):
                    val = round(val, 1)
            table.cell(r, c).text = str(val)
            table.cell(r, c).text_frame.paragraphs[0].alignment = (
                PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            )

    _style_table(table, team_color, _readable_text_color(team_color))
    _add_footer(slide, prs, team_name, team_color)
    return slide


def add_image_slide(prs, title, image_path, team_name, team_color, caption=None):
    slide = _blank_slide(prs)
    _add_background(slide, prs, BODY_BG)

    _add_textbox(slide, Inches(0.5), Inches(0.35), Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.6),
                 title, size=28, bold=True, color=DARK_TEXT)

    max_w = Inches(SLIDE_WIDTH_IN - 1.6)
    max_h = Inches(5.1)
    top = Inches(1.2)

    from PIL import Image
    with Image.open(image_path) as im:
        img_w, img_h = im.size
    aspect = img_w / img_h

    if (max_w / max_h) > aspect:
        height = max_h
        width = Emu(int(height * aspect))
    else:
        width = max_w
        height = Emu(int(width / aspect))

    left = Emu(int((prs.slide_width - width) / 2))

    slide.shapes.add_picture(str(image_path), left, top, width=width, height=height)

    if caption:
        _add_textbox(slide, Inches(0.5), Inches(6.5), Inches(SLIDE_WIDTH_IN - 1.0), Inches(0.4),
                     caption, size=12, color=MUTED_TEXT, align=PP_ALIGN.CENTER)

    _add_footer(slide, prs, team_name, team_color)
    return slide


# ============================================================
# MAIN ENTRY POINT
# ============================================================

def generate_report(
    comparison_csv,
    team_name,
    team_color,
    opponent_name,
    opponent_color,
    output_path,
    match_title="Match Report",
    team_player_stats_csv=None,
    opponent_player_stats_csv=None,
    avg_position_image=None,
    subtitle=None,
):
    """
    Build the full .pptx report and write it to output_path.

    comparison_csv: path (or file-like) to the team KPI comparison CSV
        produced by stats_processor.build_team_comparison(), with
        columns ['Metric', team_name, opponent_name].
    team_color / opponent_color: hex strings, e.g. '#0B3D91'.
    team_player_stats_csv / opponent_player_stats_csv: optional paths
        to per-player stat CSVs (from
        stats_processor.player_stats_to_csv_df()) for the "top
        players" slides.
    avg_position_image: optional path to a PNG/JPG of the team's
        average position map.
    """

    team_rgb = _hex_to_rgb(team_color)
    opponent_rgb = _hex_to_rgb(opponent_color)

    comparison_df = pd.read_csv(comparison_csv)
    # Coerce KPI columns to numeric where possible (PPDA etc. arrive as
    # strings after a CSV round-trip); leave non-numeric as-is.
    for col in (team_name, opponent_name):
        comparison_df[col] = pd.to_numeric(comparison_df[col], errors="coerce")

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)

    # 1. Title slide
    add_title_slide(prs, match_title, team_name, team_rgb, opponent_name, opponent_rgb, subtitle=subtitle)

    # 2. KPI table slide(s), paginated
    n_metric_rows = len(comparison_df)
    n_chunks = max(1, math.ceil(n_metric_rows / METRICS_PER_TABLE_SLIDE))
    for i in range(n_chunks):
        chunk = comparison_df.iloc[i * METRICS_PER_TABLE_SLIDE: (i + 1) * METRICS_PER_TABLE_SLIDE]
        label = f"Part {i + 1} of {n_chunks}" if n_chunks > 1 else None
        add_kpi_table_slide(prs, "Team KPI Comparison", chunk, team_name, opponent_name,
                             team_rgb, opponent_rgb, section_label=label)

    # 3. KPI comparison chart slide
    add_kpi_chart_slide(prs, comparison_df, team_name, opponent_name, team_rgb, opponent_rgb)

    # 4. Top players slides (optional)
    if team_player_stats_csv is not None:
        team_players_df = pd.read_csv(team_player_stats_csv)
        add_top_players_slide(prs, f"{team_name} - Top Performers", team_players_df, team_name, team_rgb)

    if opponent_player_stats_csv is not None:
        opp_players_df = pd.read_csv(opponent_player_stats_csv)
        add_top_players_slide(prs, f"{opponent_name} - Top Performers", opp_players_df, opponent_name, opponent_rgb)

    # 5. Average position image slide (optional)
    if avg_position_image is not None:
        add_image_slide(prs, f"{team_name} - Average Position", avg_position_image, team_name, team_rgb)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return str(output_path)
