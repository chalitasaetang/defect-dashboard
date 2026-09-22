import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
from io import BytesIO
import os
import urllib.request

import pymupdf  # PyMuPDF — used to turn the report PDF page into a JPG

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart

st.set_page_config(layout="wide", page_icon="📊")

# ---------------------------------------------------------------
# Professional, calm color palette
# ---------------------------------------------------------------
COLOR_PRIMARY = "#2C3E6B"       # deep navy-blue — headers / accents
COLOR_ACCENT = "#C0392B"        # muted brick red — defect emphasis
COLOR_BLUE_BAR = "#1B3A6B"      # deep navy blue for the main defect bar chart
COLOR_BLUE_BAR_LINE = "#0F2647"
COLOR_GOOD = "#2E7D5B"          # muted green — within target
COLOR_BAD = "#C0392B"           # muted red — over target
COLOR_BG_CARD = "#F7F8FA"       # light neutral card background
COLOR_BORDER = "#DCE1E8"        # soft border
COLOR_TEXT_MUTED = "#6B7280"

# ---------------------------------------------------------------
# Date-range theme: the template's primary color and the Top 5 bar
# chart color both switch based on how many days are selected —
# 1 day = blue, 2-27 days = yellow, 28+ days (a full month) = red.
# Applies to both the on-screen Dashboard and the exported JPG.
# ---------------------------------------------------------------
THEME_COLORS = {
    "blue":  {"primary": "#0EA5E9", "bar": "#0284C7", "bar_line": "#075985"},
    "green": {"primary": "#0A9B6D", "bar": "#0A9B6D", "bar_line": "#06724F"},
    "navy":  {"primary": "#0B0F3D", "bar": "#0B0F3D", "bar_line": "#000000"},
}

def theme_for_day_count(n_days: int) -> str:
    """1 day -> blue, 2-27 days -> green, 28+ days (a full month) -> navy."""
    if n_days <= 1:
        return "blue"
    elif n_days <= 27:
        return "green"
    else:
        return "navy"

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Noto+Sans+Thai:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] {{
            font-family: 'Inter', 'Noto Sans Thai', sans-serif;
        }}
        .main {{ background-color: #FCFCFD; }}
        h1, h2, h3 {{ color: {COLOR_PRIMARY}; font-family: 'Inter', 'Noto Sans Thai', sans-serif; }}
        div[data-testid="stMetric"] {{
            background-color: {COLOR_BG_CARD};
            border: 1px solid {COLOR_BORDER};
            border-radius: 10px;
            padding: 10px 14px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Dashboard สรุปตำหนิหลังคัดเกรด")

# ---------------------------------------------------------------
# Translation maps for values coming from the Excel file
# ---------------------------------------------------------------
LOCATION_EN = {
    "เครื่องตัด": "Cutter",
    "เครื่องขัด": "Sander",
    "หน่วยเกรด": "Grader",
    "อื่นๆ": "Other",
    "คลังสินค้า": "Ware house",
    "รถยก": "Forklift",
    "PRESS": "Press",
}

def loc_en(v):
    if pd.isna(v):
        return "-"
    return LOCATION_EN.get(str(v).strip(), str(v))


# ---------------------------------------------------------------
# Thai-capable font for the report (bundled next to app.py; downloads once as fallback)
# ---------------------------------------------------------------
THAI_FONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "NotoSansThai.ttf")
THAI_FONT_URL = (
    "https://raw.githubusercontent.com/google/fonts/main/ofl/notosansthai/"
    "NotoSansThai%5Bwdth%2Cwght%5D.ttf"
)

def ensure_thai_font() -> str | None:
    """Return a usable Thai TTF path. Prefers the font bundled next to app.py;
    falls back to downloading NotoSansThai once (needs internet) if not bundled.
    Needed because defect names (ตำหนิ) stay in Thai per spec."""
    global THAI_FONT_DEBUG
    if os.path.exists(THAI_FONT_PATH):
        size = os.path.getsize(THAI_FONT_PATH)
        if size > 1000:
            return THAI_FONT_PATH
        THAI_FONT_DEBUG = f"Bundled font file found but is only {size} bytes (expected ~200KB) — likely corrupted or an incomplete upload."
    try:
        urllib.request.urlretrieve(THAI_FONT_URL, THAI_FONT_PATH)
        if os.path.getsize(THAI_FONT_PATH) > 1000:
            return THAI_FONT_PATH
    except Exception as e:
        THAI_FONT_DEBUG = f"Bundled font missing, and download fallback failed: {e}"
    return None

_THAI_FONT_REGISTERED = False
THAI_FONT_NAME = "Helvetica"
THAI_FONT_OK = False
THAI_FONT_DEBUG = ""

def register_thai_font():
    global _THAI_FONT_REGISTERED, THAI_FONT_NAME, THAI_FONT_OK, THAI_FONT_DEBUG
    if _THAI_FONT_REGISTERED:
        return THAI_FONT_NAME
    font_path = ensure_thai_font()
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont("NotoSansThai", font_path))
            THAI_FONT_NAME = "NotoSansThai"
            THAI_FONT_OK = True
        except Exception as e:
            THAI_FONT_NAME = "Helvetica"
            THAI_FONT_OK = False
            THAI_FONT_DEBUG = f"Font file found at {font_path} but ReportLab could not load it: {e}"
    _THAI_FONT_REGISTERED = True
    return THAI_FONT_NAME


def build_defect_chart_drawing(top5_df, width_mm=95, height_mm=76, primary_hex=None, bar_hex=None, bar_line_hex=None) -> Drawing:
    """Native ReportLab vector bar chart of Top 5 defects — no external image libs needed."""
    primary_hex = primary_hex or COLOR_PRIMARY
    bar_hex = bar_hex or COLOR_BLUE_BAR
    bar_line_hex = bar_line_hex or COLOR_BLUE_BAR_LINE

    width = width_mm * mm
    height = height_mm * mm
    drawing = Drawing(width, height)

    data_sorted = top5_df.sort_values("จำนวน(Cu)", ascending=False)
    values = [round(v, 2) for v in data_sorted["จำนวน(Cu)"].tolist()]
    has_pct_col = "%ของตำหนิทั้งหมด" in data_sorted.columns
    pcts = data_sorted["%ของตำหนิทั้งหมด"].tolist() if has_pct_col else [None] * len(values)
    # Keep category labels short for the chart (defect names can be long in Thai)
    cats = [str(n)[:12] for n in data_sorted["ตำหนิ"].tolist()]

    chart = VerticalBarChart()
    chart.x = 14 * mm
    chart.y = 16 * mm
    chart.width = width - 20 * mm
    chart.height = height - 24 * mm
    chart.data = [values]
    chart.categoryAxis.categoryNames = cats
    chart.categoryAxis.labels.fontSize = 7
    chart.categoryAxis.labels.fontName = "NotoSansThai" if THAI_FONT_OK else "Helvetica"
    chart.categoryAxis.labels.angle = 0
    chart.categoryAxis.labels.boxAnchor = "n"
    chart.valueAxis.labels.fontSize = 7
    chart.valueAxis.labels.fontName = "NotoSansThai" if THAI_FONT_OK else "Helvetica"
    chart.valueAxis.valueMin = 0
    chart.bars[0].fillColor = colors.HexColor(bar_hex)
    chart.barWidth = 9
    chart.groupSpacing = 10
    chart.bars.strokeColor = colors.HexColor(bar_line_hex)
    chart.bars.strokeWidth = 0.5

    drawing.add(chart)

    # Value labels above each bar — matches the dashboard's "X.XX m³ (Y.Y%)" format
    if values:
        max_v = max(values) if max(values) > 0 else 1
        plot_h = chart.height
        n = len(values)
        bar_area_w = chart.width
        slot_w = bar_area_w / n
        for i, v in enumerate(values):
            bar_h = (v / max_v) * plot_h if max_v else 0
            x_center = chart.x + slot_w * i + slot_w / 2
            # Stagger every other label slightly higher so labels on adjacent
            # bars of similar height don't collide horizontally.
            stagger = (3 * mm) if (i % 2 == 1) else 0
            y_top = chart.y + bar_h + 2 * mm + stagger
            pct = pcts[i]
            label_text = f"{v:.2f} m3 ({pct:.1f}%)" if pct is not None else f"{v:.2f} m3"
            drawing.add(String(x_center, y_top, label_text, fontSize=5.8,
                                fillColor=colors.HexColor(primary_hex),
                                textAnchor="middle", fontName="Helvetica-Bold"))
    return drawing


def build_pdf_report(
    line_display, period_label, production_m3,
    pct_defect, target_defect_pct, total_defect_m3,
    pct_reject, target_reject_pct, reject_m3_input, reject_from_data,
    grade_summary_df, loc_summary_df, top5_df, defect_df_full,
    primary_hex=None, bar_hex=None, bar_line_hex=None, header_mode="range",
) -> bytes:
    """Compact single-page (A4) defect summary report, laid out to mirror the
    dashboard: Grade + Source breakdown stacked on the left, Top-5 chart on
    the right — for a professional, executive-ready one-page report.
    Built as a PDF first, then rasterised to JPG by pdf_bytes_to_jpg().

    primary_hex / bar_hex / bar_line_hex: theme colors picked by the caller
    based on the selected date range (1 day = blue, 2-27 days = yellow,
    28+ days = red); default to the standard navy-blue theme if not given.
    header_mode: "single_day" -> header reads "As of <period_label>";
    "monthly" -> header reads "Monthly report : <period_label>" (period_label
    should be a "<Month> <BE year>" string, e.g. "August 2569"); "range"
    (default) -> header reads "Period : <period_label>".
    """
    font_name = register_thai_font()
    primary_hex = primary_hex or COLOR_PRIMARY
    bar_hex = bar_hex or COLOR_BLUE_BAR
    bar_line_hex = bar_line_hex or COLOR_BLUE_BAR_LINE
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=8 * mm, bottomMargin=8 * mm, leftMargin=12 * mm, rightMargin=12 * mm,
    )

    h1 = ParagraphStyle("h1", fontName=font_name, fontSize=12.5, leading=15.5, spaceAfter=0)
    h2 = ParagraphStyle("h2", fontName=font_name, fontSize=11.5, leading=14, textColor=colors.HexColor(primary_hex), spaceBefore=4, spaceAfter=3)
    body = ParagraphStyle("body", fontName=font_name, fontSize=9.5, leading=12, textColor=colors.HexColor(COLOR_TEXT_MUTED))
    small = ParagraphStyle("small", fontName=font_name, fontSize=8.5, leading=12.5, textColor=colors.HexColor(COLOR_TEXT_MUTED))
    small_bold = ParagraphStyle("small_bold", fontName=font_name, fontSize=9.5, leading=12, textColor=colors.HexColor(primary_hex), spaceBefore=3)

    elems = []

    # --- Header banner (theme-colored bar, mirrors the dashboard's line-name header) ---
    period_word = {"single_day": "As of", "monthly": "Monthly report :"}.get(header_mode, "Period :")
    header_text = f"{line_display} _ Defect Board after Grading &nbsp;&nbsp;{period_word} {period_label}"
    header_tab = Table([[Paragraph(f"<font color='white'>{header_text}</font>", h1)]], colWidths=[171 * mm])
    header_tab.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(primary_hex)),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    elems.append(header_tab)
    elems.append(Spacer(1, 6))

    def status_txt(val, target):
        return "Over Target" if val > target else "Within Target"

    def make_summary_card(title, value_text, target_text, status_over, extra_label, extra_value, accent_hex):
        status_color = colors.HexColor(COLOR_BAD) if status_over else colors.HexColor(COLOR_GOOD)
        status_label = "⚠ Over Target" if status_over else "✓ Within Target"
        card_style = ParagraphStyle("card_title", fontName=font_name, fontSize=8.5, leading=10, textColor=colors.HexColor(COLOR_TEXT_MUTED))
        card_value_style = ParagraphStyle("card_value", fontName=font_name, fontSize=19, leading=22, textColor=status_color)
        card_extra_style = ParagraphStyle("card_extra", fontName=font_name, fontSize=7.5, leading=9.5, textColor=colors.HexColor(COLOR_TEXT_MUTED))
        card_status_style = ParagraphStyle("card_status", fontName=font_name, fontSize=7.5, leading=9.5, textColor=status_color)
        card_target_style = ParagraphStyle("card_target", fontName=font_name, fontSize=6.5, leading=8, textColor=colors.HexColor(COLOR_TEXT_MUTED), alignment=2)

        inner = Table(
            [[Paragraph(f"Target {target_text}", card_target_style)],
             [Paragraph(title, card_style)],
             [Paragraph(value_text, card_value_style)],
             [Paragraph(f"{extra_label}: <b>{extra_value}</b>", card_extra_style)],
             [Paragraph(status_label, card_status_style)]],
            colWidths=[54 * mm],
        )
        inner.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(COLOR_BG_CARD)),
            ("LINEBEFORE", (0, 0), (0, -1), 3, accent_hex),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(COLOR_BORDER)),
        ]))
        return inner

    card1 = make_summary_card(
        "Total Defect %", f"{pct_defect:.2f}%", f"{target_defect_pct:.2f}%",
        pct_defect > target_defect_pct, "Total Defect", f"{total_defect_m3:,.2f} m3",
        colors.HexColor(COLOR_BAD if pct_defect > target_defect_pct else COLOR_GOOD),
    )
    card2 = make_summary_card(
        "Reject after press %", f"{pct_reject:.2f}%", f"{target_reject_pct:.2f}%",
        pct_reject > target_reject_pct, "Reject after press", f"{reject_m3_input:,.2f} m3",
        colors.HexColor(COLOR_BAD if pct_reject > target_reject_pct else COLOR_GOOD),
    )
    card3_style_value = ParagraphStyle("card3_value", fontName=font_name, fontSize=19, leading=22, textColor=colors.HexColor(primary_hex))
    card3_title_style = ParagraphStyle("card3_title", fontName=font_name, fontSize=8.5, leading=10, textColor=colors.HexColor(COLOR_TEXT_MUTED))
    card3_extra_style = ParagraphStyle("card3_extra", fontName=font_name, fontSize=7.5, leading=9.5, textColor=colors.HexColor(COLOR_TEXT_MUTED))
    card3_inner = Table(
        [[Paragraph("Production Volume", card3_title_style)],
         [Paragraph(f"{production_m3:,.2f} m3", card3_style_value)],
         [Paragraph(f"Reject from data (Grade=REJECT): <b>{reject_from_data:,.2f} m3</b>", card3_extra_style)]],
        colWidths=[54 * mm],
    )
    card3_inner.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(COLOR_BG_CARD)),
        ("LINEBEFORE", (0, 0), (0, -1), 3, colors.HexColor(primary_hex)),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(COLOR_BORDER)),
    ]))

    summary_row = Table([[card1, card2, card3_inner]], colWidths=[57 * mm, 57 * mm, 57 * mm])
    summary_row.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))
    elems.append(summary_row)
    elems.append(Spacer(1, 8))

    # --- Right column: Grade table stacked above Source table ---
    gdata = [["Grade", "m3", "%"]]
    if grade_summary_df is not None and not grade_summary_df.empty:
        for _, r in grade_summary_df.iterrows():
            gdata.append([str(r["เกรด"]), f"{r['จำนวน(Cu)']:.2f}", f"{r['%']:.2f}%"])
    ldata = [["Source", "m3"]]
    if loc_summary_df is not None and not loc_summary_df.empty:
        for _, r in loc_summary_df.iterrows():
            ldata.append([loc_en(r["จาก"]), f"{r['จำนวน(Cu)']:.2f}"])

    gtab = Table(gdata, colWidths=[30 * mm, 22 * mm, 20 * mm])
    gtab.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLOR_ACCENT)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor(COLOR_BORDER)),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor(COLOR_BG_CARD)]),
    ]))
    ltab = Table(ldata, colWidths=[52 * mm, 22 * mm])
    ltab.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(primary_hex)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor(COLOR_BORDER)),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor(COLOR_BG_CARD)]),
    ]))
    left_stack_content = Table(
        [[Paragraph("Defect by Grade", h2)],
         [gtab],
         [Spacer(1, 8)],
         [Paragraph("Defect by Source", h2)],
         [ltab]],
        colWidths=[76 * mm],
    )
    left_stack_content.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    # --- Chart column (now on the left, mirrors the dashboard) ---
    chart_col_elems = [Paragraph("Top 5 Defect — Chart", h2)]
    if top5_df is not None and not top5_df.empty:
        chart_col_elems.append(build_defect_chart_drawing(
            top5_df, width_mm=95, height_mm=76,
            primary_hex=primary_hex, bar_hex=bar_hex, bar_line_hex=bar_line_hex,
        ))
    else:
        chart_col_elems.append(Paragraph("No data", small))
    chart_stack = Table([[e] for e in chart_col_elems], colWidths=[95 * mm])
    chart_stack.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    two_col = Table([[chart_stack, left_stack_content]], colWidths=[95 * mm, 76 * mm])
    two_col.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elems.append(two_col)
    elems.append(Spacer(1, 5))

    # Top 5 defects — detail (top 3 records + grade breakdown per defect, matching the dashboard)
    elems.append(Paragraph("Defect TOP 5 — Details", h2))
    elems.append(Spacer(1, 5))
    if top5_df is not None and not top5_df.empty:
        has_pcs = "จำนวน (PCS)" in defect_df_full.columns
        for rank, (_, row) in enumerate(top5_df.sort_values("จำนวน(Cu)", ascending=False).iterrows(), start=1):
            defect_name = row["ตำหนิ"]  # kept in Thai per spec
            defect_total = row["จำนวน(Cu)"]
            defect_pct = row["%ของตำหนิทั้งหมด"]
            elems.append(Paragraph(
                f"<b>{rank}. {defect_name} — {defect_total:.2f} m3 ({defect_pct:.2f}%)</b>", small_bold
            ))
            sub_all = defect_df_full[defect_df_full["ตำหนิ"] == defect_name].copy()
            sub_top3 = sub_all.sort_values("จำนวน(Cu)", ascending=False).head(3)
            for _, srow in sub_top3.iterrows():
                info_text = srow["ข้อมูล"] if pd.notna(srow["ข้อมูล"]) else "(no data)"
                grade = srow["เกรด"] if pd.notna(srow["เกรด"]) else "-"
                pcs_val = srow["จำนวน (PCS)"] if has_pcs and pd.notna(srow["จำนวน (PCS)"]) else None
                pcs_text = f"{int(pcs_val)} sheets " if pcs_val is not None else ""
                elems.append(Paragraph(
                    f"- {info_text} {pcs_text}({srow['จำนวน(Cu)']:.4f} Cu.) — Grade {grade}", small
                ))
            grade_sub = (
                sub_all.groupby("เกรด")["จำนวน(Cu)"].sum().reset_index()
                .sort_values("จำนวน(Cu)", ascending=False)
            )
            grade_line = " | ".join(f"Grade {g['เกรด']}: {g['จำนวน(Cu)']:.2f} m3" for _, g in grade_sub.iterrows())
            elems.append(Paragraph(f"Grade breakdown: {grade_line}", small))
            elems.append(Spacer(1, 5))
    else:
        elems.append(Paragraph("No defect data in selected period", small))

    doc.build(elems)
    return buf.getvalue()


def pdf_bytes_to_jpg(pdf_bytes: bytes, dpi: int = 200, quality: int = 92) -> bytes:
    """Render page 1 of the report PDF to a JPG (A4, single page)."""
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = doc[0]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False)
        return pix.tobytes("jpeg", jpg_quality=quality)
    finally:
        doc.close()


# ---------------------------------------------------------------
# Helpers: Buddhist Era date formatting (kept as BE, per spec)
# ---------------------------------------------------------------
def fmt_be(dt) -> str:
    if pd.isna(dt):
        return "-"
    return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"  # year already stored as BE in the file


# ---------------------------------------------------------------
# File upload + inputs
# ---------------------------------------------------------------

@st.cache_data(show_spinner="Reading Excel file…")
def load_excel_data(file_bytes: bytes):
    """Parse both sheets once per uploaded file. Cached on the file's bytes,
    so touching the date picker or target % fields does NOT re-parse the
    (potentially 79,000-row) Excel file on every rerun — only a genuinely
    new upload does."""
    xls = pd.ExcelFile(BytesIO(file_bytes))

    # Sheet name for the defect log changed from "DATA" to "ตำหนิ" in newer
    # exports; support both so older files still work.
    defect_sheet_name = "ตำหนิ" if "ตำหนิ" in xls.sheet_names else ("DATA" if "DATA" in xls.sheet_names else None)
    if defect_sheet_name is None:
        return None, None, "Could not find the defect sheet (expected 'ตำหนิ' or 'DATA') in the uploaded file."

    # Header is on the 2nd row of the sheet (index 1)
    df = pd.read_excel(xls, sheet_name=defect_sheet_name, header=1)

    required_cols = ["วันที่เกรด", "ตำหนิ", "จำนวน(Cu)", "เกรด", "จาก", "ข้อมูล"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        return None, None, f"Missing required columns: {', '.join(missing)}"

    df["วันที่เกรด"] = pd.to_datetime(df["วันที่เกรด"], errors="coerce")

    # Sheet 2: daily Production Volume + Reject after press, one row per day
    production_sheet_name = "ยอดผลิต" if "ยอดผลิต" in xls.sheet_names else None
    prod_df = pd.DataFrame()
    if production_sheet_name:
        prod_df = pd.read_excel(xls, sheet_name=production_sheet_name, header=1)
        prod_df.columns = [str(c).strip() for c in prod_df.columns]
        date_col = next((c for c in prod_df.columns if "วันที่" in c), None)
        prod_col = next((c for c in prod_df.columns if "ยอดผลิต" in c), None)
        reject_col = next((c for c in prod_df.columns if "Reject" in c or "reject" in c), None)
        if date_col and prod_col and reject_col:
            prod_df[date_col] = pd.to_datetime(prod_df[date_col], errors="coerce")
            prod_df[prod_col] = pd.to_numeric(prod_df[prod_col], errors="coerce").fillna(0)
            prod_df[reject_col] = pd.to_numeric(prod_df[reject_col], errors="coerce").fillna(0)
            prod_df = prod_df.rename(columns={date_col: "วันที่", prod_col: "ยอดผลิต(m3)", reject_col: "Reject(m3)"})
            prod_df = prod_df.dropna(subset=["วันที่"])
        else:
            prod_df = pd.DataFrame()

    return df, prod_df, None


uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df, prod_df, load_error = load_excel_data(uploaded_file.getvalue())
        if load_error:
            st.error(load_error)
            st.stop()

        valid_dates = df["วันที่เกรด"].dropna()
        if valid_dates.empty:
            st.error("No valid dates found in the 'วันที่เกรด' column")
            st.stop()

        if prod_df.empty:
            st.warning(
                "⚠️ Sheet 'ยอดผลิต' (Production Volume + Reject after press) not found or has an unexpected "
                "format — Defect % and Reject % cannot be calculated. Expected columns: วันที่, ยอดผลิต (m3), "
                "Reject After press (m3)."
            )

        min_d = valid_dates.min().date()
        max_d = valid_dates.max().date()

        st.markdown("### ⚙️ Report Settings")

        line_options = ["MDF LINE 1", "MDF LINE 2", "MDF LINE 5", "Other (custom)"]
        c0a, c0b = st.columns([2, 2])
        with c0a:
            production_line = st.selectbox("🏭 Production Line", line_options, index=1)
        with c0b:
            if production_line == "Other (custom)":
                production_line = st.text_input("Enter production line name", value="")

        c1, c2, c3 = st.columns(3)
        with c1:
            date_range = st.date_input(
                "Date Range (BE, as in source file)",
                value=(min_d, max_d),
                min_value=min_d,
                max_value=max_d,
                help="Year shown as in the original file (Buddhist Era) — not converted to CE.",
            )
        with c2:
            target_defect_pct = st.number_input("Target Defect %", min_value=0.0, value=1.0, step=0.1, format="%.2f")
        with c3:
            target_reject_pct = st.number_input("Target Reject after press %", min_value=0.0, value=1.0, step=0.1, format="%.2f")

        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_d, end_d = date_range
        else:
            start_d, end_d = min_d, max_d

        # Date-range theme: recompute on every rerun based on the selected
        # range's day count (inclusive), applied to both the Dashboard and
        # the exported JPG.
        n_days_selected = (end_d - start_d).days + 1
        is_single_day = n_days_selected <= 1
        theme_key = theme_for_day_count(n_days_selected)
        THEME_PRIMARY = THEME_COLORS[theme_key]["primary"]
        THEME_BAR = THEME_COLORS[theme_key]["bar"]
        THEME_BAR_LINE = THEME_COLORS[theme_key]["bar_line"]
        # header_mode drives the report/dashboard title: a single day gets
        # "As of ...", a 28-31 day span (a full month) gets "Monthly report
        # : <Month> <BE year>", anything else keeps "Period : <range>".
        header_mode = "single_day" if is_single_day else ("monthly" if theme_key == "navy" else "range")

        mask = (df["วันที่เกรด"].dt.date >= start_d) & (df["วันที่เกรด"].dt.date <= end_d)
        df_period = df.loc[mask].copy()

        # Defect-only rows (ตำหนิ is not null)
        defect_df = df_period[df_period["ตำหนิ"].notna()].copy()
        defect_df["จำนวน(Cu)"] = pd.to_numeric(defect_df["จำนวน(Cu)"], errors="coerce").fillna(0)

        total_defect_m3 = defect_df["จำนวน(Cu)"].sum()

        # Production Volume + Reject after press — summed from sheet "ยอดผลิต"
        # over the selected date range (no manual entry needed anymore).
        production_m3 = 0.0
        reject_m3_input = 0.0
        if not prod_df.empty:
            prod_mask = (prod_df["วันที่"].dt.date >= start_d) & (prod_df["วันที่"].dt.date <= end_d)
            prod_period = prod_df.loc[prod_mask]
            production_m3 = prod_period["ยอดผลิต(m3)"].sum()
            reject_m3_input = prod_period["Reject(m3)"].sum()

        pct_defect = (total_defect_m3 / production_m3 * 100) if production_m3 > 0 else 0.0
        pct_reject = (reject_m3_input / production_m3 * 100) if production_m3 > 0 else 0.0

        reject_from_data = defect_df.loc[defect_df["เกรด"] == "REJECT", "จำนวน(Cu)"].sum()

        # Pre-compute summaries (also needed by the JPG export button near the top)
        grade_summary = pd.DataFrame()
        loc_summary = pd.DataFrame()
        top5 = pd.DataFrame()
        if not defect_df.empty:
            grade_summary = (
                defect_df.groupby("เกรด")["จำนวน(Cu)"].sum().reset_index()
                .sort_values("จำนวน(Cu)", ascending=False)
            )
            grade_summary["%"] = (grade_summary["จำนวน(Cu)"] / production_m3 * 100) if production_m3 > 0 else 0

            loc_summary = (
                defect_df.groupby("จาก")["จำนวน(Cu)"].sum().reset_index()
                .sort_values("จำนวน(Cu)", ascending=False)
            )

            top5 = (
                defect_df.groupby("ตำหนิ")["จำนวน(Cu)"].sum().reset_index()
                .sort_values("จำนวน(Cu)", ascending=False)
                .head(5)
            )
            top5["%ของตำหนิทั้งหมด"] = (top5["จำนวน(Cu)"] / total_defect_m3 * 100) if total_defect_m3 > 0 else 0

        MONTH_NAMES_EN = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December",
        ]

        def fmt_monthly_label(a_date) -> str:
            """'<Month> <BE year>' using the start date's month, e.g. 'August 2569'.
            The year is taken as stored in the file (Buddhist Era), matching fmt_be."""
            return f"{MONTH_NAMES_EN[a_date.month - 1]} {a_date.year}"

        if header_mode == "single_day":
            period_label_early = fmt_be(pd.Timestamp(start_d))
        elif header_mode == "monthly":
            period_label_early = fmt_monthly_label(start_d)
        else:
            period_label_early = f"{fmt_be(pd.Timestamp(start_d))} - {fmt_be(pd.Timestamp(end_d))}"

        st.markdown("---")

        line_display = production_line if production_line else "Unspecified Line"
        st.markdown(
            f'<div style="background: linear-gradient(90deg, {THEME_PRIMARY} 0%, #4A5B99 100%); '
            f'padding: 18px 24px; border-radius: 12px; margin-bottom: 12px;">'
            f'<span style="color: white; font-size: 28px; font-weight: 700;">🏭 {line_display}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        # ---------------------------------------------------------------
        # JPG export — built on demand (not on every rerun) and cached,
        # since generating it (vector chart + tables + rasterising) is not free.
        # ---------------------------------------------------------------
        @st.cache_data(show_spinner=False)
        def _cached_jpg_bytes(
            line_display, period_label, production_m3,
            pct_defect, target_defect_pct, total_defect_m3,
            pct_reject, target_reject_pct, reject_m3_input, reject_from_data,
            grade_summary_json, loc_summary_json, top5_json, defect_df_json,
            primary_hex, bar_hex, bar_line_hex, header_mode,
        ):
            grade_summary_df = pd.read_json(BytesIO(grade_summary_json.encode()), orient="split")
            loc_summary_df = pd.read_json(BytesIO(loc_summary_json.encode()), orient="split")
            top5_df = pd.read_json(BytesIO(top5_json.encode()), orient="split")
            defect_df_full = pd.read_json(BytesIO(defect_df_json.encode()), orient="split")
            pdf_bytes = build_pdf_report(
                line_display=line_display, period_label=period_label, production_m3=production_m3,
                pct_defect=pct_defect, target_defect_pct=target_defect_pct, total_defect_m3=total_defect_m3,
                pct_reject=pct_reject, target_reject_pct=target_reject_pct,
                reject_m3_input=reject_m3_input, reject_from_data=reject_from_data,
                grade_summary_df=grade_summary_df, loc_summary_df=loc_summary_df,
                top5_df=top5_df, defect_df_full=defect_df_full,
                primary_hex=primary_hex, bar_hex=bar_hex, bar_line_hex=bar_line_hex,
                header_mode=header_mode,
            )
            return pdf_bytes_to_jpg(pdf_bytes)

        # Check the Thai font status now (not just inside build_pdf_report) so
        # the warning below reflects the real state, with a specific reason.
        register_thai_font()

        dl_col, note_col = st.columns([1, 3])
        with dl_col:
            prepare_jpg = st.button("🖼️ Prepare JPG for download", use_container_width=True, type="primary")
        with note_col:
            if not THAI_FONT_OK:
                reason = f" ({THAI_FONT_DEBUG})" if THAI_FONT_DEBUG else ""
                st.warning(
                    f"⚠️ Thai font not available for JPG export{reason}. Defect names in the image may not render "
                    "correctly. Try re-uploading NotoSansThai.ttf to the repo (as a binary file, not via 'Edit' "
                    "in the browser), or check that this machine has internet access."
                )

        if prepare_jpg:
            try:
                with st.spinner("Building JPG…"):
                    jpg_bytes = _cached_jpg_bytes(
                        line_display, period_label_early, production_m3,
                        pct_defect, target_defect_pct, total_defect_m3,
                        pct_reject, target_reject_pct, reject_m3_input, reject_from_data,
                        grade_summary.to_json(orient="split"), loc_summary.to_json(orient="split"),
                        top5.to_json(orient="split"), defect_df.to_json(orient="split"),
                        THEME_PRIMARY, THEME_BAR, THEME_BAR_LINE, header_mode,
                    )
                st.download_button(
                    label="⬇️ Download JPG",
                    data=jpg_bytes,
                    file_name=f"dashboard_{line_display}_{start_d}_{end_d}.jpg".replace(" ", "_"),
                    mime="image/jpeg",
                    use_container_width=True,
                )
            except Exception as jpg_err:
                st.error(f"⚠️ Failed to generate JPG: {jpg_err}")

        st.markdown("<div style='margin-top:4px;'></div>", unsafe_allow_html=True)

        # ---------------------------------------------------------------
        # Summary boxes: Total Defect %, Reject %
        # ---------------------------------------------------------------
        period_label = period_label_early
        if header_mode == "single_day":
            period_label_display = f"As of {period_label}"
        elif header_mode == "monthly":
            period_label_display = f"Monthly report : {period_label}"
        else:
            period_label_display = f"Period: {period_label}"

        def summary_box(title, value_pct, target_pct, extra_m3_label, extra_m3_value):
            over = value_pct > target_pct
            status_color = COLOR_BAD if over else COLOR_GOOD
            status_text = '⚠️ Over Target' if over else '✅ Within Target'
            st.markdown(
                f'<div style="background-color:{COLOR_BG_CARD}; border:1px solid {COLOR_BORDER}; '
                f'border-left:5px solid {status_color}; border-radius:10px; padding:18px; '
                f'position:relative; margin-bottom:10px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);">'
                f'<div style="position:absolute; top:10px; right:16px; font-size:12px; color:{COLOR_TEXT_MUTED}; '
                f'background-color:white; border:1px solid {COLOR_BORDER}; border-radius:6px; padding:2px 8px;">'
                f'🎯 Target {target_pct:.2f}%</div>'
                f'<div style="font-size:15px; color:{COLOR_TEXT_MUTED}; font-weight:600;">{title}</div>'
                f'<div style="font-size:38px; font-weight:800; color:{status_color}; margin:4px 0;">{value_pct:.2f}%</div>'
                f'<div style="font-size:13px; color:{COLOR_TEXT_MUTED};">{extra_m3_label}: <b>{extra_m3_value:,.2f} m³</b></div>'
                f'<div style="font-size:12px; color:{status_color}; font-weight:600; margin-top:4px;">{status_text}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.markdown(f"#### 🗓️ {period_label_display}")
        box1, box2, box3 = st.columns(3)
        with box1:
            summary_box("Total Defect %", pct_defect, target_defect_pct, "Total Defect", total_defect_m3)
        with box2:
            summary_box("Reject after press %", pct_reject, target_reject_pct, "Reject after press", reject_m3_input)
        with box3:
            st.markdown(
                f'<div style="background-color:{COLOR_BG_CARD}; border:1px solid {COLOR_BORDER}; border-left:5px solid {THEME_PRIMARY}; border-radius:10px; padding:18px; margin-bottom:10px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);">'
                f'<div style="font-size:15px; color:{COLOR_TEXT_MUTED}; font-weight:600;">Production Volume</div>'
                f'<div style="font-size:38px; font-weight:800; color:{THEME_PRIMARY}; margin:4px 0;">{production_m3:,.2f}</div>'
                f'<div style="font-size:13px; color:{COLOR_TEXT_MUTED};">m³</div>'
                f'<div style="font-size:12px; color:{COLOR_TEXT_MUTED}; margin-top:4px;">Reject from data (Grade=REJECT): <b>{reject_from_data:,.2f} m³</b></div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        if production_m3 <= 0:
            st.warning(
                "⚠️ No Production Volume data found for the selected date range in sheet 'ยอดผลิต' — "
                "Defect % and Reject after press % cannot be calculated."
            )

        st.markdown("---")

        # ---------------------------------------------------------------
        # Left: Grade breakdown (top) + Source breakdown (bottom), stacked
        # Right: Top 5 Defect chart — sized to match the left column's height
        # ---------------------------------------------------------------
        def render_chip_row(title_html, df_rows, label_col, translate_fn=None, pct_col=None, pct_denominator=None):
            st.markdown(title_html, unsafe_allow_html=True)
            if df_rows.empty:
                st.caption("No defect data in selected period")
                return
            chips = []
            for _, r in df_rows.iterrows():
                label = translate_fn(r[label_col]) if translate_fn else r[label_col]
                if pct_col is not None:
                    pct_val = r[pct_col]
                elif pct_denominator:
                    pct_val = (r["จำนวน(Cu)"] / pct_denominator * 100) if pct_denominator else 0
                else:
                    pct_val = 0
                chips.append(
                    f'<div style="flex:1 1 0; display:flex; flex-direction:column; align-items:center; justify-content:center; '
                    f'background-color:{COLOR_BG_CARD}; border:1px solid {COLOR_BORDER}; '
                    f'border-radius:10px; padding:12px 6px; min-width:0; text-align:center;">'
                    f'<span style="font-size:12px; color:{COLOR_TEXT_MUTED}; font-weight:600; white-space:nowrap;">{label}</span>'
                    f'<span style="font-size:17px; font-weight:800; color:{THEME_PRIMARY}; margin-top:2px;">{r["จำนวน(Cu)"]:.2f} m³</span>'
                    f'<span style="font-size:11px; color:{COLOR_ACCENT}; margin-top:2px;">{pct_val:.2f}%</span>'
                    f'</div>'
                )
            st.markdown(
                f"<div style='display:flex; gap:8px; width:100%;'>{''.join(chips)}</div>",
                unsafe_allow_html=True,
            )

        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.markdown(
                f"<div style='font-size:18px; font-weight:700; color:{THEME_PRIMARY}; margin-bottom:6px;'>🔵 Top 5 Defect — {period_label_display}</div>",
                unsafe_allow_html=True,
            )
            if not top5.empty:
                top5_sorted = top5.sort_values("จำนวน(Cu)", ascending=False).copy()
                top5_sorted["label"] = top5_sorted.apply(
                    lambda r: f"{r['จำนวน(Cu)']:.2f} m³ ({r['%ของตำหนิทั้งหมด']:.1f}%)", axis=1
                )

                fig = px.bar(
                    top5_sorted,
                    x="ตำหนิ",
                    y="จำนวน(Cu)",
                    text="label",
                    color_discrete_sequence=[THEME_BAR],
                )
                fig.update_traces(
                    textposition="outside",
                    marker_line_color=THEME_BAR_LINE, marker_line_width=1,
                    textfont_size=13, textfont_color="black",
                )
                fig.update_layout(
                    xaxis_title="", yaxis_title="Cu. (m³)", height=300,
                    plot_bgcolor="white", paper_bgcolor="white",
                    font=dict(color="black", size=13),
                    margin=dict(t=20, b=10, l=10, r=10),
                )
                fig.update_yaxes(gridcolor=COLOR_BORDER, tickfont=dict(color="black", size=13),
                                  title_font=dict(color="black", size=13))
                fig.update_xaxes(tickfont=dict(color="black", size=13))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.caption("No defect data in selected period")

        with col_right:
            render_chip_row(
                f"<div style='font-size:18px; font-weight:700; color:{THEME_PRIMARY}; margin-bottom:6px;'>🏷️ Defect by Grade</div>",
                grade_summary, "เกรด", translate_fn=lambda g: f"Grade {g}", pct_col="%",
            )
            st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
            render_chip_row(
                f"<div style='font-size:18px; font-weight:700; color:{THEME_PRIMARY}; margin-bottom:6px;'>🏭 Defect by Source</div>",
                loc_summary, "จาก", translate_fn=loc_en, pct_denominator=total_defect_m3,
            )

        st.markdown("---")

        # -----------------------------------------------------------
        # Detail per top-5 defect: top 3 records from column S (ข้อมูล)
        # -----------------------------------------------------------
        if not top5.empty:
            st.markdown("### 📋 Defect TOP 5")

            has_pcs_col = "จำนวน (PCS)" in defect_df.columns

            for rank, (_, row) in enumerate(top5_sorted.iterrows(), start=1):
                defect_name = row["ตำหนิ"]  # kept in Thai per spec
                defect_total = row["จำนวน(Cu)"]
                defect_pct = row["%ของตำหนิทั้งหมด"]

                st.markdown(
                    f'<div style="border-left:4px solid {COLOR_ACCENT}; padding-left:12px; margin-top:14px;">'
                    f'<span style="font-size:17px; font-weight:700; color:{THEME_PRIMARY};">'
                    f'{rank}. {defect_name} — {defect_total:.2f} m³ ({defect_pct:.2f}%)</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

                sub_all = defect_df[defect_df["ตำหนิ"] == defect_name].copy()
                sub_top3 = sub_all.sort_values("จำนวน(Cu)", ascending=False).head(3)

                for _, srow in sub_top3.iterrows():
                    info_text = srow["ข้อมูล"] if pd.notna(srow["ข้อมูล"]) else "(no data)"
                    grade = srow["เกรด"] if pd.notna(srow["เกรด"]) else "-"
                    pcs_val = srow["จำนวน (PCS)"] if has_pcs_col and pd.notna(srow["จำนวน (PCS)"]) else None
                    pcs_text = f"{int(pcs_val)} sheets " if pcs_val is not None else ""
                    st.markdown(
                        f'<div style="font-size:14px; margin:2px 0;">- {info_text} {pcs_text}'
                        f'({srow["จำนวน(Cu)"]:.4f} Cu.) — Grade {grade}</div>',
                        unsafe_allow_html=True,
                    )

                # Grade subtotal for this defect, largest m3 first — same font size as the top-3 detail above
                grade_sub = (
                    sub_all.groupby("เกรด")["จำนวน(Cu)"]
                    .sum()
                    .reset_index()
                    .sort_values("จำนวน(Cu)", ascending=False)
                )
                st.markdown(
                    f"<div style='margin-top:6px; margin-bottom:2px; font-size:14px; color:{COLOR_TEXT_MUTED}; font-weight:600;'>Grade breakdown for this defect:</div>",
                    unsafe_allow_html=True,
                )
                for _, grow in grade_sub.iterrows():
                    st.markdown(
                        f'<div style="margin-left:8px; font-size:14px;">▸ Grade <b>{grow["เกรด"]}</b> '
                        f'<span style="color:{COLOR_ACCENT}; font-weight:600;">{grow["จำนวน(Cu)"]:.2f} m³</span></div>',
                        unsafe_allow_html=True,
                    )

                st.markdown("")
        else:
            st.info("No defect data in the selected period — Top 5 chart unavailable")

        st.markdown("---")

        # ---------------------------------------------------------------
        # Raw data table
        # ---------------------------------------------------------------
        st.header("📉 Full Data Table — Selected Period")
        st.dataframe(df_period, use_container_width=True)

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("💡 Please upload your Excel file above to open the dashboard.")
