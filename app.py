# SUIVI_AUTO/app.py
from __future__ import annotations

from io import BytesIO
from datetime import datetime

import pandas as pd
import polars as pl
import streamlit as st
import plotly.express as px
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from utils import normalize, add_segment, add_intermediaire
from export_utils import _pdf_table_generic, fmt_int, build_product_chart_png
from tasks.suivi_hebdomadaire_auto import run as run_suivi_hebdo
from dash import run as run_dash

# Import sécurisé (évite de faire planter l'app si le fichier n'existe pas)
try:
    from tasks.taux_renouvellement import run as run_taux_renouvellement
    HAS_TAUX_RENOUV = True
except Exception:
    run_taux_renouvellement = None
    HAS_TAUX_RENOUV = False


st.set_page_config(page_title="SUIVI AUTO - Plateforme Automobile", layout="wide")


def load_css(path: str = "assets/styles.css") -> None:
    """Charge le CSS global (si présent)."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except Exception:
        pass


load_css()


def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\u00A0", " ", regex=False)
        .str.upper()
        .str.replace(" ", "_", regex=False)
        .str.replace(r"[^A-Z0-9_]+", "_", regex=True)
        .str.replace(r"_+", "_", regex=True)
        .str.strip("_")
    )
    return df


def _find_col(df: pd.DataFrame, keys: list[str]) -> str | None:
    cols = set(df.columns)
    for k in keys:
        if k in cols:
            return k
    return None


def _to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def _metrics_block(
    label: str,
    cotations: int,
    convertis: int,
    montant_total: float,
    montant_cotation: float,
) -> None:
    taux = (convertis / cotations) if cotations else 0.0
    part = (montant_total / montant_cotation) if montant_cotation else 0.0
    st.markdown(f"### {label}")
    st.metric("Cotations", f"{cotations:,}".replace(",", " "))
    st.metric("Convertis", f"{convertis:,}".replace(",", " "))
    st.metric("Taux de conversion", f"{taux*100:.1f}%")
    st.metric("Montant total", f"{montant_total:,.0f}".replace(",", " "))
    st.metric("Montant cotation", f"{montant_cotation:,.0f}".replace(",", " "))
    st.metric("Part réalisée", f"{part*100:.1f}%")


def _label_fix(text: str) -> str:
    return (
        str(text)
        .replace("Commerce mondial", "Global Business")
        .replace("Courtisans", "Courtiers")
        .replace("Courtisan", "Courtier")
    )


def _compute_lob_table(
    quotes_df: pd.DataFrame,
    policies_df: pd.DataFrame,
    col_quote_id: str,
    col_quote_amount: str,
    col_branche: str,
    col_policy_quote: str,
    col_policy_amount: str,
    col_policy_no: str | None,
) -> pd.DataFrame:
    if not col_branche:
        return pd.DataFrame()

    qlob = quotes_df.copy()
    qlob[col_quote_amount] = (
        qlob[col_quote_amount]
        .astype(str)
        .str.replace(r"[^\d,.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
    )
    qlob[col_quote_amount] = _to_num(qlob[col_quote_amount])
    qlob[col_branche] = qlob[col_branche].astype(str).str.strip().replace("", pd.NA)
    qlob = qlob.dropna(subset=[col_branche])

    q_ids = qlob[col_quote_id].dropna().astype(str).unique().tolist()
    pmatch = policies_df[policies_df[col_policy_quote].isin(q_ids)].copy()
    pmatch[col_policy_amount] = _to_num(pmatch[col_policy_amount])

    pol_map = (
        qlob[[col_quote_id, col_branche]]
        .dropna()
        .drop_duplicates(subset=[col_quote_id])
        .rename(columns={col_quote_id: "_QID", col_branche: "Branche"})
    )
    pmatch = pmatch.merge(
        pol_map,
        left_on=col_policy_quote,
        right_on="_QID",
        how="left",
    )

    cot = (
        qlob.groupby(col_branche)
        .agg(
            Cotations=(col_quote_id, "size"),
            Montant_Cotations=(col_quote_amount, "sum"),
        )
        .reset_index()
        .rename(columns={col_branche: "Étiquettes de lignes"})
    )
    conv = (
        pmatch.groupby("Branche")
        .agg(
            # Compte les cotations converties (clé de correspondance)
            Convertis=(col_policy_quote, "nunique"),
            Montant_convertis=(col_policy_amount, "sum"),
        )
        .reset_index()
        .rename(columns={"Branche": "Étiquettes de lignes"})
    )

    lob = cot.merge(conv, on="Étiquettes de lignes", how="left").fillna(0)
    lob["Taux réalisé en CA"] = lob["Montant_convertis"] / lob["Montant_Cotations"].replace(0, pd.NA)
    lob["Taux conversion"] = lob["Convertis"] / lob["Cotations"].replace(0, pd.NA)
    return lob


def build_courtiers_pdf(
    period_label: str,
    global_metrics: dict,
    courtier_metrics: dict,
    company_metrics: dict,
    kpi_courtiers: dict,
    top20_raw: pd.DataFrame | None,
    no_conv_raw: pd.DataFrame | None,
    lob: pd.DataFrame | None,
    comp_prev_m: pd.DataFrame | None,
    comp_prev_y: pd.DataFrame | None,
) -> bytes:
    buff = BytesIO()
    c = canvas.Canvas(buff, pagesize=landscape(A4))
    W, H = landscape(A4)
    page_num = 1

    def _footer() -> None:
        c.setFillColor(colors.HexColor("#888888"))
        c.setFont("Helvetica", 9)
        c.drawRightString(W - 12 * mm, 8 * mm, f"Page {page_num}")

    def _next_page() -> None:
        nonlocal page_num
        _footer()
        c.showPage()
        page_num += 1

    if top20_raw is None:
        top20_raw = pd.DataFrame()
    if no_conv_raw is None:
        no_conv_raw = pd.DataFrame()
    if lob is None:
        lob = pd.DataFrame()

    def _draw_title(title: str, y: float) -> float:
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(W / 2, y, title)
        return y - 10 * mm

    def _kpi_table_df(metrics: dict, label: str) -> pd.DataFrame:
        rows = [
            ("Segment", label),
            ("Cotations", metrics.get("cotations", 0)),
            ("Convertis", metrics.get("convertis", 0)),
            ("Taux de conversion", metrics.get("convertis", 0) / metrics.get("cotations", 1)),
            ("Montant total", metrics.get("montant_total", 0)),
            ("Montant cotation", metrics.get("montant_cotation", 0)),
            ("Part réalisée", metrics.get("montant_total", 0) / metrics.get("montant_cotation", 1)),
        ]
        out = pd.DataFrame(rows, columns=["Indicateur", "Valeur"])
        out["Valeur"] = out.apply(
            lambda r: f"{float(r['Valeur'])*100:.1f}%" if r["Indicateur"] in {"Taux de conversion", "Part réalisée"} else fmt_int(r["Valeur"]),
            axis=1,
        )
        return out

    # Page 1: synthèse (rendu identique à Streamlit)
    c.setFillColor(colors.white)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y = H - 14 * mm
    c.setFillColor(colors.HexColor("#7ed37e"))
    c.setFont("Helvetica-Bold", 12.5)
    c.drawCentredString(W / 2, y, "Synthèse Globale")
    y -= 8 * mm

    def _draw_metric_card(x: float, y_top: float, w: float, h: float, label: str, value: str, value_color) -> None:
        c.setStrokeColor(colors.HexColor("#d0d7e2"))
        c.setLineWidth(1)
        c.setFillColor(colors.HexColor("#f3f6fa"))
        c.roundRect(x, y_top - h, w, h, 10, stroke=1, fill=1)
        c.setFillColor(colors.HexColor("#2b2f3a"))
        c.setFont("Helvetica", 10)
        c.drawString(x + 4 * mm, y_top - 6.5 * mm, label)
        c.setFillColor(value_color)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x + 4 * mm, y_top - 15 * mm, value)

    def _draw_column(x: float, y_top: float, title: str, metrics: dict) -> None:
        c.setFillColor(colors.HexColor("#1b2a41"))
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x, y_top, title)
        def _safe_div(num, den):
            try:
                den_val = float(den)
                return float(num) / den_val if den_val else 0.0
            except Exception:
                return 0.0

        rows = [
            ("Cotations", fmt_int(metrics.get("cotations", 0)), colors.HexColor("#7ed37e")),
            ("Convertis", fmt_int(metrics.get("convertis", 0)), colors.HexColor("#7ed37e")),
            ("Taux de conversion", f"{_safe_div(metrics.get('convertis', 0), metrics.get('cotations', 0))*100:.1f}%", colors.HexColor("#7ed37e")),
            ("Montant total", fmt_int(metrics.get("montant_total", 0)), colors.HexColor("#ffb067")),
            ("Montant cotation", fmt_int(metrics.get("montant_cotation", 0)), colors.HexColor("#7ed37e")),
            ("Part réalisée", f"{_safe_div(metrics.get('montant_total', 0), metrics.get('montant_cotation', 0))*100:.1f}%", colors.HexColor("#7ed37e")),
        ]
        card_w = 58 * mm
        card_h = 18 * mm
        gap = 5 * mm
        y_cursor = y_top - 8 * mm
        for label, val, col in rows:
            _draw_metric_card(x, y_cursor, card_w, card_h, label, val, col)
            y_cursor -= (card_h + gap)

    col_w = (W - 28 * mm - 10 * mm) / 3
    x0 = 14 * mm
    _draw_column(x0, y, "Global Business", global_metrics)
    _draw_column(x0 + col_w + 5 * mm, y, "Courtiers", courtier_metrics)
    _draw_column(x0 + 2 * (col_w + 5 * mm), y, "Compagnie", company_metrics)

    # KPI Courtiers (sur la page 1, cartes)
    kpi_y = 14 * mm
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 11.5)
    c.drawString(14 * mm, kpi_y + 12 * mm, "KPI Courtiers")
    kpi_w = col_w
    kpi_h = 16 * mm
    kpi_gap = 5 * mm
    _draw_metric_card(
        x0,
        kpi_y + 6 * mm,
        kpi_w,
        kpi_h,
        "Courtiers total",
        fmt_int(kpi_courtiers.get("total", 0)),
        colors.HexColor("#7ed37e"),
    )
    _draw_metric_card(
        x0 + (kpi_w + kpi_gap),
        kpi_y + 6 * mm,
        kpi_w,
        kpi_h,
        "Courtiers actifs",
        fmt_int(kpi_courtiers.get("actifs", 0)),
        colors.HexColor("#7ed37e"),
    )
    _draw_metric_card(
        x0 + 2 * (kpi_w + kpi_gap),
        kpi_y + 6 * mm,
        kpi_w,
        kpi_h,
        "Taux d’activité",
        kpi_courtiers.get("taux", "0.0%"),
        colors.HexColor("#7ed37e"),
    )
    _next_page()

    # Page 2: Top 20 + graphique
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y = H - 16 * mm
    y = _draw_title("Les 20 meilleurs Courtiers", y)
    c.setFont("Helvetica", 10.5)
    c.drawString(14 * mm, y, f"Période : {period_label}")
    y -= 8 * mm

    def _prep_table(df_in: pd.DataFrame | None, cols: list[str]) -> pd.DataFrame:
        if df_in is None or df_in.empty:
            return pd.DataFrame([{cols[0]: "Aucune donnée"}], columns=cols)
        d = df_in.copy()
        for col in cols:
            if col not in d.columns:
                d[col] = ""
        d = d[cols].copy()
        for col in d.columns:
            if col in {"Cotations", "Convertis", "Montant_cotation", "Montant_total"}:
                d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0).map(fmt_int)
            if col in {"Taux conversion", "Part réalisée"}:
                d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0).map(lambda x: f"{x*100:.1f}%")
        return d

    def _draw_table(
        df_in: pd.DataFrame,
        x0: float,
        y: float,
        col_ws: list[float],
        row_type_map: dict[int, str] | None = None,
    ) -> float:
        def _row_style(row):
            rtype = row_type_map.get(row.name, "") if row_type_map else ""
            if rtype == "SECTION":
                return (colors.HexColor("#0f1115"), colors.HexColor("#7ed37e"), True)
            return (
                colors.white if row.name % 2 == 0 else colors.HexColor("#f6f7f9"),
                colors.HexColor("#1b2a41"),
                False,
            )

        def _cell_style(row, col):
            rtype = row_type_map.get(row.name, "") if row_type_map else ""
            if rtype == "SECTION":
                return (colors.HexColor("#0f1115"), colors.HexColor("#7ed37e"), True)
            if col in {"Cotations", "Montant_cotation"}:
                return (colors.HexColor("#eef7ff"), colors.HexColor("#1b2a41"), False)
            return None

        def _cell_align(row, col):
            rtype = row_type_map.get(row.name, "") if row_type_map else ""
            if rtype == "SECTION":
                return "center"
            return None

        def _cell_fontsize(row, col):
            rtype = row_type_map.get(row.name, "") if row_type_map else ""
            if rtype == "SECTION":
                return 12
            return None

        return _pdf_table_generic(
            c,
            df_in,
            x0=x0,
            start_y=y,
            col_ws=col_ws,
            row_h=7.2 * mm,
            header_bg=colors.HexColor("#1f7a3a"),
            header_fg=colors.white,
            grid_color=colors.HexColor("#c7d1db"),
            get_row_style=_row_style,
            align_right_cols=set(df_in.columns[1:]),
            get_cell_style=_cell_style,
            get_cell_align=_cell_align,
            get_cell_fontsize=_cell_fontsize,
        )

    def _top20_chart_png(df_in: pd.DataFrame) -> bytes:
        if df_in is None or df_in.empty:
            return b""
        dfc = df_in.copy()
        if "Courtier" not in dfc.columns or "Montant_total" not in dfc.columns:
            return b""
        dfc["Montant_total"] = pd.to_numeric(dfc["Montant_total"], errors="coerce").fillna(0)
        dfc = dfc.sort_values("Montant_total", ascending=True)
        labels = dfc["Courtier"].astype(str).tolist()
        values = (dfc["Montant_total"] / 1_000_000).tolist()

        fig, ax = plt.subplots(figsize=(11.8, 4.2), dpi=130)
        ax.set_facecolor("white")
        fig.patch.set_facecolor("white")
        ax.barh(labels, values, color="#2e6bd9")
        ax.set_xlabel("Chiffre d'affaires (Millions FCFA)")
        ax.set_ylabel("Courtier")
        ax.set_title("Top 20 Courtiers — Montant total", fontsize=12, fontweight="bold")
        max_v = max(values) if values else 0
        step = 50
        tick_max = (int(max_v / step) + 1) * step if max_v else step
        ax.set_xlim(0, tick_max)
        ax.set_xticks(list(range(0, tick_max + 1, step)))
        ax.grid(axis="x", linestyle=":", alpha=0.5)

        buf = BytesIO()
        plt.tight_layout()
        fig.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.read()

    def _lob_chart_png(df_in: pd.DataFrame, value_col: str, title: str, pct: bool = False) -> bytes:
        if df_in is None or df_in.empty:
            return b""
        if "Étiquettes de lignes" not in df_in.columns or value_col not in df_in.columns:
            return b""
        dfc = df_in.copy()
        dfc[value_col] = (
            dfc[value_col]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        dfc[value_col] = pd.to_numeric(dfc[value_col], errors="coerce").fillna(0)
        labels = dfc["Étiquettes de lignes"].astype(str).tolist()
        values = dfc[value_col].tolist()
        if pct:
            values = [v * 100 for v in values]

        fig, ax = plt.subplots(figsize=(11.8, 4.2), dpi=130)
        ax.set_facecolor("white")
        fig.patch.set_facecolor("white")
        ax.bar(labels, values, color="#2e6bd9")
        ax.set_xlabel("Étiquettes de lignes")
        ax.set_ylabel("Taux (%)" if pct else "Montant (FCFA)")
        ax.set_title(title, fontsize=12, fontweight="bold")
        ax.grid(axis="y", linestyle=":", alpha=0.5)
        ax.tick_params(axis="x", rotation=25)

        buf = BytesIO()
        plt.tight_layout()
        fig.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.read()

    top_cols = ["Courtier", "Cotations", "Convertis", "Montant_cotation", "Montant_total", "Taux conversion", "Part réalisée"]
    top_df = _prep_table(top20_raw, top_cols)
    no_df = _prep_table(no_conv_raw, top_cols)

    col_ws = [55 * mm, 20 * mm, 20 * mm, 28 * mm, 28 * mm, 20 * mm, 20 * mm]
    scale = (W - 28 * mm) / sum(col_ws)
    col_ws = [w * scale for w in col_ws]
    y_table = y
    if not top_df.empty:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(14 * mm, y_table, "Top 20 (Montant total)")
        y_table -= 6 * mm
        _draw_table(top_df.head(20), 14 * mm, y_table, col_ws)
    else:
        c.setFont("Helvetica", 11)
        c.drawString(14 * mm, y_table, "Aucune donnée Top 20.")

    _next_page()

    # Page suivante: graphique Top 20
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y_chart = H - 16 * mm
    y_chart = _draw_title("Top 20 Courtiers — Graphique", y_chart)
    c.setFont("Helvetica", 10.5)
    c.drawString(14 * mm, y_chart, f"Période : {period_label}")
    y_chart -= 10 * mm

    chart = _top20_chart_png(top20_raw)

    if chart:
        chart_h = 90 * mm
        chart_w = W - 28 * mm
        c.drawImage(ImageReader(BytesIO(chart)), 14 * mm, y_chart - chart_h, width=chart_w, height=chart_h, mask="auto")
    else:
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 11)
        c.drawString(14 * mm, y_chart - 10 * mm, "Aucun graphique disponible.")

    _next_page()

    # Page suivante: Courtiers sans contrat réalisé (toujours affiché)
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y2 = H - 16 * mm
    y2 = _draw_title("Courtiers sans contrat réalisé", y2)
    c.setFont("Helvetica", 10.5)
    c.drawString(14 * mm, y2, f"Période : {period_label}")
    y2 -= 8 * mm
    _draw_table(no_df.head(30), 14 * mm, y2, col_ws)
    _next_page()

    # Page 3: LOB (tableau)
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y = H - 16 * mm
    y = _draw_title("Split by Line of Business", y)
    c.setFont("Helvetica", 10.5)
    c.drawString(14 * mm, y, f"Période : {period_label}")
    y -= 8 * mm

    lob_cols = [
        "Étiquettes de lignes",
        "Cotations",
        "Convertis",
        "Montant_Cotations",
        "Montant_convertis",
        "Taux conversion",
        "Taux réalisé en CA",
    ]
    lob_order = [
        "Auto",
        "CAUTION",
        "IA",
        "MRH",
        "MRP",
        "RC",
        "Santé",
        "TRANSPORT_MARCHANDISES",
        "VOYAGE",
    ]
    col_ws = [45 * mm, 20 * mm, 20 * mm, 30 * mm, 30 * mm, 22 * mm, 22 * mm]
    scale = (W - 28 * mm) / sum(col_ws)
    col_ws = [w * scale for w in col_ws]

    def _draw_table_lob(df_in: pd.DataFrame) -> float:
        return _pdf_table_generic(
            c,
            df_in,
            x0=14 * mm,
            start_y=y,
            col_ws=col_ws,
            row_h=7.0 * mm,
            header_bg=colors.HexColor("#1f7a3a"),
            header_fg=colors.white,
            grid_color=colors.HexColor("#c7d1db"),
            body_font=("Helvetica", 9.5),
            get_row_style=lambda row: (
                colors.white if row.name % 2 == 0 else colors.HexColor("#fafbfc"),
                colors.HexColor("#1b2a41"),
                False,
            ),
            align_right_cols=set(df_in.columns[1:]),
            get_cell_style=lambda row, col: (
                (colors.HexColor("#e8f4e8"), colors.HexColor("#1b2a41"), False)
                if col in {"Cotations", "Montant_Cotations"}
                else None
            ),
        )

    if lob is not None and not lob.empty:
        lob_df = lob.copy()
        for col in lob_cols:
            if col not in lob_df.columns:
                lob_df[col] = 0
        lob_df = lob_df[lob_cols].copy()
        lob_df["Étiquettes de lignes"] = lob_df["Étiquettes de lignes"].astype(str)
        lob_df["__order"] = lob_df["Étiquettes de lignes"].apply(
            lambda x: lob_order.index(x) if x in lob_order else len(lob_order)
        )
        lob_df = lob_df.sort_values(["__order", "Étiquettes de lignes"]).drop(columns=["__order"])
        lob_df_fmt = lob_df.copy()
        for col in ["Cotations", "Convertis", "Montant_Cotations", "Montant_convertis"]:
            if col in lob_df_fmt.columns:
                lob_df_fmt[col] = (
                    pd.to_numeric(lob_df_fmt[col], errors="coerce")
                    .fillna(0)
                    .map(fmt_int)
                )
        for col in ["Taux conversion", "Taux réalisé en CA"]:
            if col in lob_df_fmt.columns:
                lob_df_fmt[col] = (
                    pd.to_numeric(lob_df_fmt[col], errors="coerce")
                    .fillna(0)
                    .map(lambda x: f"{x*100:.1f}%")
                )
        _draw_table_lob(lob_df_fmt)
    else:
        _draw_table_lob(pd.DataFrame([{"Étiquettes de lignes": "Aucune donnée"}], columns=lob_cols))

    _next_page()

    # Page 4: graphiques LOB
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y_g = H - 16 * mm
    y_g = _draw_title("Graphiques — Line of Business", y_g)
    c.setFont("Helvetica", 10.5)
    c.drawString(14 * mm, y_g, f"Période : {period_label}")
    y_g -= 10 * mm

    lob_chart = lob.copy()
    if not lob_chart.empty and "Étiquettes de lignes" in lob_chart.columns:
        lob_chart["__order"] = lob_chart["Étiquettes de lignes"].astype(str).apply(
            lambda x: lob_order.index(x) if x in lob_order else len(lob_order)
        )
        lob_chart = lob_chart.sort_values(["__order", "Étiquettes de lignes"]).drop(columns=["__order"])

    ch1 = _lob_chart_png(lob_chart, "Montant_convertis", "Montant convertis par Line of Business", pct=False)
    ch2 = _lob_chart_png(lob_chart, "Taux conversion", "Taux de conversion par Line of Business", pct=True)

    chart_h = 70 * mm
    chart_w = W - 28 * mm
    if ch1:
        c.drawImage(ImageReader(BytesIO(ch1)), 14 * mm, y_g - chart_h, width=chart_w, height=chart_h, mask="auto")
        y_g = y_g - chart_h - 8 * mm
    else:
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 11)
        c.drawString(14 * mm, y_g - 10 * mm, "Aucun graphique Montant convertis.")
        y_g -= 18 * mm
    if ch2:
        c.drawImage(ImageReader(BytesIO(ch2)), 14 * mm, y_g - chart_h, width=chart_w, height=chart_h, mask="auto")
    else:
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 11)
        c.drawString(14 * mm, y_g - 10 * mm, "Aucun graphique Taux de conversion.")
    _next_page()

    # Page 5: Comparative Periodique
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)
    y = H - 16 * mm
    y = _draw_title("Comparative Periodique", y)
    c.setFont("Helvetica", 10.5)
    c.drawString(14 * mm, y, f"Période : {period_label}")
    y -= 12 * mm

    if comp_prev_m is not None and not comp_prev_m.empty:
        row_type_map = {}
        if "_TYPE" in comp_prev_m.columns:
            row_type_map = comp_prev_m["_TYPE"].to_dict()
            comp_prev_m = comp_prev_m.drop(columns=["_TYPE"])
        c.setFont("Helvetica-Bold", 12)
        c.drawString(14 * mm, y, "Comparatif période (mois précédent)")
        y -= 10 * mm
        cols = list(comp_prev_m.columns)
        col_ws = [36 * mm, 40 * mm, 26 * mm, 26 * mm, 20 * mm, 20 * mm]
        scale = (W - 28 * mm) / sum(col_ws)
        col_ws = [w * scale for w in col_ws]
        y = _draw_table(comp_prev_m, 14 * mm, y, col_ws, row_type_map=row_type_map) - 12 * mm
    else:
        c.setFont("Helvetica", 11)
        c.drawString(14 * mm, y, "Aucune donnée comparatif (mois précédent).")
        y -= 12 * mm

    if comp_prev_y is not None and not comp_prev_y.empty:
        row_type_map = {}
        if "_TYPE" in comp_prev_y.columns:
            row_type_map = comp_prev_y["_TYPE"].to_dict()
            comp_prev_y = comp_prev_y.drop(columns=["_TYPE"])
        c.setFont("Helvetica-Bold", 12)
        c.drawString(14 * mm, y, "Comparatif période (même mois année précédente)")
        y -= 10 * mm
        cols = list(comp_prev_y.columns)
        col_ws = [36 * mm, 40 * mm, 26 * mm, 26 * mm, 20 * mm, 20 * mm]
        scale = (W - 28 * mm) / sum(col_ws)
        col_ws = [w * scale for w in col_ws]
        _draw_table(comp_prev_y, 14 * mm, y, col_ws, row_type_map=row_type_map)
    else:
        c.setFont("Helvetica", 11)
        c.drawString(14 * mm, y, "Aucune donnée comparatif (année précédente).")

    _footer()
    c.save()
    buff.seek(0)
    return buff.read()


def run_analyse_courtiers() -> None:
    st.header("🤝 Analyse Courtiers")

    st.markdown(
        """
        <style>
        .courtiers-wrap {
            background: #0f1115;
            border-radius: 18px;
            padding: 18px 18px 6px 18px;
            box-shadow: 0 8px 22px rgba(0,0,0,0.25);
        }
        .courtiers-title {
            color: #ff8c3a;
            font-weight: 800;
            font-size: 22px;
            text-align: center;
            margin: 4px 0 14px 0;
        }
        .courtiers-subtitle {
            color: #7ed37e;
            font-weight: 700;
            font-size: 18px;
            text-align: center;
            margin: 0 0 12px 0;
        }
        .metric-card {
            background: #141824;
            border: 1px solid #2a2f3d;
            border-radius: 14px;
            padding: 14px 12px;
            margin-bottom: 10px;
        }
        .metric-title {
            color: #c7c7c7;
            font-size: 13px;
            margin-bottom: 4px;
        }
        .metric-value {
            color: #7ed37e;
            font-weight: 800;
            font-size: 20px;
        }
        .metric-value.orange {
            color: #ffb067;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.header("📂 Données Analyse Courtiers")
    quotes_file = st.sidebar.file_uploader(
        "Fichier cotations",
        type=["xlsx", "xls"],
        key="courtiers_quotes_file",
    )
    policies_file = st.sidebar.file_uploader(
        "Fichier polices",
        type=["xlsx", "xls"],
        key="courtiers_policies_file",
    )

    if not quotes_file or not policies_file:
        st.info("Charge les 2 fichiers Excel pour lancer l'analyse.")
        st.stop()

    try:
        quotes_df = pl.read_excel(quotes_file).to_pandas()
        policies_df = pl.read_excel(policies_file).to_pandas()
    except Exception as e:
        st.error(f"Erreur lecture fichiers : {e}")
        st.stop()

    quotes_df = _normalize_cols(quotes_df)
    policies_df = _normalize_cols(policies_df)

    col_quote_id = _find_col(quotes_df, ["N_COTATION", "NO_COTATION", "NUM_COTATION", "N_COTATION_"])
    col_quote_date = _find_col(quotes_df, ["DATE_CREATION", "DATE_DE_CREATION", "DATE"])
    col_quote_courtier = _find_col(quotes_df, ["COURTIER", "BROKER"])
    col_quote_amount = _find_col(quotes_df, ["COUT_TOTAL", "COUT", "PRIX", "TOTAL"])

    col_policy_quote = _find_col(policies_df, ["BROKER_QUOTE_NO", "BROKER_QUOTE", "QUOTE_NO"])
    col_policy_amount = _find_col(policies_df, ["AMOUNT", "TOTAL", "PRIME_TTC"])
    col_policy_no = _find_col(policies_df, ["POLICY_NO", "POLICYNO", "POLICY_NUMBER", "POLICY"])
    col_policy_broker = _find_col(policies_df, ["BROKER", "COURTIER"])
    col_policy_status = _find_col(policies_df, ["STATUS", "STATUT"])
    col_policy_issue = _find_col(policies_df, ["ISSUE_DATE", "DATE_EMISSION", "DATE_EMISSION_", "ISSUE DATE"])

    missing = []
    if not col_quote_id:
        missing.append("N° Cotation")
    if not col_quote_date:
        missing.append("DATE CREATION")
    if not col_quote_amount:
        missing.append("Coût Total / Prix")
    if not col_policy_quote:
        missing.append("Broker Quote No")
    if not col_policy_amount:
        missing.append("Amount")

    if missing:
        st.error(f"Colonnes manquantes : {', '.join(missing)}")
        st.stop()

    quotes_df[col_quote_date] = pd.to_datetime(quotes_df[col_quote_date], errors="coerce")
    quotes_df[col_quote_id] = quotes_df[col_quote_id].astype("string").str.strip()
    policies_df[col_policy_quote] = policies_df[col_policy_quote].astype("string").str.strip()
    if col_policy_issue:
        policies_df[col_policy_issue] = pd.to_datetime(policies_df[col_policy_issue], errors="coerce")

    min_date = quotes_df[col_quote_date].min()
    max_date = quotes_df[col_quote_date].max()
    if pd.isna(min_date) or pd.isna(max_date):
        st.error("Aucune date valide dans DATE CREATION.")
        st.stop()

    date_range = st.sidebar.date_input(
        "Période (DATE CREATION)",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
        key="courtiers_date_range",
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_d, end_d = date_range
    else:
        start_d, end_d = min_date.date(), max_date.date()

    # ---------------- Filters (interactive) ----------------
    st.sidebar.header("🎛️ Filtres Analyse Courtiers")
    col_branche = _find_col(quotes_df, ["BRANCHE", "BUSINESS_TYPE"])
    col_statut = _find_col(quotes_df, ["STATUT", "STATUS"])
    col_user = _find_col(quotes_df, ["UTILISATEUR", "USER", "SUBSCRIBER"])

    quotes_date = quotes_df[
        (quotes_df[col_quote_date].dt.date >= start_d)
        & (quotes_df[col_quote_date].dt.date <= end_d)
    ].copy()

    sel_branche = None
    sel_courtier = None
    sel_statut = None
    sel_user = None

    if col_branche:
        branche_vals = sorted(quotes_date[col_branche].dropna().astype(str).unique().tolist())
        sel_branche = st.sidebar.selectbox("Branche", ["Tous"] + branche_vals)
    if col_quote_courtier:
        courtier_vals = sorted(quotes_date[col_quote_courtier].dropna().astype(str).unique().tolist())
        sel_courtier = st.sidebar.selectbox("Courtier", ["Tous"] + courtier_vals)
    if col_statut:
        statut_vals = sorted(quotes_date[col_statut].dropna().astype(str).unique().tolist())
        sel_statut = st.sidebar.selectbox("Statut", ["Tous"] + statut_vals)
    if col_user:
        user_vals = sorted(quotes_date[col_user].dropna().astype(str).unique().tolist())
        sel_user = st.sidebar.selectbox("Utilisateur", ["Tous"] + user_vals)

    def apply_filters(df_in: pd.DataFrame) -> pd.DataFrame:
        df_out = df_in.copy()
        if col_branche and sel_branche and sel_branche != "Tous":
            df_out = df_out[df_out[col_branche].astype(str) == sel_branche].copy()
        if col_quote_courtier and sel_courtier and sel_courtier != "Tous":
            df_out = df_out[df_out[col_quote_courtier].astype(str) == sel_courtier].copy()
        if col_statut and sel_statut and sel_statut != "Tous":
            df_out = df_out[df_out[col_statut].astype(str) == sel_statut].copy()
        if col_user and sel_user and sel_user != "Tous":
            df_out = df_out[df_out[col_user].astype(str) == sel_user].copy()
        return df_out

    quotes_filtered_all = apply_filters(quotes_df)
    quotes_f = quotes_filtered_all[
        (quotes_filtered_all[col_quote_date].dt.date >= start_d)
        & (quotes_filtered_all[col_quote_date].dt.date <= end_d)
    ].copy()

    st.markdown(
        f"## Rapport des Cotations - Quotidien ({start_d.strftime('%d/%m/%Y')}-{end_d.strftime('%d/%m/%Y')})"
    )
    st.markdown("<div class='courtiers-subtitle'>Synthèse Globale</div>", unsafe_allow_html=True)

    def compute_metrics(qdf: pd.DataFrame, pdf: pd.DataFrame) -> dict:
        qdf = qdf.copy()
        qdf[col_quote_amount] = (
            qdf[col_quote_amount]
            .astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(",", ".", regex=False)
        )
        q_ids = qdf[col_quote_id].dropna().astype(str).unique().tolist()
        matched = pdf[pdf[col_policy_quote].isin(q_ids)].copy()
        if col_policy_no:
            convertis = pdf[col_policy_no].dropna().astype(str).nunique()
        else:
            convertis = len(pdf)
        return {
            "cotations": int(len(qdf)),
            "convertis": int(convertis),
            "montant_total": float(_to_num(matched[col_policy_amount]).sum()),
            "montant_cotation": float(_to_num(qdf[col_quote_amount]).sum()),
        }

    policies_base = policies_df.copy()
    if col_policy_broker:
        policies_base[col_policy_broker] = (
            policies_base[col_policy_broker]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    if col_policy_status:
        policies_base = policies_base[
            ~policies_base[col_policy_status]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
            .isin(["CANCELLED", "FAILED"])
        ].copy()

    global_metrics = compute_metrics(quotes_f, policies_base)

    if col_quote_courtier:
        quotes_courtier = quotes_f[quotes_f[col_quote_courtier].notna()].copy()
        quotes_courtier = quotes_courtier[quotes_courtier[col_quote_courtier].astype(str).str.strip() != ""]
    else:
        quotes_courtier = quotes_f.copy()
    if col_policy_broker:
        pol_courtier = policies_base[policies_base[col_policy_broker] != ""]
    else:
        pol_courtier = policies_base
    courtier_metrics = compute_metrics(quotes_courtier, pol_courtier)

    if col_quote_courtier:
        quotes_company = quotes_f[quotes_f[col_quote_courtier].isna() | (quotes_f[col_quote_courtier].astype(str).str.strip() == "")]
    else:
        quotes_company = quotes_f.iloc[0:0].copy()
    if col_policy_broker:
        pol_company = policies_base[policies_base[col_policy_broker] == ""]
    else:
        pol_company = policies_base.iloc[0:0].copy()
    company_metrics = compute_metrics(quotes_company, pol_company)

    st.markdown("<div class='courtiers-wrap'>", unsafe_allow_html=True)
    st.markdown(
        f"<div class='courtiers-title'>Rapport des Cotations - Quotidien ({start_d.strftime('%d/%m/%Y')}-{end_d.strftime('%d/%m/%Y')})</div>",
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    for col, label, metrics in [
        (c1, "Global Business", global_metrics),
        (c2, "Courtiers", courtier_metrics),
        (c3, "Compagnie", company_metrics),
    ]:
        label = _label_fix(label)
        with col:
            st.markdown(f"### {label}")
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-title">Cotations</div>
                    <div class="metric-value">{metrics["cotations"]:,}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Convertis</div>
                    <div class="metric-value">{metrics["convertis"]:,}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Taux de conversion</div>
                    <div class="metric-value">{(metrics["convertis"]/metrics["cotations"]*100) if metrics["cotations"] else 0:.1f}%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Montant total</div>
                    <div class="metric-value orange">{metrics["montant_total"]:,.0f}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Montant cotation</div>
                    <div class="metric-value">{metrics["montant_cotation"]:,.0f}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Part réalisée</div>
                    <div class="metric-value">{(metrics["montant_total"]/metrics["montant_cotation"]*100) if metrics["montant_cotation"] else 0:.1f}%</div>
                </div>
                """.replace(",", " "),
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    # ================= COMPARATIF PERIODE =================
    def _period_label(s: pd.Timestamp, e: pd.Timestamp) -> str:
        months_fr = [
            "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
        ]
        if s.month == e.month and s.year == e.year:
            return f"{months_fr[s.month-1]} ({s.strftime('%d')}-{e.strftime('%d')})"
        return f"{s.strftime('%d/%m')}–{e.strftime('%d/%m')}"

    def _filter_policies_by_period(pdf: pd.DataFrame, s: pd.Timestamp, e: pd.Timestamp) -> pd.DataFrame:
        if not col_policy_issue:
            return pdf
        return pdf[
            (pdf[col_policy_issue].dt.date >= s.date())
            & (pdf[col_policy_issue].dt.date <= e.date())
        ].copy()

    def _metrics_table(s: pd.Timestamp, e: pd.Timestamp) -> dict:
        q_period = quotes_filtered_all[
            (quotes_filtered_all[col_quote_date].dt.date >= s.date())
            & (quotes_filtered_all[col_quote_date].dt.date <= e.date())
        ].copy()

        if col_quote_courtier:
            q_courtier = q_period[q_period[col_quote_courtier].notna()].copy()
            q_courtier = q_courtier[q_courtier[col_quote_courtier].astype(str).str.strip() != ""]
            q_company = q_period[q_period[col_quote_courtier].isna() | (q_period[col_quote_courtier].astype(str).str.strip() == "")]
        else:
            q_courtier = q_period.copy()
            q_company = q_period.iloc[0:0].copy()

        pol_period = _filter_policies_by_period(policies_base, s, e)
        pol_courtier_p = pol_period if not col_policy_broker else pol_period[pol_period[col_policy_broker] != ""]
        pol_company_p = pol_period if not col_policy_broker else pol_period[pol_period[col_policy_broker] == ""]

        return {
            "Global Business": compute_metrics(q_period, pol_period),
            "Courtiers": compute_metrics(q_courtier, pol_courtier_p),
            "Compagnie": compute_metrics(q_company, pol_company_p),
        }

    cur_s = pd.Timestamp(start_d)
    cur_e = pd.Timestamp(end_d)
    prev_month_s = (cur_s - pd.DateOffset(months=1))
    prev_month_e = (cur_e - pd.DateOffset(months=1))
    prev_year_s = (cur_s - pd.DateOffset(years=1))
    prev_year_e = (cur_e - pd.DateOffset(years=1))

    cur_label = _period_label(cur_s, cur_e)
    prev_m_label = _period_label(prev_month_s, prev_month_e)
    prev_y_label = _period_label(prev_year_s, prev_year_e)

    cur_label_m = f"{cur_label} (N)"
    prev_m_label_m = f"{prev_m_label} (N-1)"
    cur_label_y = f"{cur_label} (N)"
    prev_y_label_y = f"{prev_y_label} (N-1)"

    st.subheader("📅 Comparatif période (mois précédent)")
    cur_metrics = _metrics_table(cur_s, cur_e)
    prev_m_metrics = _metrics_table(prev_month_s, prev_month_e)

    def _build_compare_table(cur: dict, prev: dict, cur_label_col: str, prev_label_col: str) -> pd.DataFrame:
        rows = []
        indicators = [
            ("Cotations", "cotations"),
            ("Convertis", "convertis"),
            ("Taux de conversion", "taux_conversion"),
            ("Montant Cotations", "montant_cotation"),
            ("Montant convertis", "montant_total"),
            ("Part réalisée", "part_realisee"),
        ]
        for seg in ["Global Business", "Courtiers", "Compagnie"]:
            rows.append({
                "Segment": seg,
                "Indicateur": "",
                cur_label_col: "",
                prev_label_col: "",
                "Écart absolu": "",
                "Écart relatif": "",
                "_TYPE": "SECTION",
            })
            for ind_label, key in indicators:
                cur_vals = cur[seg].copy()
                prev_vals = prev[seg].copy()
                cur_vals["taux_conversion"] = (
                    cur_vals["convertis"] / cur_vals["cotations"] if cur_vals["cotations"] else 0.0
                )
                cur_vals["part_realisee"] = (
                    cur_vals["montant_total"] / cur_vals["montant_cotation"] if cur_vals["montant_cotation"] else 0.0
                )
                prev_vals["taux_conversion"] = (
                    prev_vals["convertis"] / prev_vals["cotations"] if prev_vals["cotations"] else 0.0
                )
                prev_vals["part_realisee"] = (
                    prev_vals["montant_total"] / prev_vals["montant_cotation"] if prev_vals["montant_cotation"] else 0.0
                )
                cur_v = cur_vals[key]
                prev_v = prev_vals[key]
                delta = cur_v - prev_v
                rel = (delta / prev_v) if prev_v else 0.0
                rows.append({
                    "Segment": "",
                    "Indicateur": ind_label,
                    cur_label_col: cur_v,
                    prev_label_col: prev_v,
                    "Écart absolu": delta,
                    "Écart relatif": rel,
                    "_TYPE": "DATA",
                })
        return pd.DataFrame(rows)

    table_prev_m = _build_compare_table(cur_metrics, prev_m_metrics, cur_label_m, prev_m_label_m)

    def _format_compare(df_in: pd.DataFrame, labels: list[str]) -> pd.DataFrame:
        df_out = df_in.copy()
        rate_inds = {"Taux de conversion", "Part réalisée"}
        section_mask = df_out.get("_TYPE", "") == "SECTION"
        sep_mask = df_out.get("_TYPE", "") == "SEP"
        rate_mask = df_out["Indicateur"].isin(rate_inds) & (~section_mask)
        for col in labels + ["Écart absolu"]:
            if col not in df_out.columns:
                continue
            df_out[col] = pd.to_numeric(df_out[col], errors="coerce")
            df_out.loc[rate_mask, col] = df_out.loc[rate_mask, col].apply(
                lambda x: f"{(x or 0)*100:.1f}%"
            )
            df_out.loc[~rate_mask & ~section_mask & ~sep_mask, col] = df_out.loc[~rate_mask & ~section_mask & ~sep_mask, col].apply(
                lambda x: f"{(x or 0):,.0f}".replace(",", " ")
            )
            df_out.loc[section_mask, col] = ""
            df_out.loc[sep_mask, col] = ""
        if "Écart relatif" in df_out.columns:
            df_out["Écart relatif"] = pd.to_numeric(df_out["Écart relatif"], errors="coerce").fillna(0)
            df_out.loc[section_mask, "Écart relatif"] = ""
            df_out.loc[sep_mask, "Écart relatif"] = ""
            df_out.loc[~section_mask & ~sep_mask, "Écart relatif"] = df_out.loc[~section_mask & ~sep_mask, "Écart relatif"].apply(
                lambda x: f"{x*100:.1f}%"
            )
        return df_out

    st.dataframe(
        _format_compare(table_prev_m, [cur_label_m, prev_m_label_m]).style.apply(
            lambda r: (
                (["background-color:#0f1115;color:#7ed37e;font-weight:800;font-size:16px;text-align:center;"] * len(r)
                 if r.get("_TYPE", "") == "SECTION" else
                 (["background-color:#1b1f1d;color:#ffffff;"] * len(r)
                  if r.name % 2 == 0 else ["background-color:#2c2f2c;color:#ffffff;"] * len(r)))
            ),
            axis=1,
        ).set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", "#1f7a3a"),
                ("color", "white"),
                ("font-weight", "700"),
                ("text-align", "center"),
                ("border", "1px solid #2b3a2f"),
            ]},
            {"selector": "tbody td", "props": [
                ("border", "1px solid #3a3a3a"),
                ("text-align", "center"),
            ]},
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "center"),
                ("font-weight", "700"),
            ]},
        ]).hide(axis="columns", subset=["_TYPE"]),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("📅 Comparatif période (même mois année précédente)")
    prev_y_metrics = _metrics_table(prev_year_s, prev_year_e)
    table_prev_y = _build_compare_table(cur_metrics, prev_y_metrics, cur_label_y, prev_y_label_y)
    st.dataframe(
        _format_compare(table_prev_y, [cur_label_y, prev_y_label_y]).style.apply(
            lambda r: (
                (["background-color:#0f1115;color:#7ed37e;font-weight:800;font-size:16px;text-align:center;"] * len(r)
                 if r.get("_TYPE", "") == "SECTION" else
                 (["background-color:#1b1f1d;color:#ffffff;"] * len(r)
                  if r.name % 2 == 0 else ["background-color:#2c2f2c;color:#ffffff;"] * len(r)))
            ),
            axis=1,
        ).set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", "#1f7a3a"),
                ("color", "white"),
                ("font-weight", "700"),
                ("text-align", "center"),
                ("border", "1px solid #2b3a2f"),
            ]},
            {"selector": "tbody td", "props": [
                ("border", "1px solid #3a3a3a"),
                ("text-align", "center"),
            ]},
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "center"),
                ("font-weight", "700"),
            ]},
        ]).hide(axis="columns", subset=["_TYPE"]),
        use_container_width=True,
        hide_index=True,
    )

    # ---------------- KPI Courtiers ----------------
    total_courtiers = 0
    courtiers_realises = 0
    part_realises = 0.0
    if col_quote_courtier and col_policy_broker:
        total_courtiers = (
            quotes_f[col_quote_courtier]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )
        pol_active = policies_base.copy()
        if col_policy_quote and not quotes_f.empty:
            q_ids = quotes_f[col_quote_id].dropna().astype(str).unique().tolist()
            pol_active = pol_active[pol_active[col_policy_quote].isin(q_ids)].copy()
        courtiers_realises = (
            pol_active[col_policy_broker]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )
        part_realises = (courtiers_realises / total_courtiers) if total_courtiers else 0.0

    # ---------------- PDF Export ----------------
    st.sidebar.header("⬇️ Export PDF (Analyse Courtiers)")
    if st.sidebar.button("🧾 Générer PDF", use_container_width=True, key="btn_pdf_courtiers"):
        top20_pdf = None
        no_conv_pdf = None
        lob_pdf = None
        if col_quote_courtier:
            if "detail" in locals():
                detail_pdf = detail.copy()
            else:
                detail_pdf = pd.DataFrame()
            if detail_pdf.empty:
                qd_pdf = quotes_f.copy()
                qd_pdf[col_quote_amount] = (
                    qd_pdf[col_quote_amount]
                    .astype(str)
                    .str.replace(r"[^\d,.\-]", "", regex=True)
                    .str.replace(",", ".", regex=False)
                )
                qd_pdf[col_quote_amount] = _to_num(qd_pdf[col_quote_amount])
                qt_pdf = (
                    qd_pdf.groupby(col_quote_courtier)
                    .agg(
                        Cotations=(col_quote_id, "count"),
                        Montant_cotation=(col_quote_amount, "sum"),
                    )
                    .reset_index()
                    .rename(columns={col_quote_courtier: "Courtier"})
                )
                pol_pdf = policies_base.copy()
                pol_pdf[col_policy_amount] = _to_num(pol_pdf[col_policy_amount])
                pol_pdf = pol_pdf[pol_pdf[col_policy_quote].isin(qd_pdf[col_quote_id].dropna().astype(str).unique().tolist())]
                if col_policy_broker:
                    pt_pdf = (
                        pol_pdf.groupby(col_policy_broker)
                        .agg(
                            Convertis=(col_policy_no, "nunique") if col_policy_no else (col_policy_quote, "count"),
                            Montant_total=(col_policy_amount, "sum"),
                        )
                        .reset_index()
                        .rename(columns={col_policy_broker: "Courtier"})
                    )
                else:
                    pt_pdf = pd.DataFrame(columns=["Courtier", "Convertis", "Montant_total"])
                detail_pdf = qt_pdf.merge(pt_pdf, on="Courtier", how="left").fillna(0)
                detail_pdf["Taux conversion"] = detail_pdf["Convertis"] / detail_pdf["Cotations"].replace(0, pd.NA)
                detail_pdf["Part réalisée"] = detail_pdf["Montant_total"] / detail_pdf["Montant_cotation"].replace(0, pd.NA)

            if "Courtier" in detail_pdf.columns:
                detail_pdf = detail_pdf[detail_pdf["Courtier"].astype(str).str.strip() != ""].copy()
            top20_pdf = detail_pdf.sort_values("Montant_total", ascending=False).head(20).copy()
            no_conv_pdf = detail_pdf[detail_pdf["Convertis"] == 0].copy()
        if col_branche:
            lob_pdf = _compute_lob_table(
                quotes_f,
                policies_base,
                col_quote_id,
                col_quote_amount,
                col_branche,
                col_policy_quote,
                col_policy_amount,
                col_policy_no,
            )

        kpi_courtiers = {
            "total": total_courtiers if col_quote_courtier else 0,
            "actifs": courtiers_realises if col_quote_courtier else 0,
            "taux": f"{part_realises*100:.1f}%" if col_quote_courtier else "0.0%",
        }

        pdf_bytes = build_courtiers_pdf(
            period_label=f"{start_d.strftime('%d/%m/%Y')} → {end_d.strftime('%d/%m/%Y')}",
            global_metrics=global_metrics,
            courtier_metrics=courtier_metrics,
            company_metrics=company_metrics,
            kpi_courtiers=kpi_courtiers,
            top20_raw=top20_pdf,
            no_conv_raw=no_conv_pdf,
            lob=lob_pdf,
            comp_prev_m=_format_compare(table_prev_m, [cur_label_m, prev_m_label_m]),
            comp_prev_y=_format_compare(table_prev_y, [cur_label_y, prev_y_label_y]),
        )
        st.session_state.courtiers_pdf = pdf_bytes

    if st.session_state.get("courtiers_pdf"):
        st.sidebar.download_button(
            "📄 Télécharger le rapport PDF",
            data=st.session_state.courtiers_pdf,
            file_name=f"rapport_courtiers_{start_d.strftime('%Y%m%d')}_{end_d.strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


        st.subheader("📌 KPI Courtiers")
        k1, k2, k3 = st.columns(3)
        with k1:
            st.metric("Courtiers total", f"{total_courtiers:,}".replace(",", " "))
        with k2:
            st.metric("Courtiers actifs", f"{courtiers_realises:,}".replace(",", " "))
        with k3:
            st.metric("Taux d’activité", f"{part_realises*100:.1f}%")

    if col_quote_courtier:
        st.subheader("Détail par courtier")
        qd = quotes_f.copy()
        qd[col_quote_amount] = (
            qd[col_quote_amount]
            .astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(",", ".", regex=False)
        )
        qd[col_quote_amount] = _to_num(qd[col_quote_amount])
        qt = (
            qd.groupby(col_quote_courtier)
            .agg(
                Cotations=(col_quote_id, "count"),
                Montant_cotation=(col_quote_amount, "sum"),
            )
            .reset_index()
        )
        qt = qt.rename(columns={col_quote_courtier: "Courtier"})

        pol = policies_df.copy()
        pol[col_policy_amount] = _to_num(pol[col_policy_amount])
        pol = pol[pol[col_policy_quote].isin(qd[col_quote_id].dropna().astype(str).unique().tolist())]
        if col_policy_broker:
            pt = (
                pol.groupby(col_policy_broker)
                .agg(
                    Convertis=(col_policy_quote, "nunique"),
                    Montant_total=(col_policy_amount, "sum"),
                )
                .reset_index()
            )
            pt = pt.rename(columns={col_policy_broker: "Courtier"})
        else:
            pt = pd.DataFrame(columns=["Courtier", "Convertis", "Montant_total"])

        detail = qt.merge(pt, on="Courtier", how="left").fillna(0)
        detail["Taux conversion"] = detail["Convertis"] / detail["Cotations"].replace(0, pd.NA)
        detail["Part réalisée"] = detail["Montant_total"] / detail["Montant_cotation"].replace(0, pd.NA)
        detail = detail.sort_values("Montant_total", ascending=False)

        st.dataframe(
            detail.style.format({
                "Cotations": lambda x: f"{int(x):,}".replace(",", " "),
                "Convertis": lambda x: f"{int(x):,}".replace(",", " "),
                "Montant_cotation": lambda x: f"{x:,.0f}".replace(",", " "),
                "Montant_total": lambda x: f"{x:,.0f}".replace(",", " "),
                "Taux conversion": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                "Part réalisée": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("🏆 Top 20 courtiers (plusieurs critères)")
        top_tabs = st.tabs([
            "Nombre de contrats convertis",
            "Montant total",
            "Taux de conversion",
            "Montant cotation",
            "Part réalisée",
        ])
        sort_specs = [
            ("Convertis", top_tabs[0]),
            ("Montant_total", top_tabs[1]),
            ("Taux conversion", top_tabs[2]),
            ("Montant_cotation", top_tabs[3]),
            ("Part réalisée", top_tabs[4]),
        ]
        for sort_col, tab in sort_specs:
            with tab:
                top20 = detail.sort_values(sort_col, ascending=False).head(20).copy()
                st.dataframe(
                    top20.style.format({
                        "Cotations": lambda x: f"{int(x):,}".replace(",", " "),
                        "Convertis": lambda x: f"{int(x):,}".replace(",", " "),
                        "Montant_cotation": lambda x: f"{x:,.0f}".replace(",", " "),
                        "Montant_total": lambda x: f"{x:,.0f}".replace(",", " "),
                        "Taux conversion": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                        "Part réalisée": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )
                fig = px.bar(
                    top20,
                    x="Courtier",
                    y=sort_col,
                    title=f"Top 20 — {sort_col}",
                )
                fig.update_layout(
                    height=420,
                    margin=dict(l=40, r=20, t=60, b=80),
                    xaxis_tickangle=-25,
                )
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("🚫 Courtiers sans contrat réalisé")
        detail_no = detail[detail["Convertis"] == 0].copy()
        st.dataframe(
            detail_no.style.format({
                "Cotations": lambda x: f"{int(x):,}".replace(",", " "),
                "Convertis": lambda x: f"{int(x):,}".replace(",", " "),
                "Montant_cotation": lambda x: f"{x:,.0f}".replace(",", " "),
                "Montant_total": lambda x: f"{x:,.0f}".replace(",", " "),
                "Taux conversion": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                "Part réalisée": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
            }),
            use_container_width=True,
            hide_index=True,
        )

    if col_branche:
        st.subheader("📊 Split by Line of Business")
        lob = _compute_lob_table(
            quotes_f,
            policies_base,
            col_quote_id,
            col_quote_amount,
            col_branche,
            col_policy_quote,
            col_policy_amount,
            col_policy_no,
        )
        lob_order = [
            "Auto",
            "CAUTION",
            "IA",
            "MRH",
            "MRP",
            "RC",
            "Santé",
            "TRANSPORT_MARCHANDISES",
            "VOYAGE",
        ]
        if not lob.empty and "Étiquettes de lignes" in lob.columns:
            lob["__order"] = lob["Étiquettes de lignes"].astype(str).apply(
                lambda x: lob_order.index(x) if x in lob_order else len(lob_order)
            )
            lob = lob.sort_values(["__order", "Étiquettes de lignes"]).drop(columns=["__order"])
        st.session_state["lob_table"] = lob.copy()

        lob_styler = (
            lob.style
                .format({
                    "Cotations": lambda x: f"{int(x):,}".replace(",", " "),
                    "Convertis": lambda x: f"{int(x):,}".replace(",", " "),
                    "Montant_Cotations": lambda x: f"{x:,.0f}".replace(",", " "),
                    "Montant_convertis": lambda x: f"{x:,.0f}".replace(",", " "),
                    "Taux conversion": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                    "Taux réalisé en CA": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                })
                .set_table_styles([
                    {"selector": "thead th", "props": [
                        ("background-color", "#1f7a3a"),
                        ("color", "#ffffff"),
                        ("font-weight", "700"),
                        ("text-align", "center"),
                        ("border-bottom", "1px solid #c7d1db"),
                    ]},
                    {"selector": "tbody td", "props": [
                        ("border-bottom", "1px solid #eef0f2"),
                        ("padding", "10px"),
                        ("font-size", "14px"),
                    ]},
                    {"selector": "tbody td:nth-child(2), tbody td:nth-child(4)", "props": [
                        ("background-color", "#e8f4e8"),
                    ]},
                    {"selector": "tbody td:nth-child(1)", "props": [
                        ("font-weight", "700"),
                    ]},
                ])
        )
        st.dataframe(
            lob_styler,
            use_container_width=True,
            hide_index=True,
        )

        c1, c2 = st.columns(2)
        with c1:
            fig_lob_ca = px.bar(
                lob,
                x="Étiquettes de lignes",
                y="Montant_convertis",
                title="Montant convertis par Line of Business",
            )
            fig_lob_ca.update_layout(
                height=380,
                margin=dict(l=40, r=20, t=60, b=80),
                xaxis_tickangle=-25,
            )
            st.plotly_chart(fig_lob_ca, use_container_width=True)
        with c2:
            fig_lob_tx = px.bar(
                lob,
                x="Étiquettes de lignes",
                y="Taux conversion",
                title="Taux de conversion par Line of Business",
            )
            fig_lob_tx.update_layout(
                height=380,
                margin=dict(l=40, r=20, t=60, b=80),
                xaxis_tickangle=-25,
                yaxis_tickformat=".0%",
            )
            st.plotly_chart(fig_lob_tx, use_container_width=True)

        st.subheader("🚫 Branches sans conversion")
        no_conv = lob[lob["Convertis"] == 0].copy()
        st.dataframe(
            no_conv.style.format({
                "Cotations": lambda x: f"{int(x):,}".replace(",", " "),
                "Convertis": lambda x: f"{int(x):,}".replace(",", " "),
                "Montant_Cotations": lambda x: f"{x:,.0f}".replace(",", " "),
                "Montant_convertis": lambda x: f"{x:,.0f}".replace(",", " "),
                "Taux conversion": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
                "Taux réalisé en CA": lambda x: f"{x*100:.1f}%" if pd.notna(x) else "0.0%",
            }),
            use_container_width=True,
            hide_index=True,
        )


# =========================
#  SIDEBAR - UPLOAD GLOBAL
#  (Permet d'utiliser "Taux de renouvellement" même sans passer par Suivi hebdo)
# =========================
def load_data_global_sidebar() -> None:
    st.sidebar.header("📂 Données (Global)")

    # init session
    if "df_base" not in st.session_state:
        st.session_state.df_base = pd.DataFrame()

    files = st.sidebar.file_uploader(
        "Charger un ou plusieurs fichiers Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="uploader_excel_global",
        help="Ces fichiers alimentent df_base (utilisé par le module Taux de renouvellement).",
    )

    c1, c2 = st.sidebar.columns(2)
    with c1:
        if st.button("🧹 Vider", use_container_width=True, key="btn_clear_global"):
            st.session_state.df_base = pd.DataFrame()
            st.sidebar.success("✅ Données vidées.")
            st.rerun()

    with c2:
        if st.button("🔄 Recharger", use_container_width=True, key="btn_reload_global"):
            st.rerun()

    # lecture fichiers
    if files:
        dfs = []
        for f in files:
            try:
                dfs.append(pl.read_excel(f).to_pandas())
            except Exception as e:
                st.sidebar.error(f"Erreur lecture {getattr(f, 'name', 'fichier')} : {e}")

        if dfs:
            df_raw = pd.concat(dfs, ignore_index=True)

            # Normalisation + champs essentiels
            df = normalize(df_raw)
            df = add_segment(df)
            df = add_intermediaire(df)

            st.session_state.df_base = df
            st.sidebar.success(f"✅ Données chargées : {len(df):,} lignes".replace(",", " "))

    # info date_create (si dispo)
    df_info = st.session_state.get("df_base")
    if isinstance(df_info, pd.DataFrame) and not df_info.empty and "DATE_CREATE" in df_info.columns:
        try:
            dmin = pd.to_datetime(df_info["DATE_CREATE"], errors="coerce").min()
            dmax = pd.to_datetime(df_info["DATE_CREATE"], errors="coerce").max()
            if pd.notna(dmin) and pd.notna(dmax):
                st.sidebar.markdown(
                    f"🗓️ **Date create** : {dmin.date()} → {dmax.date()}"
                )
        except Exception:
            pass


load_data_global_sidebar()


# =========================
#  MENU TÂCHES
# =========================
st.sidebar.title("📌 TÂCHES")

task = st.sidebar.radio(
    "Choisir une tâche",
    [
        "📈 Suivi hebdomadaire Auto",
        "🔁 Taux de renouvellement",
        "🤝 Analyse Courtiers",
        "📊 Dashboard Production",
    ],
    index=0,
    key="task_menu",
)


# =========================
#  ROUTING
#  (⚠️ Ne touche pas ton module Suivi Hebdo)
# =========================
if task == "📈 Suivi hebdomadaire Auto":
    # IMPORTANT: on ne change rien à ce module
    run_suivi_hebdo()

elif task == "🔁 Taux de renouvellement":
    if not HAS_TAUX_RENOUV or run_taux_renouvellement is None:
        st.error("Module 'taux_renouvellement' introuvable. Vérifie que tasks/taux_renouvellement.py existe.")
        st.stop()

    dfb = st.session_state.get("df_base")
    if dfb is None or (isinstance(dfb, pd.DataFrame) and dfb.empty):
        st.info("👉 Charge tes fichiers Excel dans la barre latérale (section « Données (Global) »).")
        st.stop()

    run_taux_renouvellement()

else:
    if task == "🤝 Analyse Courtiers":
        run_analyse_courtiers()
    else:
        run_dash()
