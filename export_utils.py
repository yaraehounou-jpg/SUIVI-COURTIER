# SUIVI_AUTO/export_utils.py
from __future__ import annotations

from io import BytesIO
from datetime import datetime

import pandas as pd

from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


# ----------------- helpers -----------------
def fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except Exception:
        return "0"


def safe_pct(x, digits: int = 1) -> str:
    try:
        return f"{float(x)*100:.{digits}f}%"
    except Exception:
        return "0.0%"


# ---------------- PDF: cards ----------------
def _fit_font_size(text: str, font_name: str, max_size: float, min_size: float, max_width: float) -> float:
    """Retourne une taille de police ajustee pour tenir dans la largeur."""
    size = max_size
    txt = "" if text is None else str(text)
    while size > min_size and stringWidth(txt, font_name, size) > max_width:
        size -= 0.5
    return max(size, min_size)


def _truncate_to_width(text: str, font_name: str, font_size: float, max_width: float) -> str:
    """Tronque proprement un texte pour tenir dans une largeur donnee."""
    txt = "" if text is None else str(text)
    if stringWidth(txt, font_name, font_size) <= max_width:
        return txt
    ell = "..."
    out = txt
    while out and stringWidth(out + ell, font_name, font_size) > max_width:
        out = out[:-1]
    return (out + ell) if out else ell


def _split_pct_lines(pct_text: str) -> list[str]:
    """Prepare 1-2 lines for KPI pct block so values stay readable in PDF."""
    txt = "" if pct_text is None else str(pct_text).strip()
    if not txt:
        return []
    # Support explicit line break from streamlit html strings.
    txt = txt.replace("<br>", "\n")
    parts = [p.strip() for p in txt.split("\n") if p.strip()]
    if len(parts) >= 2:
        return parts[:2]
    # If a long single line with separators, split at last separator.
    if " | " in txt and len(txt) > 26:
        left, right = txt.rsplit(" | ", 1)
        return [left.strip(), right.strip()]
    return [txt]


def _pdf_card(
    c,
    x,
    y,
    w,
    h,
    value,
    label,
    value_color,
    bg_color,
    border_color,
    border_w=2,
    pct=None,
    pct_color=colors.black,
    pct_align: str = "right",
    pct_y_offset: float = 0.0,
    pct_x_offset: float = 0.0,
    pct_font_max: float = 10.5,
):
    c.setLineWidth(border_w)
    c.setStrokeColor(border_color)
    c.setFillColor(bg_color)
    c.roundRect(x, y, w, h, 10, stroke=1, fill=1)

    c.setFillColor(value_color)
    value_size = _fit_font_size(str(value), "Helvetica-Bold", 19, 11, w - 20)
    c.setFont("Helvetica-Bold", value_size)
    c.drawString(x + 10, y + h - 28, str(value))

    if pct is not None:
        # pct block: right aligned, one or two lines, full card width
        pct_lines = _split_pct_lines(str(pct))
        if pct_lines:
            pct_max_w = max(40, w - 20)
            longest = max(pct_lines, key=len)
            pct_size = _fit_font_size(longest, "Helvetica-Bold", float(pct_font_max), 7.0, pct_max_w)
            c.setFont("Helvetica-Bold", pct_size)
            c.setFillColor(pct_color)
            y0 = y + h - 40 + float(pct_y_offset)
            step = max(8, pct_size + 1)
            for i, line in enumerate(pct_lines[:2]):
                line_txt = _truncate_to_width(line, "Helvetica-Bold", pct_size, pct_max_w)
                if pct_align == "center":
                    c.drawCentredString(x + w / 2 + float(pct_x_offset), y0 - i * step, line_txt)
                else:
                    c.drawRightString(x + w - 10 + float(pct_x_offset), y0 - i * step, line_txt)

    # label en bas a gauche, taille reduite et tronquee
    c.setFillColor(colors.black)
    label_size = _fit_font_size(str(label), "Helvetica-BoldOblique", 9.5, 7.5, w - 20)
    c.setFont("Helvetica-BoldOblique", label_size)
    label_txt = _truncate_to_width(str(label), "Helvetica-BoldOblique", label_size, w - 20)
    c.drawString(x + 10, y + 10, label_txt)


def _pdf_ratio_card(c, x, y, w, h, ratio):
    c.setLineWidth(3)
    c.setStrokeColor(colors.HexColor("#6b5f84"))
    c.setFillColor(colors.HexColor("#6b5f84"))
    c.roundRect(x, y, w, h, 10, stroke=1, fill=1)

    c.setFillColor(colors.HexColor("#f1d26a"))
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(x + w / 2, y + h / 2 + 6, ratio)

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(x + w / 2, y + h / 2 - 18, "Ratio Objectif Annuel")


# ---------------- PDF: table generic ----------------
def _pdf_table_generic(
    c,
    df_table: pd.DataFrame,
    x0: float,
    start_y: float,
    col_ws: list[float],
    row_h: float = 8.5 * mm,
    header_bg=colors.HexColor("#F4D35E"),
    header_fg=colors.black,
    header_font=("Helvetica-Bold", 10.5),
    body_font=("Helvetica", 10.0),
    grid_color=colors.HexColor("#B0B0B0"),
    get_row_style=None,  # fn(row)->(bg, fg, bold:bool)
    align_right_cols: set[str] | None = None,
    get_cell_style=None,  # fn(row, col)->(bg, fg, bold:bool) or None
    get_cell_align=None,  # fn(row, col)->"left"|"center"|"right"
    get_cell_fontsize=None,  # fn(row, col)->int|float
):
    align_right_cols = align_right_cols or set()
    cols = list(df_table.columns)

    # Header
    c.setStrokeColor(grid_color)
    c.setLineWidth(1)
    c.setFillColor(header_bg)
    c.rect(x0, start_y - row_h, sum(col_ws), row_h, stroke=1, fill=1)

    c.setFillColor(header_fg)
    c.setFont(*header_font)
    cx = x0
    for i, col in enumerate(cols):
        c.drawCentredString(cx + col_ws[i] / 2, start_y - row_h + 3.0 * mm, str(col))
        cx += col_ws[i]

    y = start_y - row_h

    # Body
    for r in range(len(df_table)):
        row = df_table.iloc[r]
        if get_row_style:
            bg, fg, is_bold = get_row_style(row)
        else:
            bg, fg, is_bold = (colors.whitesmoke, colors.black, False)

        cx = x0
        for i, col in enumerate(cols):
            val = "" if pd.isna(row[col]) else str(row[col])

            if get_cell_style:
                cell_style = get_cell_style(row, col)
            else:
                cell_style = None

            if cell_style:
                cell_bg, cell_fg, cell_bold = cell_style
            else:
                cell_bg, cell_fg, cell_bold = bg, fg, is_bold

            c.setFillColor(cell_bg)
            c.rect(cx, y - row_h, col_ws[i], row_h, stroke=1, fill=1)

            c.setFillColor(cell_fg)
            font_name = "Helvetica-Bold" if cell_bold else body_font[0]
            font_size = body_font[1]
            if get_cell_fontsize:
                fs = get_cell_fontsize(row, col)
                if fs:
                    font_size = fs
            c.setFont(font_name, font_size)

            align = None
            if get_cell_align:
                align = get_cell_align(row, col)

            if align == "right" or (align is None and col in align_right_cols):
                c.drawRightString(cx + col_ws[i] - 3 * mm, y - row_h + 3.0 * mm, val)
            elif align == "center" or (align is None and i != 0):
                c.drawCentredString(cx + col_ws[i] / 2, y - row_h + 3.0 * mm, val)
            elif align == "left" or (align is None and i == 0):
                c.drawString(cx + 3 * mm, y - row_h + 3.0 * mm, val)

            cx += col_ws[i]

        y -= row_h

    return y


# ---------------- Chart (PNG bytes) : PRODUIT ----------------
def build_product_chart_png(tab_produit: pd.DataFrame, title: str) -> bytes:
    dfc = tab_produit.copy()
    if dfc is None or dfc.empty:
        return b""

    dfc = dfc[dfc["PRODUIT"].astype(str).str.lower() != "total"].copy()
    dfc["Prime TTC (FCFA)"] = pd.to_numeric(dfc["Prime TTC (FCFA)"], errors="coerce").fillna(0)
    dfc["% Prime TTC"] = pd.to_numeric(dfc["% Prime TTC"], errors="coerce").fillna(0)

    dfc = dfc.sort_values("Prime TTC (FCFA)", ascending=False)

    produits = dfc["PRODUIT"].astype(str).tolist()
    ca_fcfa = dfc["Prime TTC (FCFA)"].tolist()
    pct = dfc["% Prime TTC"].tolist()

    ca_m = [v / 1_000_000 for v in ca_fcfa]
    total_m = (sum(ca_fcfa) / 1_000_000) if sum(ca_fcfa) else 0

    fig, ax = plt.subplots(figsize=(11.5, 4.0), dpi=130)
    fig.patch.set_facecolor("#fbf3d6")
    ax.set_facecolor("#fbf3d6")

    bars = ax.bar(produits, ca_m, edgecolor="black", linewidth=1.2)

    if total_m > 0:
        ax.axhline(total_m, linestyle="--", linewidth=1.6, label="Total CA (tous produits)")

    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_ylabel("Chiffre d’Affaires (Millions FCFA)")
    ax.set_xlabel("Produit")

    top = max(ca_m) if ca_m else 0
    for b, v_fcfa, p in zip(bars, ca_fcfa, pct):
        ax.text(
            b.get_x() + b.get_width() / 2,
            b.get_height() + (0.03 * top if top > 0 else 0.1),
            f"{v_fcfa/1_000_000:.0f} M FCFA\n({p*100:.1f}%)",
            ha="center",
            va="bottom",
            fontsize=9.5,
            fontweight="bold",
        )

    ax.tick_params(axis="x", rotation=35)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.legend(loc="upper right", fontsize=9)

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ---------------- Chart (PNG bytes) : MENSUEL (2 charts) ----------------
def build_monthly_report_charts_png(tab_month: pd.DataFrame, end_date_str: str) -> tuple[bytes, bytes]:
    if tab_month is None or tab_month.empty:
        return b"", b""

    dfm = tab_month.copy()
    dfm = dfm[~dfm["Mois"].astype(str).str.lower().str.startswith("total")].copy()

    for col in ["CA par Mois", "Objectif Mensuel", "Ratio Objectif Mensuel"]:
        dfm[col] = pd.to_numeric(dfm[col], errors="coerce").fillna(0)

    order_fr = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
    months_present = [m for m in order_fr if m in dfm["Mois"].astype(str).unique().tolist()]
    order_used = months_present or order_fr
    dfm["Mois"] = pd.Categorical(dfm["Mois"], categories=order_used, ordered=True)
    dfm = dfm.sort_values("Mois")

    ca_m = (dfm["CA par Mois"] / 1_000_000).tolist()
    obj_m = (dfm["Objectif Mensuel"] / 1_000_000).tolist()
    ratio_pct = (dfm["Ratio Objectif Mensuel"] * 100).tolist()
    mois = dfm["Mois"].astype(str).tolist()

    # Chart 1
    fig1, ax1 = plt.subplots(figsize=(12.5, 4.0), dpi=130)
    ax1.bar(mois, ca_m, edgecolor="black", linewidth=1.0)
    ax1.plot(mois, obj_m, linestyle="--", marker="o", linewidth=1.8, label="Objectif mensuel")

    if order_used:
        title_range = f"{order_used[0][:3]} → {order_used[-1][:3]}"
    else:
        title_range = "Mois disponibles"
    ax1.set_title(f"Chiffre d’affaires mensuel — {title_range}", fontsize=15, fontweight="bold")
    ax1.set_ylabel("Montant (Millions FCFA)")
    ax1.grid(axis="y", linestyle=":", alpha=0.35)
    ax1.legend(loc="upper right", fontsize=10)
    ax1.tick_params(axis="x", rotation=20)

    ymax = max(max(ca_m), max(obj_m)) if ca_m and obj_m else 0
    for i, (y, r) in enumerate(zip(ca_m, ratio_pct)):
        ax1.text(i, y + (0.03 * ymax if ymax > 0 else 0.2), f"{r:.1f}%", ha="center", va="bottom", fontsize=10, fontweight="bold")

    buf1 = BytesIO()
    plt.tight_layout()
    fig1.savefig(buf1, format="png", bbox_inches="tight")
    plt.close(fig1)
    buf1.seek(0)

    # Chart 2
    fig2, ax2 = plt.subplots(figsize=(12.5, 4.0), dpi=130)
    bars = ax2.bar(mois, ca_m, edgecolor="black", linewidth=1.0)
    ax2.set_title(f"Évolution mensuelle du chiffre d’affaires (Janvier au {end_date_str})", fontsize=15, fontweight="bold")
    ax2.set_ylabel("Chiffre d'affaires (Millions FCFA)")
    ax2.grid(axis="y", linestyle=":", alpha=0.35)
    ax2.tick_params(axis="x", rotation=20)

    ymax2 = max(ca_m) if ca_m else 0
    for b, v in zip(bars, dfm["CA par Mois"].tolist()):
        ax2.text(
            b.get_x() + b.get_width() / 2,
            b.get_height() + (0.03 * ymax2 if ymax2 > 0 else 0.2),
            fmt_int(v),
            ha="center",
            va="bottom",
            fontsize=9.5,
            fontweight="bold",
        )

    buf2 = BytesIO()
    plt.tight_layout()
    fig2.savefig(buf2, format="png", bbox_inches="tight")
    plt.close(fig2)
    buf2.seek(0)

    return buf1.read(), buf2.read()


# ---------------- Chart (PNG bytes) : CA moyen journalier (jour de semaine) ----------------
def build_weekday_avg_chart_png(
    tab_week: pd.DataFrame,
    ref_month: str = "Avril",
    max_compare: int = 3,
) -> bytes:
    """
    Graphique lignes: CA moyen journalier par jour de la semaine.
    - tab_week format attendu: colonnes = ["Mois","dimanche","lundi","mardi","mercredi","jeudi","vendredi","samedi","Total"]
    - Affiche: ref_month (réf.) + les derniers mois disponibles (max_compare)
    """
    if tab_week is None or tab_week.empty:
        return b""

    df = tab_week.copy()
    if "Mois" not in df.columns:
        return b""

    # enlever Total ligne
    df = df[~df["Mois"].astype(str).str.strip().str.lower().eq("total")].copy()

    days = ["dimanche","lundi","mardi","mercredi","jeudi","vendredi","samedi"]
    # garder colonnes existantes
    days = [d for d in days if d in df.columns]
    if not days:
        return b""

    # mapping ordre mois
    order_fr = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
    month_to_idx = {m.lower(): i for i, m in enumerate(order_fr)}

    df["_m"] = df["Mois"].astype(str).str.strip()
    df["_mi"] = df["_m"].str.lower().map(month_to_idx).fillna(999).astype(int)

    # numeric
    for d in days:
        df[d] = pd.to_numeric(df[d], errors="coerce").fillna(0)

    # choix des mois à tracer
    ref_key = ref_month.strip().lower()
    has_ref = (df["_m"].str.lower() == ref_key).any()

    df_sorted = df.sort_values("_mi").copy()
    others = df_sorted[df_sorted["_m"].str.lower() != ref_key].copy()

    # derniers mois dispo (en gardant l'ordre chronologique)
    others_tail = others[others["_mi"] != 999].tail(max_compare)
    months_plot = []
    if has_ref:
        months_plot.append(ref_key)
    months_plot += others_tail["_m"].str.lower().tolist()

    if not months_plot:
        # fallback: derniers 3 mois
        months_plot = df_sorted.tail(max_compare)["_m"].str.lower().tolist()

    # dataset final
    dfp = df_sorted[df_sorted["_m"].str.lower().isin(months_plot)].copy()

    # labels title
    months_lbl = dfp["_m"].tolist()
    if has_ref:
        other_lbl = [m for m in months_lbl if m.lower() != ref_key]
        title = f"CA moyen journalier par jour de la semaine - {ref_month} (réf.) vs " + " & ".join(other_lbl)
    else:
        title = "CA moyen journalier par jour de la semaine"

    fig, ax = plt.subplots(figsize=(12.5, 4.0), dpi=130)

    x = list(range(len(days)))
    ax.set_xticks(x)
    ax.set_xticklabels(days, rotation=45, ha="right")

    for _, r in dfp.iterrows():
        label = str(r["_m"])
        if has_ref and label.strip().lower() == ref_key:
            label = f"{label} (référence)"
            ax.plot(x, [r[d] for d in days], marker="o", linewidth=2.4, label=label)
        else:
            ax.plot(x, [r[d] for d in days], marker="o", linewidth=2.0, label=label)

    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_xlabel("Jour de la semaine")
    ax.set_ylabel("CA moyen journalier (FCFA)")
    ax.grid(True, linestyle=":", alpha=0.35)
    ax.legend(loc="upper left", fontsize=9)

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def build_top_bottom_table(tab_month: pd.DataFrame, n: int = 3) -> pd.DataFrame:
    if tab_month is None or tab_month.empty:
        return pd.DataFrame(columns=["Type", "Mois", "CA (FCFA)", "Ratio"])

    dfm = tab_month.copy()
    dfm = dfm[~dfm["Mois"].astype(str).str.lower().str.startswith("total")].copy()

    dfm["CA par Mois"] = pd.to_numeric(dfm["CA par Mois"], errors="coerce").fillna(0)
    dfm["Ratio Objectif Mensuel"] = pd.to_numeric(dfm["Ratio Objectif Mensuel"], errors="coerce").fillna(0)

    top = dfm.sort_values("CA par Mois", ascending=False).head(n).copy()
    bot = dfm.sort_values("CA par Mois", ascending=True).head(n).copy()

    top["Type"] = "TOP"
    bot["Type"] = "BOTTOM"

    out = pd.concat([top, bot], ignore_index=True)
    out = out[["Type", "Mois", "CA par Mois", "Ratio Objectif Mensuel"]].rename(
        columns={"CA par Mois": "CA (FCFA)", "Ratio Objectif Mensuel": "Ratio"}
    )

    out["CA (FCFA)"] = out["CA (FCFA)"].map(fmt_int)
    out["Ratio"] = out["Ratio"].map(lambda x: safe_pct(x, 1))
    return out


# =========================
# PAGE 6: COMPARATIFS HELPERS
# =========================
def _format_comp_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Attendu (depuis ton app):
    colonnes: ["Indicateur","Période actuelle","Période référence","Δ (valeur)","Δ (%)"]
    Valeurs numériques -> formatées proprement pour PDF.
    """
    cols = ["Indicateur", "Période actuelle", "Période référence", "Δ (valeur)", "Δ (%)"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    out = df.copy()

    # normaliser colonnes provenant du Streamlit (semaine)
    rename_map = {
        "Semaine courante": "Période actuelle",
        "Semaine précédente": "Période référence",
        "Écart": "Δ (valeur)",
        "Variation %": "Δ (%)",
        "Variation % ": "Δ (%)",
    }
    for src, dst in rename_map.items():
        if src in out.columns and dst not in out.columns:
            out = out.rename(columns={src: dst})

    # sécuriser colonnes
    for c in cols:
        if c not in out.columns:
            out[c] = 0

    for c in ["Période actuelle", "Période référence", "Δ (valeur)"]:
        out[c] = (
            pd.to_numeric(out[c], errors="coerce")
            .fillna(0)
            .map(fmt_int)
        )

    out["Δ (%)"] = pd.to_numeric(out["Δ (%)"], errors="coerce").fillna(0).map(lambda x: safe_pct(x, 1))
    out["Indicateur"] = out["Indicateur"].astype(str)

    return out[cols]


def _style_comp_row(row):
    # Mise en forme simple : lignes alternées + surlignage pool/hors pool
    ind = str(row.get("Indicateur", "")).strip().lower()

    if ind in ["pool", "hors pool", "ca (prime ttc)"]:
        return (colors.HexColor("#FFF1C6"), colors.black, True)  # jaune clair + bold
    return (colors.whitesmoke, colors.black, False)


def build_week_compare_chart_png(df_in: pd.DataFrame) -> bytes:
    if df_in is None or df_in.empty:
        return b""

    dfc = df_in.copy()
    rename_map = {
        "Semaine courante": "Période actuelle",
        "Semaine précédente": "Période référence",
        "Écart": "Δ (valeur)",
        "Variation %": "Δ (%)",
    }
    for src, dst in rename_map.items():
        if src in dfc.columns and dst not in dfc.columns:
            dfc = dfc.rename(columns={src: dst})

    if "Indicateur" not in dfc.columns:
        return b""

    dfc = dfc[dfc["Indicateur"].isin(["Nombre de contrats", "Renouvellements", "Nouvelles affaires"])].copy()
    if dfc.empty:
        return b""

    for col in ["Période actuelle", "Période référence"]:
        if col in dfc.columns:
            dfc[col] = (
                dfc[col]
                .astype(str)
                .str.replace(" ", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
            dfc[col] = pd.to_numeric(dfc[col], errors="coerce").fillna(0)
        else:
            dfc[col] = 0

    labels = dfc["Indicateur"].tolist()
    cur = dfc["Période actuelle"].tolist()
    prev = dfc["Période référence"].tolist()

    x = range(len(labels))
    fig, ax = plt.subplots(figsize=(11.8, 3.8), dpi=130)
    ax.bar([i - 0.2 for i in x], cur, width=0.4, label="Semaine courante", color="#0B63CE")
    ax.bar([i + 0.2 for i in x], prev, width=0.4, label="Semaine précédente", color="#86C5FF")

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=0)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.legend(loc="upper right", fontsize=9)
    ax.set_title("Comparatif semaines", fontsize=12, fontweight="bold")

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def build_month_week_chart_png(df_in: pd.DataFrame, col_cur: str, col_prev: str, title: str) -> bytes:
    if df_in is None or df_in.empty:
        return b""
    if col_cur not in df_in.columns or col_prev not in df_in.columns:
        return b""

    dfc = df_in.copy()
    labels = dfc["Semaine"].astype(str).tolist()
    cur = pd.to_numeric(dfc[col_cur], errors="coerce").fillna(0).tolist()
    prev = pd.to_numeric(dfc[col_prev], errors="coerce").fillna(0).tolist()

    x = range(len(labels))
    fig, ax = plt.subplots(figsize=(11.8, 3.8), dpi=130)
    ax.bar([i - 0.2 for i in x], cur, width=0.4, label="Mois courant", color="#0B63CE")
    ax.bar([i + 0.2 for i in x], prev, width=0.4, label="Mois précédent", color="#86C5FF")

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=0)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.legend(loc="upper right", fontsize=9)
    ax.set_title(title, fontsize=12, fontweight="bold")

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ---------------- PDF EXPORT (6 pages) ----------------
def export_pdf_dashboard(
    k: dict,
    segment: str,
    intermediaire: str,
    start_date: str,
    end_date: str,
    objectif_annuel: float,
    tab_produit: pd.DataFrame,
    tab_month: pd.DataFrame,
    tab_week: pd.DataFrame,  # ✅ CA moyen journalier
    avg_contracts_daily: float | None = None,
    avg_ca_daily: float | None = None,
    ca_1_month: float | None = None,
    nb_1_month: int | None = None,

    # ✅ PAGE 6 (comparatifs)
    comp_week: pd.DataFrame | None = None,
    week_n_label: str = "",
    week_n1_label: str = "",
    comp_ref: pd.DataFrame | None = None,
    ref_mode_label: str = "",
    ref_label: str = "",
    comp_m3: pd.DataFrame | None = None,
    m3_title: str = "",
    comp_month_week: pd.DataFrame | None = None,
    month_week_label: str = "",
    comp_month_week_y: pd.DataFrame | None = None,
    month_week_label_y: str = "",
    comp_period_month: pd.DataFrame | None = None,
    comp_period_label: str = "",
    comp_pool_period_year: pd.DataFrame | None = None,
    comp_pool_period_label: str = "",

    # ✅ TOP 10 AGENTS
    top_agents_ca: pd.DataFrame | None = None,
    top_agents_cnt: pd.DataFrame | None = None,
    top_pool_ca: pd.DataFrame | None = None,
    top_pool_cnt: pd.DataFrame | None = None,
    one_month_feb_table: pd.DataFrame | None = None,
    one_month_feb_label: str = "",
) -> bytes:
    buff = BytesIO()
    c = canvas.Canvas(buff, pagesize=landscape(A4))
    W, H = landscape(A4)

    # ================= PAGE 1 : dashboard =================
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)

    banner_h = 18 * mm
    c.setFillColorRGB(0.85, 0.62, 0.53)
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.5)
    c.rect(12 * mm, H - banner_h - 12 * mm, W - 24 * mm, banner_h, stroke=1, fill=1)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(W / 2, H - banner_h - 6.5 * mm, "PLATEFORME AUTOMOBILE")

    c.setFont("Helvetica", 10.5)
    y_info = H - banner_h - 22 * mm
    c.drawString(14 * mm, y_info, f"Segment : {segment}")
    c.drawString(14 * mm, y_info - 6 * mm, f"Intermédiaire : {intermediaire}")
    c.drawString(14 * mm, y_info - 12 * mm, f"Période : {start_date} → {end_date}")
    c.drawRightString(W - 14 * mm, y_info - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

    margin = 12 * mm
    gap = 8 * mm
    card_h = 26 * mm

    y1 = H - banner_h - 62 * mm
    w4 = (W - 2 * margin - 3 * gap) / 4

    _pdf_card(c, margin + 0 * (w4 + gap), y1, w4, card_h, fmt_int(k.get("taxes", 0)), "Taxes",
              colors.HexColor("#d6793f"), colors.white, colors.black, 2)
    ca_evol_pdf = k.get("ca_evol_text", None)
    ca_evol_up = bool(k.get("ca_evol_up", False))
    ca_evol_down = bool(k.get("ca_evol_down", False))
    ca_evol_color = colors.HexColor("#0b8f3a") if ca_evol_up else (colors.HexColor("#c62828") if ca_evol_down else colors.HexColor("#444444"))
    _pdf_card(
        c,
        margin + 1 * (w4 + gap),
        y1,
        w4,
        card_h,
        fmt_int(k.get("ca", 0)),
        "Chiffre d'affaires",
        colors.HexColor("#2e6bd9"),
        colors.HexColor("#d7d7d7"),
        colors.HexColor("#d6793f"),
        3,
        pct=ca_evol_pdf,
        pct_color=ca_evol_color,
    )
    _pdf_card(c, margin + 2 * (w4 + gap), y1, w4, card_h, fmt_int(k.get("nb_agents_fideles", 0)), "Agents fidèles",
              colors.HexColor("#d6793f"), colors.white, colors.black, 2)
    _pdf_card(c, margin + 3 * (w4 + gap), y1, w4, card_h, fmt_int(k.get("nb_agents_actifs", 0)), "Agents actifs",
              colors.HexColor("#d6793f"), colors.white, colors.black, 2)

    y2 = y1 - card_h - 12 * mm
    w3 = (W - 2 * margin - 2 * gap) / 3

    pool_value_pdf = fmt_int(k.get("montant_pool", 0))
    pool_pct_pdf = safe_pct(k.get("pct_pool", 0), 1)
    if "montant_reverser_pool" in k:
        pool_value_pdf = f"{fmt_int(k.get('montant_pool', 0))} | {fmt_int(k.get('montant_reverser_pool', 0))}"
    if "pool_pct_text" in k:
        pool_pct_pdf = str(k.get("pool_pct_text"))
    elif "pct_reverser_pool_ca" in k:
        pool_pct_pdf = f"{safe_pct(k.get('pct_pool', 0), 1)} | {safe_pct(k.get('pct_reverser_pool_ca', 0), 1)}"
    pool_evol_up = bool(k.get("pool_evol_up", False))
    pool_evol_down = bool(k.get("pool_evol_down", False))
    pool_pct_color = colors.HexColor("#0b8f3a") if pool_evol_up else (colors.HexColor("#c62828") if pool_evol_down else colors.HexColor("#444444"))
    _pdf_card(
        c,
        margin,
        y2,
        w3,
        card_h,
        pool_value_pdf,
        "Montant POOL TPV | Prime de cession",
        colors.HexColor("#2e6bd9"),
        colors.HexColor("#d7d7d7"),
        colors.HexColor("#d6793f"),
        3,
        pct=pool_pct_pdf,
        pct_color=pool_pct_color,
    )
    _pdf_ratio_card(c, margin + (w3 + gap), y2 - 2 * mm, w3, card_h + 4 * mm, safe_pct(k.get("ratio_obj", 0), 1))
    hors_pct_pdf = str(k.get("hors_pool_pct_text")) if "hors_pool_pct_text" in k else safe_pct(k.get("pct_hors_pool", 0), 1)
    hors_evol_up = bool(k.get("hors_evol_up", False))
    hors_evol_down = bool(k.get("hors_evol_down", False))
    hors_pct_color = colors.HexColor("#0b8f3a") if hors_evol_up else (colors.HexColor("#c62828") if hors_evol_down else colors.HexColor("#444444"))
    _pdf_card(c, margin + 2 * (w3 + gap), y2, w3, card_h, fmt_int(k.get("montant_hors_pool", 0)), "Montant hors pool",
              colors.HexColor("#2e6bd9"), colors.HexColor("#d7d7d7"), colors.HexColor("#d6793f"), 3,
              pct=hors_pct_pdf, pct_color=hors_pct_color)

    y3 = y2 - card_h - 14 * mm
    ren_value_pdf = fmt_int(k.get("renouvellements", 0))
    ren_label_pdf = "Renouvellements"
    ren_pct_pdf = safe_pct(k.get("pct_ren", 0), 1)
    if "ren_non_moto" in k:
        ren_value_pdf = f"{fmt_int(k.get('renouvellements', 0))} | {fmt_int(k.get('ren_non_moto', 0))}"
        ren_label_pdf = "Renouvellement | Hors motos"
        ren_pct_pdf = f"{safe_pct(k.get('pct_ren', 0), 1)} | {safe_pct(k.get('pct_ren_non_moto', 0), 1)}"
    _pdf_card(
        c,
        margin,
        y3,
        w3,
        card_h,
        ren_value_pdf,
        ren_label_pdf,
        colors.HexColor("#1b1b1b"),
        colors.white,
        colors.black,
        2,
        pct=ren_pct_pdf,
    )

    contrats_value_pdf = fmt_int(k.get("souscriptions", 0))
    contrats_label_pdf = "Souscriptions"
    if "nb_contrats_sans_moto" in k:
        contrats_value_pdf = f"{fmt_int(k.get('souscriptions', 0))} | {fmt_int(k.get('nb_contrats_sans_moto', 0))}"
        contrats_label_pdf = "Souscriptions | Hors motos"
    _pdf_card(
        c,
        margin + (w3 + gap),
        y3,
        w3,
        card_h,
        contrats_value_pdf,
        contrats_label_pdf,
        colors.HexColor("#1b1b1b"),
        colors.white,
        colors.black,
        2,
    )

    new_value_pdf = fmt_int(k.get("nouvelles_affaires", 0))
    new_label_pdf = "Nouvelles affaires"
    new_pct_pdf = safe_pct(k.get("pct_new", 0), 1)
    if "new_non_moto" in k:
        new_value_pdf = f"{fmt_int(k.get('nouvelles_affaires', 0))} | {fmt_int(k.get('new_non_moto', 0))}"
        new_label_pdf = "Nouvelles affaires | Hors motos"
        new_pct_pdf = f"{safe_pct(k.get('pct_new', 0), 1)} | {safe_pct(k.get('pct_new_non_moto', 0), 1)}"
    _pdf_card(
        c,
        margin + 2 * (w3 + gap),
        y3,
        w3,
        card_h,
        new_value_pdf,
        new_label_pdf,
        colors.HexColor("#1b1b1b"),
        colors.white,
        colors.black,
        2,
        pct=new_pct_pdf,
    )

    if (
        avg_contracts_daily is not None
        or avg_ca_daily is not None
        or ca_1_month is not None
        or nb_1_month is not None
    ):
        card_h_small = 18 * mm
        y4 = y3 - card_h_small - 10 * mm
        w2 = (W - 2 * margin - gap) / 2
        w3_small = (W - 2 * margin - 2 * gap) / 3
        w4_small = (W - 2 * margin - 3 * gap) / 4
        avg_contracts_val = 0 if avg_contracts_daily is None else avg_contracts_daily
        avg_ca_val = 0 if avg_ca_daily is None else avg_ca_daily
        ca_1_month_val = 0 if ca_1_month is None else ca_1_month
        nb_1_month_val = 0 if nb_1_month is None else nb_1_month
        if ca_1_month is None and nb_1_month is None:
            _pdf_card(
                c,
                margin,
                y4,
                w2,
                card_h_small,
                f"{avg_contracts_val:,.1f}".replace(",", " "),
                "Contrats / jour (moy.)",
                colors.HexColor("#1b1b1b"),
                colors.white,
                colors.black,
                2,
            )
            _pdf_card(
                c,
                margin + (w2 + gap),
                y4,
                w2,
                card_h_small,
                fmt_int(avg_ca_val),
                "CA / jour (moy.)",
                colors.HexColor("#1b1b1b"),
                colors.white,
                colors.black,
                2,
            )
        else:
            if nb_1_month is None or ca_1_month is None:
                _pdf_card(
                    c,
                    margin,
                    y4,
                    w3_small,
                    card_h_small,
                    f"{avg_contracts_val:,.1f}".replace(",", " "),
                    "Contrats / jour (moy.)",
                    colors.HexColor("#1b1b1b"),
                    colors.white,
                    colors.black,
                    2,
                )
                _pdf_card(
                    c,
                    margin + (w3_small + gap),
                    y4,
                    w3_small,
                    card_h_small,
                    fmt_int(avg_ca_val),
                    "CA / jour (moy.)",
                    colors.HexColor("#1b1b1b"),
                    colors.white,
                    colors.black,
                    2,
                )
                if ca_1_month is not None:
                    _pdf_card(
                        c,
                        margin + 2 * (w3_small + gap),
                        y4,
                        w3_small,
                        card_h_small,
                        fmt_int(ca_1_month_val),
                        "Montant contrats 1 mois",
                        colors.HexColor("#1b1b1b"),
                        colors.white,
                        colors.black,
                        2,
                    )
                else:
                    _pdf_card(
                        c,
                        margin + 2 * (w3_small + gap),
                        y4,
                        w3_small,
                        card_h_small,
                        fmt_int(nb_1_month_val),
                        "Contrats 1 mois",
                        colors.HexColor("#1b1b1b"),
                        colors.white,
                        colors.black,
                        2,
                    )
            else:
                _pdf_card(
                    c,
                    margin,
                    y4,
                    w4_small,
                    card_h_small,
                    f"{avg_contracts_val:,.1f}".replace(",", " "),
                    "Contrats / jour (moy.)",
                    colors.HexColor("#1b1b1b"),
                    colors.white,
                    colors.black,
                    2,
                )
                _pdf_card(
                    c,
                    margin + (w4_small + gap),
                    y4,
                    w4_small,
                    card_h_small,
                    fmt_int(avg_ca_val),
                    "CA / jour (moy.)",
                    colors.HexColor("#1b1b1b"),
                    colors.white,
                    colors.black,
                    2,
                )
                _pdf_card(
                    c,
                    margin + 2 * (w4_small + gap),
                    y4,
                    w4_small,
                    card_h_small,
                    fmt_int(nb_1_month_val),
                    "Contrats 1 mois",
                    colors.HexColor("#1b1b1b"),
                    colors.white,
                    colors.black,
                    2,
                    pct=(str(k.get("one_month_count_pct_text")) if "one_month_count_pct_text" in k else safe_pct(k.get("pct_1_month_non_moto", 0), 1)),
                    pct_color=(
                        colors.HexColor("#0b8f3a")
                        if bool(k.get("one_month_count_evol_up", False))
                        else (
                            colors.HexColor("#c62828")
                            if bool(k.get("one_month_count_evol_down", False))
                            else colors.HexColor("#444444")
                        )
                    ),
                    pct_align="right",
                    pct_y_offset=4,
                    pct_x_offset=-2,
                    pct_font_max=8.5,
                )
                _pdf_card(
                    c,
                    margin + 3 * (w4_small + gap),
                    y4,
                    w4_small,
                    card_h_small,
                    fmt_int(ca_1_month_val),
                    "Montant 1 mois",
                    colors.HexColor("#1b1b1b"),
                    colors.white,
                    colors.black,
                    2,
                    pct=(str(k.get("one_month_pct_text")) if "one_month_pct_text" in k else None),
                    pct_color=(
                        colors.HexColor("#0b8f3a")
                        if bool(k.get("one_month_evol_up", False))
                        else (
                            colors.HexColor("#c62828")
                            if bool(k.get("one_month_evol_down", False))
                            else colors.HexColor("#444444")
                        )
                    ),
                    pct_align="right",
                    pct_y_offset=4,
                    pct_x_offset=-2,
                    pct_font_max=8.5,
                )

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawString(margin, 10 * mm, f"Objectif annuel : {fmt_int(objectif_annuel)}")

    c.showPage()

    # ================= PAGE 1 bis : CONTRATS 1 MOIS (ECHEANCE FEVRIER) =================
    if one_month_feb_table is not None and not one_month_feb_table.empty:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(W / 2, H - 16 * mm, "CONTRATS DUREE 1 MOIS - ECHEANCE FEVRIER")
        if one_month_feb_label:
            c.setFont("Helvetica", 11)
            c.drawString(14 * mm, H - 24 * mm, one_month_feb_label[:160])

        dff = one_month_feb_table.copy()
        for col in [
            "Souscriptions 1 mois (échéance février)",
            "Renouvelés en février",
            "Montant souscriptions (FCFA)",
            "Montant renouvelés (FCFA)",
        ]:
            if col in dff.columns:
                dff[col] = pd.to_numeric(dff[col], errors="coerce").fillna(0).map(fmt_int)
        for col in ["Part réalisée", "% renouvellement"]:
            if col in dff.columns:
                dff[col] = pd.to_numeric(dff[col], errors="coerce").fillna(0).map(lambda x: safe_pct(x, 1))

        if "PRODUIT" in dff.columns:
            cols_keep = [
                "PRODUIT",
                "Souscriptions 1 mois (échéance février)",
                "Renouvelés en février",
                "Montant souscriptions (FCFA)",
                "Montant renouvelés (FCFA)",
                "Part réalisée",
                "% renouvellement",
            ]
            dff = dff[[c for c in cols_keep if c in dff.columns]]

        # En-tetes PDF raccourcis pour eviter les chevauchements
        dff = dff.rename(
            columns={
                "Souscriptions 1 mois (échéance février)": "Sousc. 1 mois (ech. fev.)",
                "Renouvelés en février": "Renouv. fev.",
                "Montant souscriptions (FCFA)": "Montant sousc. (FCFA)",
                "Montant renouvelés (FCFA)": "Montant renouv. (FCFA)",
                "Part réalisée": "Part realisee",
                "% renouvellement": "% renouvellement",
            }
        )

        def style_one_month(row):
            is_total = str(row.iloc[0]).strip().lower() in {"grand total", "total"}
            if is_total:
                return (colors.HexColor("#1F4E79"), colors.white, True)
            return (colors.whitesmoke, colors.black, False)

        _pdf_table_generic(
            c,
            dff,
            x0=14 * mm,
            start_y=H - 34 * mm,
            col_ws=[40 * mm, 46 * mm, 28 * mm, 40 * mm, 40 * mm, 20 * mm, 18 * mm],
            row_h=9 * mm,
            get_row_style=style_one_month,
            align_right_cols={c for c in dff.columns if c != "PRODUIT"},
            header_font=("Helvetica-Bold", 7.2),
            body_font=("Helvetica", 8.0),
        )

        c.showPage()

    # ================= PAGE 2 : PRODUITS =================
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(W / 2, H - 18 * mm, "RÉPARTITION PAR PRODUIT")

    chart_png = build_product_chart_png(tab_produit, "Répartition du Chiffre d’Affaires par Produit")
    y_top = H - 30 * mm

    if chart_png:
        img_reader = ImageReader(BytesIO(chart_png))
        chart_x = 14 * mm
        chart_w = W - 28 * mm
        chart_h = 72 * mm
        chart_y = y_top - chart_h
        c.drawImage(img_reader, chart_x, chart_y, width=chart_w, height=chart_h, mask="auto")
        start_table_y = chart_y - 6 * mm
    else:
        start_table_y = y_top

    dfp = tab_produit.copy()
    if dfp is None or dfp.empty:
        dfp = pd.DataFrame(columns=["PRODUIT", "Nombre de contrats", "% Contrats", "Prime TTC (FCFA)", "% Prime TTC"])
    else:
        dfp["Nombre de contrats"] = pd.to_numeric(dfp["Nombre de contrats"], errors="coerce").fillna(0).map(fmt_int)
        dfp["Prime TTC (FCFA)"] = pd.to_numeric(dfp["Prime TTC (FCFA)"], errors="coerce").fillna(0).map(fmt_int)
        dfp["% Contrats"] = pd.to_numeric(dfp["% Contrats"], errors="coerce").fillna(0).map(lambda x: f"{x*100:.1f}%")
        dfp["% Prime TTC"] = pd.to_numeric(dfp["% Prime TTC"], errors="coerce").fillna(0).map(lambda x: f"{x*100:.1f}%")
        dfp = dfp[["PRODUIT", "Nombre de contrats", "% Contrats", "Prime TTC (FCFA)", "% Prime TTC"]]

    def style_prod(row):
        is_total = str(row["PRODUIT"]).strip().lower() == "total"
        if is_total:
            return (colors.HexColor("#1F4E79"), colors.white, True)
        return (colors.whitesmoke, colors.black, False)

    col_ws = [48 * mm, 48 * mm, 30 * mm, 58 * mm, 30 * mm]
    _pdf_table_generic(
        c,
        dfp,
        x0=14 * mm,
        start_y=start_table_y,
        col_ws=col_ws,
        row_h=9 * mm,
        get_row_style=style_prod,
        align_right_cols={"Nombre de contrats", "% Contrats", "Prime TTC (FCFA)", "% Prime TTC"},
    )

    c.showPage()

    # ================= PAGE 3 : 2 GRAPHIQUES MENSUELS =================
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(W / 2, H - 18 * mm, "CHIFFRE D’AFFAIRES MENSUEL")

    ch1, ch2 = build_monthly_report_charts_png(tab_month, end_date)

    margin_x = 14 * mm
    gap_y = 6 * mm
    available_h = H - 30 * mm - 14 * mm
    each_h = (available_h - gap_y) / 2

    y_top = H - 26 * mm
    if ch1:
        c.drawImage(ImageReader(BytesIO(ch1)), margin_x, y_top - each_h, width=W - 2 * margin_x, height=each_h, mask="auto")
    if ch2:
        c.drawImage(ImageReader(BytesIO(ch2)), margin_x, y_top - each_h - gap_y - each_h, width=W - 2 * margin_x, height=each_h, mask="auto")

    c.showPage()

    # ================= PAGE 4 : TABLE MENSUEL + TOP / BOTTOM =================
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(W / 2, H - 18 * mm, "TABLEAU MENSUEL + TOP / BOTTOM")

    df_month = tab_month.copy()
    if df_month is None or df_month.empty:
        df_month = pd.DataFrame(columns=["Mois", "CA par Mois", "Objectif Mensuel", "Ratio Objectif Mensuel"])
    else:
        df_month = df_month.copy()
        df_month.insert(0, " ", list(range(len(df_month))))

        df_month["CA par Mois"] = pd.to_numeric(df_month["CA par Mois"], errors="coerce").fillna(0).map(fmt_int)
        df_month["Objectif Mensuel"] = pd.to_numeric(df_month["Objectif Mensuel"], errors="coerce").fillna(0).map(fmt_int)
        df_month["Ratio Objectif Mensuel"] = pd.to_numeric(df_month["Ratio Objectif Mensuel"], errors="coerce").fillna(0).map(lambda x: safe_pct(x, 1))

        df_month = df_month[[" ", "Mois", "CA par Mois", "Objectif Mensuel", "Ratio Objectif Mensuel"]]

    def style_month(row):
        mois = str(row["Mois"]).strip().lower()
        is_total = mois.startswith("total")
        if is_total:
            return (colors.HexColor("#FFF1C6"), colors.black, True)
        return (colors.whitesmoke, colors.black, False)

    df_tb = build_top_bottom_table(tab_month, n=3)

    def style_tb(row):
        typ = str(row["Type"]).strip().upper()
        if typ == "TOP":
            return (colors.HexColor("#DFF2DF"), colors.HexColor("#0B5D1E"), True)   # ✅ vert
        if typ == "BOTTOM":
            return (colors.HexColor("#FBE5E5"), colors.HexColor("#8B0000"), True)
        return (colors.whitesmoke, colors.black, False)

    x_left = 14 * mm
    y_start = H - 30 * mm

    mini_w = 90 * mm
    gap = 8 * mm
    left_w = (W - 28 * mm) - mini_w - gap

    col_ws_month = [12 * mm, 42 * mm, 45 * mm, 45 * mm, 45 * mm]
    scale = left_w / sum(col_ws_month)
    col_ws_month = [w * scale for w in col_ws_month]

    _pdf_table_generic(
        c,
        df_month,
        x0=x_left,
        start_y=y_start,
        col_ws=col_ws_month,
        row_h=8.5 * mm,
        get_row_style=style_month,
        align_right_cols={" ", "CA par Mois", "Objectif Mensuel", "Ratio Objectif Mensuel"},
    )

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x_left + left_w + gap, y_start + 2 * mm, "TOP / BOTTOM (CA mensuel)")

    if df_tb is None or df_tb.empty:
        df_tb = pd.DataFrame(columns=["Type", "Mois", "CA (FCFA)", "Ratio"])

    col_ws_tb = [22 * mm, 28 * mm, 28 * mm, 12 * mm]
    scale2 = mini_w / sum(col_ws_tb)
    col_ws_tb = [w * scale2 for w in col_ws_tb]

    _pdf_table_generic(
        c,
        df_tb,
        x0=x_left + left_w + gap,
        start_y=y_start - 6 * mm,
        col_ws=col_ws_tb,
        row_h=9.0 * mm,
        get_row_style=style_tb,
        align_right_cols={"CA (FCFA)", "Ratio"},
        header_bg=colors.HexColor("#F4D35E"),
    )

    c.showPage()

    # ================= PAGE 5 : CA moyen journalier (graph + tableau) =================
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, W, H, stroke=0, fill=1)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(W / 2, H - 18 * mm, "CA MOYEN JOURNALIER PAR JOUR DE SEMAINE")

    df_week = tab_week.copy() if tab_week is not None else pd.DataFrame()
    df_week_num = pd.DataFrame()

    if df_week.empty:
        df_week = pd.DataFrame(columns=["Mois", "dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "Total"])
        df_week_num = df_week.copy()
    else:
        # version numérique pour chart
        df_week_num = df_week.copy()
        for col in df_week_num.columns:
            if col != "Mois":
                df_week_num[col] = pd.to_numeric(df_week_num[col], errors="coerce").fillna(0)

        # version print pour tableau
        for col in df_week.columns:
            if col != "Mois":
                df_week[col] = pd.to_numeric(df_week[col], errors="coerce").fillna(0).map(fmt_int)

    def style_week(row):
        is_total = str(row["Mois"]).strip().lower() == "total"
        if is_total:
            return (colors.HexColor("#FFF1C6"), colors.black, True)
        return (colors.whitesmoke, colors.black, False)

    # --- Graphique en haut ---
    chart_week = b""
    try:
        if not df_week_num.empty:
            chart_week = build_weekday_avg_chart_png(
                tab_week=df_week_num,
                ref_month="Avril",
                max_compare=3,
            )
    except Exception:
        chart_week = b""

    x0 = 14 * mm
    y_top = H - 28 * mm

    if chart_week:
        chart_h = 70 * mm
        chart_w = W - 28 * mm
        chart_y = y_top - chart_h
        c.drawImage(ImageReader(BytesIO(chart_week)), x0, chart_y, width=chart_w, height=chart_h, mask="auto")
        table_start_y = chart_y - 6 * mm
    else:
        table_start_y = y_top

    # --- Tableau en dessous ---
    col_ws_w = [40*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 24*mm]
    scale_w = (W - 28*mm) / sum(col_ws_w)
    col_ws_w = [w * scale_w for w in col_ws_w]

    _pdf_table_generic(
        c,
        df_week,
        x0=x0,
        start_y=table_start_y,
        col_ws=col_ws_w,
        row_h=8.3 * mm,
        get_row_style=style_week,
        align_right_cols=set([cc for cc in df_week.columns if cc != "Mois"]),
        header_bg=colors.HexColor("#F4D35E"),
    )

    c.showPage()

    # ================= PAGE 6 : ANALYSE COMPARATIVE =================
    # On n'affiche la page que si on a au moins une table
    has_comp = (
        (comp_week is not None and not comp_week.empty) or
        (comp_ref is not None and not comp_ref.empty) or
        (comp_m3 is not None and not comp_m3.empty)
    )

    if has_comp:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 24)
        c.drawCentredString(W / 2, H - 18 * mm, "ANALYSE COMPARATIVE")

        c.setFont("Helvetica", 10.5)
        y_info2 = H - 30 * mm
        c.drawString(14 * mm, y_info2, f"Segment : {segment}")
        c.drawString(14 * mm, y_info2 - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info2 - 12 * mm, f"Période : {start_date} → {end_date}")
        c.drawRightString(W - 14 * mm, y_info2 - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

        x0 = 14 * mm
        y = y_info2 - 22 * mm

        # Table widths (5 cols)
        col_ws_c = [60*mm, 42*mm, 42*mm, 42*mm, 25*mm]
        scale_c = (W - 28*mm) / sum(col_ws_c)
        col_ws_c = [w * scale_c for w in col_ws_c]

        def _draw_section(title: str, subtitle: str, df_in: pd.DataFrame | None, y_start: float) -> float:
            if df_in is None or df_in.empty:
                return y_start

            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 14)
            c.drawString(x0, y_start, title)

            if subtitle:
                c.setFont("Helvetica", 10.5)
                c.drawString(x0, y_start - 6 * mm, subtitle)
                y_table = y_start - 12 * mm
            else:
                y_table = y_start - 6 * mm

            dff = _format_comp_df(df_in)

            # lignes alternées + mise en avant CA/POOL/Hors POOL
            y_after = _pdf_table_generic(
                c,
                dff,
                x0=x0,
                start_y=y_table,
                col_ws=col_ws_c,
                row_h=8.4 * mm,
                get_row_style=_style_comp_row,
                align_right_cols={"Période actuelle", "Période référence", "Δ (valeur)", "Δ (%)"},
                header_bg=colors.HexColor("#F4D35E"),
            )
            return y_after - 10 * mm

        # 1) semaine N vs N-1
        y = _draw_section(
            "1) Semaine N vs Semaine N-1",
            f"N : {week_n_label} | N-1 : {week_n1_label}" if (week_n_label or week_n1_label) else "",
            comp_week,
            y,
        )
        if comp_week is not None and not comp_week.empty:
            chart_week = build_week_compare_chart_png(comp_week)
            if chart_week:
                chart_h = 55 * mm
                chart_w = W - 28 * mm
                chart_y = y - chart_h
                c.drawImage(ImageReader(BytesIO(chart_week)), x0, chart_y, width=chart_w, height=chart_h, mask="auto")
                y = chart_y - 10 * mm

        # 2) période vs référence
        y = _draw_section(
            "2) Période vs Référence",
            f"Référence : {ref_label} ({ref_mode_label})" if (ref_label or ref_mode_label) else "",
            comp_ref,
            y,
        )

        # 3) mois vs moyenne 3 mois
        y = _draw_section(
            "3) Mois en cours vs Moyenne des 3 derniers mois",
            (m3_title or "")[:140],
            comp_m3,
            y,
        )

        c.showPage()

    # ================= PAGE 7 : COMPARATIF PERIODE DU MOIS =================
    if comp_period_month is not None and not comp_period_month.empty:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(W / 2, H - 18 * mm, "COMPARATIF PERIODE DU MOIS (N-1)")

        c.setFont("Helvetica", 10.5)
        y_info_pm = H - 30 * mm
        c.drawString(14 * mm, y_info_pm, f"Segment : {segment}")
        c.drawString(14 * mm, y_info_pm - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info_pm - 12 * mm, f"Période : {start_date} → {end_date}")
        if comp_period_label:
            c.drawString(14 * mm, y_info_pm - 18 * mm, comp_period_label[:150])
        c.drawRightString(W - 14 * mm, y_info_pm - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

        dfpm = comp_period_month.copy()
        cols_pm = ["Indicateur", "Période M", "Période N-1", "Δ M vs N-1"]
        for ccol in cols_pm:
            if ccol not in dfpm.columns:
                dfpm[ccol] = 0
        for ccol in [c for c in cols_pm if c != "Indicateur"]:
            dfpm[ccol] = pd.to_numeric(dfpm[ccol], errors="coerce").fillna(0).map(fmt_int)
        dfpm = dfpm[cols_pm]

        col_ws_pm = [78 * mm, 34 * mm, 34 * mm, 34 * mm]
        scale_pm = (W - 28 * mm) / sum(col_ws_pm)
        col_ws_pm = [w * scale_pm for w in col_ws_pm]

        def _row_style_pm(row):
            return (colors.white if row.name % 2 == 0 else colors.HexColor("#f6f7f9"), colors.black, False)

        def _cell_style_pm(row, col):
            if str(col).startswith("Δ"):
                return (colors.HexColor("#fff3cd"), colors.black, False)
            return None

        _pdf_table_generic(
            c,
            dfpm,
            x0=14 * mm,
            start_y=y_info_pm - 26 * mm,
            col_ws=col_ws_pm,
            row_h=9.0 * mm,
            align_right_cols=set([c for c in cols_pm if c != "Indicateur"]),
            header_bg=colors.HexColor("#F4D35E"),
            header_font=("Helvetica-Bold", 9.0),
            body_font=("Helvetica", 8.8),
            get_row_style=_row_style_pm,
            get_cell_style=_cell_style_pm,
        )

        c.showPage()

    # ================= PAGE 7 bis : COMPARATIF POOL vs HORS POOL (N-1) =================
    if comp_pool_period_year is not None and not comp_pool_period_year.empty:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(W / 2, H - 18 * mm, "COMPARATIF POOL vs HORS POOL (N-1)")

        c.setFont("Helvetica", 10.5)
        y_info_pp = H - 30 * mm
        c.drawString(14 * mm, y_info_pp, f"Segment : {segment}")
        c.drawString(14 * mm, y_info_pp - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info_pp - 12 * mm, f"Période : {start_date} → {end_date}")
        if comp_pool_period_label:
            c.drawString(14 * mm, y_info_pp - 18 * mm, comp_pool_period_label[:150])
        c.drawRightString(W - 14 * mm, y_info_pp - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

        dpp = comp_pool_period_year.copy()
        cols_pp = [
            "Type", "Contrats M", "Contrats N-1", "Δ Contrats",
            "Renouv. M", "Renouv. N-1", "Δ Renouv.",
            "Nouv. aff. M", "Nouv. aff. N-1", "Δ Nouv.",
            "CA M", "CA N-1", "Δ CA",
        ]
        for ccol in cols_pp:
            if ccol not in dpp.columns:
                dpp[ccol] = 0
        for ccol in [c for c in cols_pp if c != "Type"]:
            dpp[ccol] = pd.to_numeric(dpp[ccol], errors="coerce").fillna(0).map(fmt_int)
        dpp = dpp[cols_pp]

        col_ws_pp = [20 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 24 * mm, 24 * mm, 24 * mm]
        scale_pp = (W - 28 * mm) / sum(col_ws_pp)
        col_ws_pp = [w * scale_pp for w in col_ws_pp]

        def _row_style_pp(row):
            is_total = str(row.get("Type", "")).strip().lower() == "total"
            if is_total:
                return (colors.HexColor("#e9ecef"), colors.black, True)
            return (colors.white if row.name % 2 == 0 else colors.HexColor("#f6f7f9"), colors.black, False)

        def _cell_style_pp(row, col):
            if str(col).startswith("Δ"):
                return (colors.HexColor("#fff3cd"), colors.black, False)
            return None

        _pdf_table_generic(
            c,
            dpp,
            x0=14 * mm,
            start_y=y_info_pp - 26 * mm,
            col_ws=col_ws_pp,
            row_h=8.6 * mm,
            align_right_cols=set([c for c in cols_pp if c != "Type"]),
            header_bg=colors.HexColor("#F4D35E"),
            header_font=("Helvetica-Bold", 8.0),
            body_font=("Helvetica", 7.8),
            get_row_style=_row_style_pp,
            get_cell_style=_cell_style_pp,
        )

        c.showPage()

    # ================= PAGE 8 : COMPARATIF SEMAINES DU MOIS =================
    if comp_month_week is not None and not comp_month_week.empty:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 24)
        c.drawCentredString(W / 2, H - 18 * mm, "COMPARATIF SEMAINES DU MOIS")

        c.setFont("Helvetica", 10.5)
        y_info_m = H - 30 * mm
        c.drawString(14 * mm, y_info_m, f"Segment : {segment}")
        c.drawString(14 * mm, y_info_m - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info_m - 12 * mm, f"Période : {start_date} → {end_date}")
        if month_week_label:
            c.drawString(14 * mm, y_info_m - 18 * mm, month_week_label)
        c.drawRightString(W - 14 * mm, y_info_m - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

        dfmw = comp_month_week.copy()
        cols = [
            "Semaine",
            "Contrats M", "Contrats M-1", "Δ Contrats",
            "Renouv. M", "Renouv. M-1", "Δ Renouv.",
            "Nouv. aff. M", "Nouv. aff. M-1", "Δ Nouv.",
            "CA M", "CA M-1", "Δ CA",
        ]
        for ccol in cols:
            if ccol not in dfmw.columns:
                dfmw[ccol] = 0
        for ccol in [c for c in cols if c != "Semaine"]:
            dfmw[ccol] = pd.to_numeric(dfmw[ccol], errors="coerce").fillna(0).map(fmt_int)
        dfmw = dfmw[cols]

        col_ws_mw = [
            24 * mm,
            18 * mm, 18 * mm, 18 * mm,
            18 * mm, 18 * mm, 18 * mm,
            18 * mm, 18 * mm, 18 * mm,
            20 * mm, 20 * mm, 18 * mm,
        ]
        scale_mw = (W - 28 * mm) / sum(col_ws_mw)
        col_ws_mw = [w * scale_mw for w in col_ws_mw]

        y_table = y_info_m - 26 * mm
        def _row_style_mw(row):
            if str(row.get("Semaine", "")) == "Total":
                return (colors.HexColor("#e9ecef"), colors.black, True)
            return (colors.white if row.name % 2 == 0 else colors.HexColor("#f6f7f9"), colors.black, False)

        def _cell_style_mw(row, col):
            if str(col).startswith("Δ") and str(row.get("Semaine", "")) != "Total":
                return (colors.HexColor("#fff3cd"), colors.black, False)
            return None

        y_after = _pdf_table_generic(
            c,
            dfmw,
            x0=14 * mm,
            start_y=y_table,
            col_ws=col_ws_mw,
            row_h=6.8 * mm,
            align_right_cols=set([c for c in cols if c != "Semaine"]),
            header_bg=colors.HexColor("#F4D35E"),
            header_font=("Helvetica-Bold", 7.2),
            body_font=("Helvetica", 7.0),
            get_row_style=_row_style_mw,
            get_cell_style=_cell_style_mw,
        )

        chart_h = 55 * mm
        chart_w = W - 28 * mm
        y_chart_top = y_after - 8 * mm

        ch1 = build_month_week_chart_png(comp_month_week, "Contrats M", "Contrats M-1", "Contrats par semaine (M vs M-1)")
        ch2 = build_month_week_chart_png(comp_month_week, "CA M", "CA M-1", "Chiffre d'affaires par semaine (M vs M-1)")

        if ch1:
            c.drawImage(ImageReader(BytesIO(ch1)), 14 * mm, y_chart_top - chart_h, width=chart_w, height=chart_h, mask="auto")
        if ch2:
            c.drawImage(ImageReader(BytesIO(ch2)), 14 * mm, y_chart_top - chart_h - 6 * mm - chart_h, width=chart_w, height=chart_h, mask="auto")

        c.showPage()

    # ================= PAGE 7 BIS : COMPARATIF SEMAINES (ANNEE PRECEDENTE) =================
    if comp_month_week_y is not None and not comp_month_week_y.empty:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(W / 2, H - 18 * mm, "COMPARATIF SEMAINES (ANNEE PRECEDENTE)")

        c.setFont("Helvetica", 10.5)
        y_info_my = H - 30 * mm
        c.drawString(14 * mm, y_info_my, f"Segment : {segment}")
        c.drawString(14 * mm, y_info_my - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info_my - 12 * mm, f"Période : {start_date} → {end_date}")
        if month_week_label_y:
            c.drawString(14 * mm, y_info_my - 18 * mm, month_week_label_y)
        c.drawRightString(W - 14 * mm, y_info_my - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

        dfmwy = comp_month_week_y.copy()
        cols_y = [
            "Semaine",
            "Contrats M", "Contrats M-1", "Δ Contrats",
            "Renouv. M", "Renouv. M-1", "Δ Renouv.",
            "Nouv. aff. M", "Nouv. aff. M-1", "Δ Nouv.",
            "CA M", "CA M-1", "Δ CA",
        ]
        for ccol in cols_y:
            if ccol not in dfmwy.columns:
                dfmwy[ccol] = 0
        for ccol in [c for c in cols_y if c != "Semaine"]:
            dfmwy[ccol] = pd.to_numeric(dfmwy[ccol], errors="coerce").fillna(0).map(fmt_int)
        dfmwy = dfmwy[cols_y]

        col_ws_my = [
            24 * mm,
            18 * mm, 18 * mm, 18 * mm,
            18 * mm, 18 * mm, 18 * mm,
            18 * mm, 18 * mm, 18 * mm,
            20 * mm, 20 * mm, 18 * mm,
        ]
        scale_my = (W - 28 * mm) / sum(col_ws_my)
        col_ws_my = [w * scale_my for w in col_ws_my]

        y_table_my = y_info_my - 26 * mm
        def _row_style_my(row):
            if str(row.get("Semaine", "")) == "Total":
                return (colors.HexColor("#e9ecef"), colors.black, True)
            return (colors.white if row.name % 2 == 0 else colors.HexColor("#f6f7f9"), colors.black, False)

        def _cell_style_my(row, col):
            if str(col).startswith("Δ") and str(row.get("Semaine", "")) != "Total":
                return (colors.HexColor("#fff3cd"), colors.black, False)
            return None

        _pdf_table_generic(
            c,
            dfmwy,
            x0=14 * mm,
            start_y=y_table_my,
            col_ws=col_ws_my,
            row_h=6.8 * mm,
            align_right_cols=set([c for c in cols_y if c != "Semaine"]),
            header_bg=colors.HexColor("#F4D35E"),
            header_font=("Helvetica-Bold", 7.2),
            body_font=("Helvetica", 7.0),
            get_row_style=_row_style_my,
            get_cell_style=_cell_style_my,
        )

        c.showPage()

        # Page 2 for renouv/new charts
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(W / 2, H - 18 * mm, "COMPARATIF SEMAINES DU MOIS (suite)")

        c.setFont("Helvetica", 10.5)
        y_info_m2 = H - 30 * mm
        c.drawString(14 * mm, y_info_m2, f"Segment : {segment}")
        c.drawString(14 * mm, y_info_m2 - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info_m2 - 12 * mm, f"Période : {start_date} → {end_date}")
        if month_week_label:
            c.drawString(14 * mm, y_info_m2 - 18 * mm, month_week_label)
        c.drawRightString(W - 14 * mm, y_info_m2 - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))

        y_chart_top2 = H - 42 * mm
        ch3 = build_month_week_chart_png(comp_month_week, "Renouv. M", "Renouv. M-1", "Renouvellements par semaine (M vs M-1)")
        ch4 = build_month_week_chart_png(comp_month_week, "Nouv. aff. M", "Nouv. aff. M-1", "Nouvelles affaires par semaine (M vs M-1)")

        if ch3:
            c.drawImage(ImageReader(BytesIO(ch3)), 14 * mm, y_chart_top2 - chart_h, width=chart_w, height=chart_h, mask="auto")
        if ch4:
            c.drawImage(ImageReader(BytesIO(ch4)), 14 * mm, y_chart_top2 - chart_h - 6 * mm - chart_h, width=chart_w, height=chart_h, mask="auto")

        c.showPage()

    # ================= PAGE 7 : TOP 10 AGENTS =================
    has_top = any(
        df is not None and not df.empty
        for df in [top_agents_ca, top_agents_cnt, top_pool_ca, top_pool_cnt]
    )

    if has_top:
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 24)
        c.drawCentredString(W / 2, H - 18 * mm, "TOP 10 AGENTS")

        c.setFont("Helvetica", 10.5)
        y_info3 = H - 30 * mm
        c.drawString(14 * mm, y_info3, f"Segment : {segment}")
        c.drawString(14 * mm, y_info3 - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info3 - 12 * mm, f"Période : {start_date} → {end_date}")
        c.drawRightString(W - 14 * mm, y_info3 - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))
        if week_n_label or week_n1_label:
            c.drawString(14 * mm, y_info3 - 18 * mm, f"Semaine N : {week_n_label} | N-1 : {week_n1_label}")

        def _top_table_df(df_in: pd.DataFrame | None, value_col: str) -> pd.DataFrame:
            if df_in is None or df_in.empty:
                return pd.DataFrame({"Agent": ["Aucune donnée"], value_col: [""]})

            df_out = df_in.copy()
            if value_col not in df_out.columns:
                other_cols = [c for c in df_out.columns if c != "Agent"]
                if other_cols:
                    df_out = df_out.rename(columns={other_cols[0]: value_col})

            if "Agent" not in df_out.columns:
                cols = list(df_out.columns)
                if cols:
                    df_out = df_out.rename(columns={cols[0]: "Agent"})
                else:
                    return pd.DataFrame({"Agent": ["Aucune donnée"], value_col: [""]})

            df_out = df_out[["Agent", value_col]].copy()
            df_out[value_col] = pd.to_numeric(df_out[value_col], errors="coerce").fillna(0).map(fmt_int)
            return df_out

        def _draw_top_block(title: str, df_in: pd.DataFrame, x0: float, y_start: float, col_w: float) -> float:
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 12.5)
            c.drawString(x0, y_start, title)

            y_table = y_start - 6 * mm
            col_ws = [col_w * 0.62, col_w * 0.38]
            _pdf_table_generic(
                c,
                df_in,
                x0=x0,
                start_y=y_table,
                col_ws=col_ws,
                row_h=7.2 * mm,
                align_right_cols={df_in.columns[1]},
                header_bg=colors.HexColor("#F4D35E"),
            )
            height = (len(df_in) + 1) * (7.2 * mm)
            return y_table - height - 6 * mm

        margin_x = 14 * mm
        gap_x = 8 * mm
        col_w = (W - 2 * margin_x - gap_x) / 2
        y_top = H - 60 * mm

        df_agents_ca = _top_table_df(top_agents_ca, "Prime TTC (FCFA)")
        df_pool_ca = _top_table_df(top_pool_ca, "Prime TTC (FCFA)")

        _draw_top_block("Top 10 Agents — Prime TTC", df_agents_ca, margin_x, y_top, col_w)
        _draw_top_block(
            "Top 10 Agents POOL TPV — Prime TTC",
            df_pool_ca,
            margin_x + col_w + gap_x,
            y_top,
            col_w,
        )
        c.showPage()

        # page suivante pour les tableaux "Nombre de contrats"
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, W, H, stroke=0, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(W / 2, H - 18 * mm, "TOP 10 AGENTS (suite)")

        c.setFont("Helvetica", 10.5)
        y_info4 = H - 30 * mm
        c.drawString(14 * mm, y_info4, f"Segment : {segment}")
        c.drawString(14 * mm, y_info4 - 6 * mm, f"Intermédiaire : {intermediaire}")
        c.drawString(14 * mm, y_info4 - 12 * mm, f"Période : {start_date} → {end_date}")
        c.drawRightString(W - 14 * mm, y_info4 - 12 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))
        if week_n_label or week_n1_label:
            c.drawString(14 * mm, y_info4 - 18 * mm, f"Semaine N : {week_n_label} | N-1 : {week_n1_label}")

        y_top2 = H - 60 * mm
        df_agents_cnt = _top_table_df(top_agents_cnt, "Nombre de contrats")
        df_pool_cnt = _top_table_df(top_pool_cnt, "Nombre de contrats")

        _draw_top_block("Top 10 Agents — Nombre de contrats", df_agents_cnt, margin_x, y_top2, col_w)
        _draw_top_block(
            "Top 10 Agents POOL TPV — Nombre de contrats",
            df_pool_cnt,
            margin_x + col_w + gap_x,
            y_top2,
            col_w,
        )

        c.showPage()

    c.save()
    buff.seek(0)
    return buff.read()
