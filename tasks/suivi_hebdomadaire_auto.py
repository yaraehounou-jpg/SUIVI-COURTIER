from __future__ import annotations

from datetime import datetime, timedelta
import re
import numpy as np
import pandas as pd
import polars as pl
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from utils import (
    normalize,
    add_segment,
    add_intermediaire,
    compute_kpis,
    tableau_produit,
    monthly_ca_table,
    weekly_avg_table,
    fmt_int,
)
from export_utils import export_pdf_dashboard

OBJECTIF_AGENT = 6_826_339_000
OBJECTIF_COURTIER = 8_000_000_000


# ---------------- CSS ----------------
def load_css(path: str = "assets/styles.css") -> None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except Exception:
        pass


# ---------------- Utils format ----------------
def fmt_space_int(x) -> str:
    return fmt_int(x)


def fmt_pct_1(x) -> str:
    try:
        return f"{float(x) * 100:.1f}%"
    except Exception:
        return "0.0%"


def get_objectif(segment: str) -> int:
    if segment in {"Agent", "Direct", "Direct + Agents", "Agents + Direct"}:
        return OBJECTIF_AGENT
    if segment in {"Courtier", "Courtiers"}:
        return OBJECTIF_COURTIER
    return OBJECTIF_AGENT + OBJECTIF_COURTIER


# ---------------- Reset exports when filters change ----------------
def reset_exports():
    st.session_state.pdf_bytes = None
    st.session_state.pdf_key = None


# ---------------- KPI HTML cards ----------------
def kpi_card(
    value: str,
    label: str,
    pct: str | None = None,
    kind: str = "default",
    pct_class: str | None = None,
) -> str:
    cls = "kpi_pct" + (f" {pct_class}" if pct_class else "")
    pct_html = f"<div class='{cls}'>{pct}</div>" if pct else ""
    return f"""
    <div class="kpi_card {kind}">
      <div class="kpi_value">{value}</div>
      <div class="kpi_footer">
        <div class="kpi_label">{label}</div>
        {pct_html}
      </div>
    </div>
    """


def ratio_card(ratio: str) -> str:
    return f"""
    <div class="kpi_card ratio">
      <div class="kpi_ratio_value">{ratio}</div>
      <div class="kpi_ratio_label">Ratio Objectif Annuel</div>
    </div>
    """


def _find_date_col(df: pd.DataFrame, keys: list[str]) -> str | None:
    for col in df.columns:
        key = re.sub(r"[^A-Z0-9]+", "_", str(col).upper()).strip("_")
        if key in keys:
            return col
    return None


def _find_col(df: pd.DataFrame, keys: list[str]) -> str | None:
    for col in df.columns:
        key = re.sub(r"[^A-Z0-9]+", "_", str(col).upper()).strip("_")
        if key in keys:
            return col
    return None


def _to_num(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("")
    s = s.str.replace("\u00A0", " ", regex=False)
    s = s.str.replace("\u202F", " ", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(r"[^0-9,\.\-]", "", regex=True)
    has_comma = s.str.contains(",", regex=False)
    has_dot = s.str.contains(r"\.", regex=True)
    s = s.where(~(has_comma & ~has_dot), s.str.replace(",", ".", regex=False))
    s = s.where(~(has_comma & has_dot), s.str.replace(",", "", regex=False))
    return pd.to_numeric(s, errors="coerce")


def _parse_dates_any(series: pd.Series) -> pd.Series:
    s = series.copy()
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    num = pd.to_numeric(s, errors="coerce")
    mask_num = dt.isna() & num.notna()
    if mask_num.any():
        dt.loc[mask_num] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask_num], unit="D")
    return dt


def _apply_segment_filter(df: pd.DataFrame, segment_choice: str) -> pd.DataFrame:
    if "SEGMENT" not in df.columns:
        return df.copy()
    if segment_choice == "Direct + Agents":
        return df[df["SEGMENT"].isin(["Agent", "Direct"])].copy()
    if segment_choice == "Courtiers":
        return df[df["SEGMENT"] == "Courtier"].copy()
    return df.copy()


def run() -> None:
    # CSS (dans run, pas au niveau global)
    load_css()

    # ---------------- Title banner ----------------
    st.markdown(
        """
        <div class="top_banner">
            <div class="top_title">PLATEFORME AUTOMOBILE</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------------- Init session ----------------
    if "df_base" not in st.session_state:
        st.session_state.df_base = pd.DataFrame()

    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None
        st.session_state.pdf_key = None

    # ---------------- Sidebar: Upload ----------------
    st.sidebar.header("📂 Données")
    files = st.sidebar.file_uploader(
        "Charger un ou plusieurs fichiers Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="uploader_files",
        on_change=reset_exports,
    )

    if files:
        dfs = []
        for f in files:
            try:
                dfs.append(pl.read_excel(f).to_pandas())
            except Exception as e:
                st.sidebar.error(f"Erreur lecture fichier {getattr(f, 'name', '')} : {e}")

        if dfs:
            df_raw = pd.concat(dfs, ignore_index=True)
            df = normalize(df_raw)
            df = add_segment(df)
            df = add_intermediaire(df)
            st.session_state.df_base = df

    df = st.session_state.df_base
    if df is None or df.empty:
        st.info("Charge tes fichiers Excel dans la sidebar.")
        st.stop()

    if "DATE_CREATE" not in df.columns:
        st.error("Colonne DATE_CREATE introuvable après normalisation.")
        st.stop()

    df_dates = df.dropna(subset=["DATE_CREATE"]).copy()
    if df_dates.empty:
        st.error("Aucune date valide dans DATE_CREATE.")
        st.stop()

    # ---------------- Sidebar: Filters ----------------
    st.sidebar.header("🎛️ Filtres")

    min_date = df_dates["DATE_CREATE"].min().date()
    max_date = df_dates["DATE_CREATE"].max().date()

    seg_options = ["Direct + Agents", "Courtiers"]
    segment_choice = st.sidebar.selectbox(
        "Segment",
        seg_options,
        index=0,
        key="segment_choice",
        on_change=reset_exports,
    )

    date_range = st.sidebar.date_input(
        "Période (DATE_CREATE)",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        key="date_range",
        on_change=reset_exports,
    )
    eval_date_b4 = st.sidebar.date_input(
        "Date d'évaluation (B4)",
        value=end_date if "end_date" in locals() else max_date,
        min_value=min_date,
        max_value=max_date,
        key="eval_date_b4",
        on_change=reset_exports,
    )

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_d, end_d = date_range
    else:
        start_d, end_d = min_date, max_date
    # si la période change, réancre la date d'évaluation sur la borne haute
    if eval_date_b4 < start_d or eval_date_b4 > end_d:
        eval_date_b4 = end_d

    # Liste intermédiaires dépend du segment
    df_seg = _apply_segment_filter(df, segment_choice)

    inter_list = sorted(df_seg["INTERMEDIAIRE"].dropna().astype(str).unique().tolist())
    inter_choice = st.sidebar.selectbox(
        "Intermédiaire",
        ["Tous"] + inter_list,
        index=0,
        key="inter_choice",
        on_change=reset_exports,
    )

    # bouton (optionnel)
    st.sidebar.button("🔄 Rafraîchir", on_click=st.rerun, use_container_width=True)

    # ---------------- Apply filters ----------------
    df_f = _apply_segment_filter(df, segment_choice)

    df_f = df_f[
        (df_f["DATE_CREATE"].dt.date >= start_d)
        & (df_f["DATE_CREATE"].dt.date <= end_d)
    ].copy()

    if inter_choice != "Tous":
        df_f = df_f[df_f["INTERMEDIAIRE"] == inter_choice].copy()

    st.caption(f"✅ Lignes filtrées : **{len(df_f):,}**".replace(",", " "))

    if df_f.empty:
        st.warning("Aucune donnée pour ces filtres.")
        st.stop()

    objectif_annuel = get_objectif(segment_choice)

    segment_label = segment_choice
    inter_label = inter_choice if inter_choice != "Tous" else (
        "Tous les agents + directs" if segment_choice == "Direct + Agents"
        else "Tous les courtiers"
    )

    start_date_str = str(start_d)
    end_date_str = str(end_d)

    # ---------------- Meta header ----------------
    info_left, info_right = st.columns([0.75, 0.25])
    with info_left:
        st.markdown(
            f"""
            <div class="meta_block">
                <div><b>Segment :</b> {segment_label}</div>
                <div><b>Intermédiaire :</b> {inter_label}</div>
                <div><b>Période :</b> {start_date_str} → {end_date_str}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with info_right:
        st.markdown(
            f"<div class='meta_time'>Dernière mise à jour : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div>",
            unsafe_allow_html=True,
        )

    # ================= KPI =================
    k = compute_kpis(df_f, objectif_annuel)
    # Evolution CA: periode actuelle vs meme periode annee precedente (meme filtres segment/intermediaire)
    prev_y_start = (pd.Timestamp(start_d) - pd.DateOffset(years=1)).date()
    prev_y_end = (pd.Timestamp(end_d) - pd.DateOffset(years=1)).date()
    df_prev_y = _apply_segment_filter(df, segment_choice)
    if inter_choice != "Tous":
        df_prev_y = df_prev_y[df_prev_y["INTERMEDIAIRE"] == inter_choice].copy()
    df_prev_y = df_prev_y[
        (df_prev_y["DATE_CREATE"].dt.date >= prev_y_start)
        & (df_prev_y["DATE_CREATE"].dt.date <= prev_y_end)
    ].copy()
    k_prev_y_kpi = compute_kpis(df_prev_y, objectif_annuel) if not df_prev_y.empty else {"ca": 0}
    ca_cur_kpi = float(k.get("ca", 0))
    ca_prev_y_kpi = float(k_prev_y_kpi.get("ca", 0))
    if ca_prev_y_kpi > 0:
        ca_evol_pct = (ca_cur_kpi - ca_prev_y_kpi) / ca_prev_y_kpi
    else:
        ca_evol_pct = 1.0 if ca_cur_kpi > 0 else 0.0
    ca_evol_up = ca_evol_pct > 0
    ca_evol_down = ca_evol_pct < 0
    ca_delta_amt = ca_cur_kpi - ca_prev_y_kpi
    ca_delta_sign = "+" if ca_delta_amt >= 0 else "-"
    if ca_evol_up:
        ca_evol_streamlit = (
            f"<span style='color:#0b8f3a;font-weight:800;'>▲ {fmt_pct_1(ca_evol_pct)} vs N-1</span>"
            f"<br><span style='color:#0b8f3a;font-weight:800;'>Δ {ca_delta_sign}{fmt_space_int(abs(ca_delta_amt))}</span>"
        )
        ca_evol_pdf = f"▲ {fmt_pct_1(ca_evol_pct)} | Δ {ca_delta_sign}{fmt_space_int(abs(ca_delta_amt))}"
    elif ca_evol_down:
        ca_evol_streamlit = (
            f"<span style='color:#c62828;font-weight:800;'>▼ {fmt_pct_1(ca_evol_pct)} vs N-1</span>"
            f"<br><span style='color:#c62828;font-weight:800;'>Δ {ca_delta_sign}{fmt_space_int(abs(ca_delta_amt))}</span>"
        )
        ca_evol_pdf = f"▼ {fmt_pct_1(ca_evol_pct)} | Δ {ca_delta_sign}{fmt_space_int(abs(ca_delta_amt))}"
    else:
        ca_evol_streamlit = (
            f"<span style='color:#444;font-weight:800;'>• {fmt_pct_1(ca_evol_pct)} vs N-1</span>"
            f"<br><span style='color:#444;font-weight:800;'>Δ {ca_delta_sign}{fmt_space_int(abs(ca_delta_amt))}</span>"
        )
        ca_evol_pdf = f"• {fmt_pct_1(ca_evol_pct)} | Δ {ca_delta_sign}{fmt_space_int(abs(ca_delta_amt))}"
    prime_nette_col = _find_col(df_f, ["PRIME_NETTE", "PRIME_NETTE_"])
    accessoires_col = _find_col(df_f, ["ACCESSOIRES", "ACCESSOIRE"])
    if prime_nette_col or accessoires_col:
        prime_nette_num = _to_num(df_f[prime_nette_col]).fillna(0) if prime_nette_col else pd.Series(0, index=df_f.index)
        accessoires_num = _to_num(df_f[accessoires_col]).fillna(0) if accessoires_col else pd.Series(0, index=df_f.index)
        ca_row = (prime_nette_num + accessoires_num)
    else:
        ca_row = _to_num(df_f["PRIME_TTC"]).fillna(0) if "PRIME_TTC" in df_f.columns else pd.Series(0, index=df_f.index)
    # Nombre de contrats hors motos (affiché à côté du total)
    col_prod_kpi = _find_col(
        df_f,
        [
            "PRODUIT",
            "BRANCHE",
            "LINE_OF_BUSINESS",
            "CATEGORIE",
        ],
    )
    if col_prod_kpi and col_prod_kpi in df_f.columns:
        prod_up = df_f[col_prod_kpi].astype("string").str.upper().fillna("")
        non_moto_mask = ~prod_up.str.contains("MOTO", na=False)
        nb_contrats_sans_moto = int(non_moto_mask.sum())
    else:
        non_moto_mask = pd.Series([True] * len(df_f), index=df_f.index)
        nb_contrats_sans_moto = int(k["souscriptions"])

    dfx_kpi = df_f.copy()
    if "EST_RENOUVELLER" not in dfx_kpi.columns:
        dfx_kpi["EST_RENOUVELLER"] = False
    dfx_kpi["EST_RENOUVELLER"] = dfx_kpi["EST_RENOUVELLER"].astype(bool)
    dfx_non_moto = dfx_kpi[non_moto_mask].copy()
    ren_non_moto = int(dfx_non_moto["EST_RENOUVELLER"].sum())
    new_non_moto = int((~dfx_non_moto["EST_RENOUVELLER"]).sum())
    pct_ren_non_moto = (ren_non_moto / nb_contrats_sans_moto) if nb_contrats_sans_moto else 0.0
    pct_new_non_moto = (new_non_moto / nb_contrats_sans_moto) if nb_contrats_sans_moto else 0.0

    # KPI Excel: NB.SI.ENS(Date d'expiration; "<=B4"), limité à Branche Automobile
    col_expiration = _find_date_col(
        df_f,
        ["DATE_D_EXPIRATION", "DATE_EXPIRATION", "DATE_DEXPIRATION", "DATE_ECHEANCE", "DATE_EXPIR"],
    )
    col_branche = _find_col(df_f, ["BRANCHE", "LINE_OF_BUSINESS", "LINE_OF_BUSINESS_"])
    if col_expiration and col_expiration in df_f.columns:
        exp_dt = _parse_dates_any(df_f[col_expiration])
        if col_branche and col_branche in df_f.columns:
            auto_mask = (
                df_f[col_branche]
                .astype("string")
                .str.upper()
                .str.contains("AUTO", na=False)
            )
        else:
            auto_mask = pd.Series([True] * len(df_f), index=df_f.index)
        nb_echeance_auto = int((exp_dt.notna() & (exp_dt.dt.date <= eval_date_b4) & auto_mask).sum())
    else:
        nb_echeance_auto = 0

    # Montant a reverser au pool = somme GARANTIE RC + GARANTIE DEFENSE RECOURS
    # sur les lignes identifiees comme contrats POOL.
    col_pool_flag = _find_col(df_f, ["POOLS_TPV", "POOL_TPV", "POOL", "CESSION_POOL_TPV_40"])
    col_gar_rc = _find_col(df_f, ["GARANTIE_RC", "GARANTIE__RC"])
    col_gar_dr = _find_col(
        df_f,
        [
            "GARANTIE_DEFENSE_RECOURS",
            "GARANTIE_DEFENSE__RECOURS",
            "GARANTIE_DEFENSE_RECOUR",
        ],
    )
    if col_pool_flag and (col_gar_rc or col_gar_dr):
        pool_raw = df_f[col_pool_flag]
        pool_num = _to_num(pool_raw).fillna(0)
        pool_txt = pool_raw.astype("string").str.upper().str.strip()
        pool_mask = (pool_num > 0) | pool_txt.isin(["TRUE", "VRAI", "OUI", "YES"])
        rc_num = _to_num(df_f[col_gar_rc]).fillna(0) if col_gar_rc else pd.Series(0, index=df_f.index)
        dr_num = _to_num(df_f[col_gar_dr]).fillna(0) if col_gar_dr else pd.Series(0, index=df_f.index)
        montant_reverser_pool = float((rc_num[pool_mask] + dr_num[pool_mask]).sum()) * 0.40
        if not df_prev_y.empty and col_pool_flag in df_prev_y.columns:
            pool_raw_prev = df_prev_y[col_pool_flag]
            pool_num_prev = _to_num(pool_raw_prev).fillna(0)
            pool_txt_prev = pool_raw_prev.astype("string").str.upper().str.strip()
            pool_mask_prev = (pool_num_prev > 0) | pool_txt_prev.isin(["TRUE", "VRAI", "OUI", "YES"])
            rc_prev = _to_num(df_prev_y[col_gar_rc]).fillna(0) if col_gar_rc and col_gar_rc in df_prev_y.columns else pd.Series(0, index=df_prev_y.index)
            dr_prev = _to_num(df_prev_y[col_gar_dr]).fillna(0) if col_gar_dr and col_gar_dr in df_prev_y.columns else pd.Series(0, index=df_prev_y.index)
            montant_reverser_pool_prev = float((rc_prev[pool_mask_prev] + dr_prev[pool_mask_prev]).sum()) * 0.40
        else:
            montant_reverser_pool_prev = 0.0
    else:
        montant_reverser_pool = 0.0
        montant_reverser_pool_prev = 0.0
    pct_pool_vs_ca = (float(k.get("montant_pool", 0)) / float(k["ca"])) if float(k["ca"]) else 0.0
    pct_reverser_pool_ca = (montant_reverser_pool / float(k["ca"])) if float(k["ca"]) else 0.0
    pct_cession_vs_pool = (montant_reverser_pool / float(k.get("montant_pool", 0))) if float(k.get("montant_pool", 0)) else 0.0
    pool_cur = float(k.get("montant_pool", 0))
    pool_prev = float(k_prev_y_kpi.get("montant_pool", 0))
    if pool_prev > 0:
        pool_evol_pct = (pool_cur - pool_prev) / pool_prev
    else:
        pool_evol_pct = 1.0 if pool_cur > 0 else 0.0
    pool_evol_up = pool_evol_pct > 0
    pool_evol_down = pool_evol_pct < 0
    if pool_evol_up:
        pool_evol_streamlit = f"<span style='color:#0b8f3a;font-weight:800;'>▲ {fmt_pct_1(pool_evol_pct)} vs N-1</span>"
        pool_evol_pdf = f"▲ {fmt_pct_1(pool_evol_pct)} vs N-1"
    elif pool_evol_down:
        pool_evol_streamlit = f"<span style='color:#c62828;font-weight:800;'>▼ {fmt_pct_1(pool_evol_pct)} vs N-1</span>"
        pool_evol_pdf = f"▼ {fmt_pct_1(pool_evol_pct)} vs N-1"
    else:
        pool_evol_streamlit = f"<span style='color:#444;font-weight:800;'>• {fmt_pct_1(pool_evol_pct)} vs N-1</span>"
        pool_evol_pdf = f"• {fmt_pct_1(pool_evol_pct)} vs N-1"

    hors_cur = float(k.get("montant_hors_pool", 0))
    hors_prev = float(k_prev_y_kpi.get("montant_hors_pool", 0))
    if hors_prev > 0:
        hors_evol_pct = (hors_cur - hors_prev) / hors_prev
    else:
        hors_evol_pct = 1.0 if hors_cur > 0 else 0.0
    hors_evol_up = hors_evol_pct > 0
    hors_evol_down = hors_evol_pct < 0
    if hors_evol_up:
        hors_evol_streamlit = f"<span style='color:#0b8f3a;font-weight:800;'>▲ {fmt_pct_1(hors_evol_pct)} vs N-1</span>"
        hors_evol_pdf = f"▲ {fmt_pct_1(hors_evol_pct)} vs N-1"
    elif hors_evol_down:
        hors_evol_streamlit = f"<span style='color:#c62828;font-weight:800;'>▼ {fmt_pct_1(hors_evol_pct)} vs N-1</span>"
        hors_evol_pdf = f"▼ {fmt_pct_1(hors_evol_pct)} vs N-1"
    else:
        hors_evol_streamlit = f"<span style='color:#444;font-weight:800;'>• {fmt_pct_1(hors_evol_pct)} vs N-1</span>"
        hors_evol_pdf = f"• {fmt_pct_1(hors_evol_pct)} vs N-1"
    pct_hors_vs_ca = (hors_cur / float(k["ca"])) if float(k["ca"]) else 0.0
    denom_pool_hors = float(k.get("montant_pool", 0)) + hors_cur
    pct_hors_share = (hors_cur / denom_pool_hors) if denom_pool_hors else 0.0

    row1 = st.columns(4)
    row2 = st.columns(3)
    row3 = st.columns(3)
    row4 = st.columns(3)

    with row1[0]:
        st.markdown(kpi_card(fmt_space_int(k["taxes"]), "Impôts"), unsafe_allow_html=True)
    with row1[1]:
        st.markdown(
            kpi_card(fmt_space_int(k["ca"]), "Chiffre d'affaires", pct=ca_evol_streamlit, kind="highlight"),
            unsafe_allow_html=True,
        )
    with row1[2]:
        st.markdown(
            kpi_card(
                fmt_space_int(nb_echeance_auto),
                "Branche = Automobile",
                pct=f"Échéance ≤ {eval_date_b4.strftime('%d/%m/%Y')}",
                kind="whiteframe",
            ),
            unsafe_allow_html=True,
        )
    with row1[3]:
        st.markdown(kpi_card(fmt_space_int(k["nb_agents_actifs"]), "Nom Agents Actifs"), unsafe_allow_html=True)

    with row2[0]:
        pool_value = f"{fmt_space_int(k['montant_pool'])} | {fmt_space_int(montant_reverser_pool)}"
        pool_shares_black = (
            f"<span style='color:#111;font-weight:800;'>{fmt_pct_1(pct_pool_vs_ca)} | {fmt_pct_1(pct_cession_vs_pool)}</span>"
        )
        pool_pct = f"{pool_shares_black}<br>{pool_evol_streamlit}"
        st.markdown(
            kpi_card(pool_value, "Montant POOL TPV | Prime de cession", pct=pool_pct, kind="highlight"),
            unsafe_allow_html=True,
        )
    with row2[1]:
        st.markdown(ratio_card(fmt_pct_1(k["ratio_obj"])), unsafe_allow_html=True)
    with row2[2]:
        st.markdown(
            kpi_card(
                fmt_space_int(k["montant_hors_pool"]),
                "Montant hors pool",
                pct=(
                    f"<span style='color:#111;font-weight:800;'>{fmt_pct_1(pct_hors_vs_ca)} | {fmt_pct_1(pct_hors_share)}</span>"
                    f"<br>{hors_evol_streamlit}"
                ),
                kind="highlight",
            ),
            unsafe_allow_html=True,
        )

    with row3[0]:
        ren_value = f"{fmt_space_int(k['renouvellements'])} | {fmt_space_int(ren_non_moto)}"
        ren_pct = f"{fmt_pct_1(k['pct_ren'])} | {fmt_pct_1(pct_ren_non_moto)}"
        st.markdown(
            kpi_card(ren_value, "Renouvellement | Hors motos", pct=ren_pct, kind="dark"),
            unsafe_allow_html=True,
        )
    with row3[1]:
        contrats_value = f"{fmt_space_int(k['souscriptions'])} | {fmt_space_int(nb_contrats_sans_moto)}"
        st.markdown(kpi_card(contrats_value, "Nombre de contrats | Hors motos", kind="dark"), unsafe_allow_html=True)
    with row3[2]:
        new_value = f"{fmt_space_int(k['nouvelles_affaires'])} | {fmt_space_int(new_non_moto)}"
        new_pct = f"{fmt_pct_1(k['pct_new'])} | {fmt_pct_1(pct_new_non_moto)}"
        st.markdown(
            kpi_card(new_value, "Nouvelles affaires | Hors motos", pct=new_pct, kind="dark"),
            unsafe_allow_html=True,
        )

    if "DATE_CREATE" in df_f.columns:
        nb_days = int(df_f["DATE_CREATE"].dt.date.nunique())
    else:
        nb_days = int((end_d - start_d).days + 1)
    nb_days = max(nb_days, 1)
    avg_contracts = k["souscriptions"] / nb_days
    avg_ca = k["ca"] / nb_days

    # Priorite: colonne DUREE (MOIS) avec conversion jours -> mois selon la regle demandee
    col_duree = _find_col(
        df_f,
        [
            "DUREE_MOIS",
            "DUREE_MOIS_",
            "DUREE_MOIS__",
            "DUREE",
            "DURATION",
        ],
    )
    if col_duree:
        duree_raw = _to_num(df_f[col_duree]).fillna(0.0)
        day_to_month_values = {30, 31, 60, 90, 365, 364, 91, 120, 150, 200, 300, 180}
        duree_mois = duree_raw.copy()
        mask_days = duree_raw.isin(day_to_month_values)
        duree_mois.loc[mask_days] = duree_raw.loc[mask_days] / 30.0

        # Contrats "1 mois" = 1 mois natif ou conversion jours proche de 1 mois (30/31 j)
        one_month_mask = np.isclose(duree_mois, 1.0, atol=0.2)
        nb_1_month = int(one_month_mask.sum())
        ca_1_month = float(ca_row[one_month_mask].sum())
    else:
        # Fallback: derive de la duree entre date effet et expiration
        col_effet = _find_date_col(
            df_f,
            [
                "DATE_EFFET",
                "DATE_D_EFFET",
                "DATE_DEFFET",
                "DATE_EFFET_DEBUT",
            ],
        )
        col_exp = _find_date_col(
            df_f,
            [
                "DATE_EXPIRATION",
                "DATE_D_EXPIRATION",
                "DATE_DEXPIRATION",
                "DATE_ECHEANCE",
            ],
        )
        if col_effet and col_exp:
            eff = _parse_dates_any(df_f[col_effet])
            exp = _parse_dates_any(df_f[col_exp])
            delta_days = (exp - eff).dt.days
            one_month_mask = (delta_days >= 28) & (delta_days <= 31)
            nb_1_month = int(one_month_mask.sum())
            ca_1_month = float(ca_row[one_month_mask].sum())
        else:
            nb_1_month = 0
            ca_1_month = 0.0

    # Meme KPI 1 mois sur N-1 pour afficher l'ecart en montant
    if not df_prev_y.empty:
        prime_nette_prev_col = _find_col(df_prev_y, ["PRIME_NETTE", "PRIME_NETTE_"])
        accessoires_prev_col = _find_col(df_prev_y, ["ACCESSOIRES", "ACCESSOIRE"])
        if prime_nette_prev_col or accessoires_prev_col:
            prime_prev_num = _to_num(df_prev_y[prime_nette_prev_col]).fillna(0) if prime_nette_prev_col else pd.Series(0, index=df_prev_y.index)
            acc_prev_num = _to_num(df_prev_y[accessoires_prev_col]).fillna(0) if accessoires_prev_col else pd.Series(0, index=df_prev_y.index)
            ca_prev_row = prime_prev_num + acc_prev_num
        else:
            ca_prev_row = _to_num(df_prev_y["PRIME_TTC"]).fillna(0) if "PRIME_TTC" in df_prev_y.columns else pd.Series(0, index=df_prev_y.index)

        col_duree_prev = _find_col(df_prev_y, ["DUREE_MOIS", "DUREE_MOIS_", "DUREE_MOIS__", "DUREE", "DURATION"])
        if col_duree_prev:
            duree_prev_raw = _to_num(df_prev_y[col_duree_prev]).fillna(0.0)
            day_to_month_values_prev = {30, 31, 60, 90, 365, 364, 91, 120, 150, 200, 300, 180}
            duree_prev_mois = duree_prev_raw.copy()
            mask_prev_days = duree_prev_raw.isin(day_to_month_values_prev)
            duree_prev_mois.loc[mask_prev_days] = duree_prev_raw.loc[mask_prev_days] / 30.0
            one_month_prev_mask = np.isclose(duree_prev_mois, 1.0, atol=0.2)
            nb_1_month_prev = int(one_month_prev_mask.sum())
            ca_1_month_prev = float(ca_prev_row[one_month_prev_mask].sum())
        else:
            col_effet_prev = _find_date_col(df_prev_y, ["DATE_EFFET", "DATE_D_EFFET", "DATE_DEFFET", "DATE_EFFET_DEBUT"])
            col_exp_prev = _find_date_col(df_prev_y, ["DATE_EXPIRATION", "DATE_D_EXPIRATION", "DATE_DEXPIRATION", "DATE_ECHEANCE"])
            if col_effet_prev and col_exp_prev:
                eff_prev = _parse_dates_any(df_prev_y[col_effet_prev])
                exp_prev = _parse_dates_any(df_prev_y[col_exp_prev])
                delta_prev_days = (exp_prev - eff_prev).dt.days
                one_month_prev_mask = (delta_prev_days >= 28) & (delta_prev_days <= 31)
                nb_1_month_prev = int(one_month_prev_mask.sum())
                ca_1_month_prev = float(ca_prev_row[one_month_prev_mask].sum())
            else:
                nb_1_month_prev = 0
                ca_1_month_prev = 0.0
    else:
        nb_1_month_prev = 0
        ca_1_month_prev = 0.0
    delta_1_month_amt = float(ca_1_month) - float(ca_1_month_prev)
    if nb_1_month_prev > 0:
        one_month_count_evol = (nb_1_month - nb_1_month_prev) / nb_1_month_prev
    else:
        one_month_count_evol = 1.0 if nb_1_month > 0 else 0.0
    if one_month_count_evol > 0:
        one_month_count_evol_streamlit = f"<span style='color:#0b8f3a;font-weight:800;'>▲ {fmt_pct_1(one_month_count_evol)} vs N-1</span>"
        one_month_count_evol_pdf = f"▲ {fmt_pct_1(one_month_count_evol)} vs N-1"
    elif one_month_count_evol < 0:
        one_month_count_evol_streamlit = f"<span style='color:#c62828;font-weight:800;'>▼ {fmt_pct_1(one_month_count_evol)} vs N-1</span>"
        one_month_count_evol_pdf = f"▼ {fmt_pct_1(one_month_count_evol)} vs N-1"
    else:
        one_month_count_evol_streamlit = f"<span style='color:#444;font-weight:800;'>• {fmt_pct_1(one_month_count_evol)} vs N-1</span>"
        one_month_count_evol_pdf = f"• {fmt_pct_1(one_month_count_evol)} vs N-1"

    if ca_1_month_prev > 0:
        one_month_amt_evol = (ca_1_month - ca_1_month_prev) / ca_1_month_prev
    else:
        one_month_amt_evol = 1.0 if ca_1_month > 0 else 0.0
    if one_month_amt_evol > 0:
        one_month_amt_evol_streamlit = f"<span style='color:#0b8f3a;font-weight:800;'>▲ {fmt_pct_1(one_month_amt_evol)} vs N-1</span>"
        one_month_amt_evol_pdf = f"▲ {fmt_pct_1(one_month_amt_evol)} vs N-1"
    elif one_month_amt_evol < 0:
        one_month_amt_evol_streamlit = f"<span style='color:#c62828;font-weight:800;'>▼ {fmt_pct_1(one_month_amt_evol)} vs N-1</span>"
        one_month_amt_evol_pdf = f"▼ {fmt_pct_1(one_month_amt_evol)} vs N-1"
    else:
        one_month_amt_evol_streamlit = f"<span style='color:#444;font-weight:800;'>• {fmt_pct_1(one_month_amt_evol)} vs N-1</span>"
        one_month_amt_evol_pdf = f"• {fmt_pct_1(one_month_amt_evol)} vs N-1"
    one_month_amt_evol_streamlit_pipe = one_month_amt_evol_streamlit.replace(">", ">| ", 1)

    with row4[0]:
        st.markdown(
            kpi_card(f"{avg_contracts:,.1f}".replace(",", " "), "Contrats / jour (moy.)"),
            unsafe_allow_html=True,
        )
    with row4[1]:
        st.markdown(
            kpi_card(fmt_space_int(avg_ca), "CA / jour (moy.)"),
            unsafe_allow_html=True,
        )
    with row4[2]:
        pct_1_month_non_moto = (nb_1_month / nb_contrats_sans_moto) if nb_contrats_sans_moto else 0.0
        pct_1_month_all = (nb_1_month / float(k["souscriptions"])) if float(k["souscriptions"]) else 0.0
        st.markdown(
            kpi_card(
                fmt_space_int(nb_1_month),
                "Contrats durée 1 mois",
                pct=(
                    f"<span style='color:#111;font-weight:800;'>{fmt_pct_1(pct_1_month_non_moto)} | {fmt_pct_1(pct_1_month_all)}</span>"
                    f"<br>{one_month_count_evol_streamlit}"
                ),
                pct_class="kpi_pct_center kpi_pct_up",
            ),
            unsafe_allow_html=True,
        )

    pct_amt_1m_vs_ca = (ca_1_month / float(k["ca"])) if float(k["ca"]) else 0.0
    st.markdown(
        kpi_card(
            fmt_space_int(ca_1_month),
            "Montant 1 mois",
            pct=(
                f"<span style='color:#111;font-weight:800;'>{fmt_pct_1(pct_amt_1m_vs_ca)}</span>"
                f"<br>{one_month_amt_evol_streamlit_pipe}"
            ),
            kind="highlight",
            pct_class="kpi_pct_center kpi_pct_up",
        ),
        unsafe_allow_html=True,
    )

    # ================= Tableau 1 mois -> echeance fevrier =================
    st.markdown("### 📋 Contrats 1 mois arrivant à échéance en février")
    col_echeance = _find_date_col(
        df_f,
        [
            "DATE_ECHEANCE",
            "DATE_ECH",
            "ECHEANCE",
            "DATE_EXPIRATION",
            "DATE_EXPIR",
            "DATE_DEXPIRATION",
        ],
    )
    col_immat = _find_col(
        df_f,
        [
            "IMMATRICULATION",
            "IMMATRICULATION",
            "NUMERO_IMMATRICULATION",
            "NUM_IMMATRICULATION",
            "PLAQUE",
        ],
    )
    target_year_feb = int(end_d.year)

    tb_1m_feb = pd.DataFrame()
    if col_echeance:
        ech_dt = _parse_dates_any(df_f[col_echeance])
        feb_ech_mask = (ech_dt.dt.month == 2) & (ech_dt.dt.year == target_year_feb)
        cohort_mask = one_month_mask & feb_ech_mask
        cohort_df = dfx_kpi.loc[cohort_mask].copy()

        if cohort_df.empty:
            st.info(f"Aucun contrat 1 mois avec échéance en février {target_year_feb} sur la période filtrée.")
        else:
            prod_col_tbl = col_prod_kpi if col_prod_kpi in cohort_df.columns else None
            if prod_col_tbl:
                cohort_df[prod_col_tbl] = cohort_df[prod_col_tbl].astype("string").fillna("INCONNU")
            else:
                cohort_df["__PROD__"] = "TOUS PRODUITS"
                prod_col_tbl = "__PROD__"

            if col_immat and (col_immat in cohort_df.columns) and (col_immat in df.columns):
                cohort_immat = cohort_df[col_immat].astype("string").str.strip().replace("", pd.NA).dropna().unique().tolist()
                if cohort_immat:
                    renew_feb = df.copy()
                    if "EST_RENOUVELLER" in renew_feb.columns:
                        renew_feb = renew_feb[renew_feb["EST_RENOUVELLER"].astype(bool)].copy()
                    if "DATE_CREATE" in renew_feb.columns:
                        renew_feb = renew_feb[
                            (renew_feb["DATE_CREATE"].dt.month == 2)
                            & (renew_feb["DATE_CREATE"].dt.year == target_year_feb)
                        ].copy()
                    renew_feb[col_immat] = renew_feb[col_immat].astype("string").str.strip().replace("", pd.NA)
                    renew_immat_set = set(renew_feb[col_immat].dropna().unique().tolist())
                else:
                    renew_immat_set = set()

                rows = []
                for prod, grp in cohort_df.groupby(prod_col_tbl, dropna=False):
                    immats = grp[col_immat].astype("string").str.strip().replace("", pd.NA).dropna().unique().tolist()
                    tot = len(immats)
                    ren = len([x for x in immats if x in renew_immat_set])
                    amt_tot = float(ca_row.loc[grp.index].sum())
                    grp_immats = grp[col_immat].astype("string").str.strip().replace("", pd.NA)
                    renew_grp_mask = grp_immats.isin(list(renew_immat_set))
                    amt_ren = float(ca_row.loc[grp.index[renew_grp_mask.fillna(False)]].sum())
                    rows.append(
                        {
                            "PRODUIT": str(prod),
                            "Souscriptions 1 mois (échéance février)": int(tot),
                            "Renouvelés en février": int(ren),
                            "Montant souscriptions (FCFA)": amt_tot,
                            "Montant renouvelés (FCFA)": amt_ren,
                            "Part réalisée": (amt_ren / amt_tot) if amt_tot else 0.0,
                            "% renouvellement": (ren / tot) if tot else 0.0,
                        }
                    )
                tb_1m_feb = pd.DataFrame(rows)
            else:
                # Fallback sans immatriculation: base ligne
                tb_1m_feb = (
                    cohort_df.groupby(prod_col_tbl, dropna=False)["EST_RENOUVELLER"]
                    .agg(["size", "sum"])
                    .reset_index()
                    .rename(
                        columns={
                            prod_col_tbl: "PRODUIT",
                            "size": "Souscriptions 1 mois (échéance février)",
                            "sum": "Renouvelés en février",
                        }
                    )
                )
                tb_1m_feb["Montant souscriptions (FCFA)"] = (
                    cohort_df.groupby(prod_col_tbl, dropna=False)
                    .apply(lambda g: float(ca_row.loc[g.index].sum()))
                    .reset_index(drop=True)
                )
                tb_1m_feb["Montant renouvelés (FCFA)"] = (
                    cohort_df.groupby(prod_col_tbl, dropna=False)
                    .apply(lambda g: float(ca_row.loc[g[g["EST_RENOUVELLER"].astype(bool)].index].sum()))
                    .reset_index(drop=True)
                )
                tb_1m_feb["Part réalisée"] = (
                    tb_1m_feb["Montant renouvelés (FCFA)"]
                    / tb_1m_feb["Montant souscriptions (FCFA)"].replace(0, np.nan)
                ).fillna(0.0)
                tb_1m_feb["% renouvellement"] = (
                    tb_1m_feb["Renouvelés en février"]
                    / tb_1m_feb["Souscriptions 1 mois (échéance février)"].replace(0, np.nan)
                ).fillna(0.0)

            total_sous = int(tb_1m_feb["Souscriptions 1 mois (échéance février)"].sum())
            total_ren = int(tb_1m_feb["Renouvelés en février"].sum())
            total_amt_sous = float(tb_1m_feb["Montant souscriptions (FCFA)"].sum())
            total_amt_ren = float(tb_1m_feb["Montant renouvelés (FCFA)"].sum())
            tb_1m_feb = tb_1m_feb.sort_values(
                by="Souscriptions 1 mois (échéance février)",
                ascending=False,
            )
            tb_1m_feb = pd.concat(
                [
                    tb_1m_feb,
                    pd.DataFrame(
                        [
                            {
                                "PRODUIT": "Grand Total",
                                "Souscriptions 1 mois (échéance février)": total_sous,
                                "Renouvelés en février": total_ren,
                                "Montant souscriptions (FCFA)": total_amt_sous,
                                "Montant renouvelés (FCFA)": total_amt_ren,
                                "Part réalisée": (total_amt_ren / total_amt_sous) if total_amt_sous else 0.0,
                                "% renouvellement": (total_ren / total_sous) if total_sous else 0.0,
                            }
                        ]
                    ),
                ],
                ignore_index=True,
            )

            st.dataframe(
                tb_1m_feb.style.format(
                    {
                        "Souscriptions 1 mois (échéance février)": lambda x: fmt_space_int(x),
                        "Renouvelés en février": lambda x: fmt_space_int(x),
                        "Montant souscriptions (FCFA)": lambda x: fmt_space_int(x),
                        "Montant renouvelés (FCFA)": lambda x: fmt_space_int(x),
                        "Part réalisée": lambda x: fmt_pct_1(x),
                        "% renouvellement": lambda x: fmt_pct_1(x),
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
    else:
        st.info("Colonne date d'échéance introuvable pour calculer le tableau des contrats 1 mois.")

    st.markdown(
        f"<div class='objectif_footer'>Objectif annuel : <b>{fmt_space_int(objectif_annuel)}</b> FCFA</div>",
        unsafe_allow_html=True,
    )

    st.divider()

    # ================= TOP AGENTS =================
    st.subheader("🏆 TOP 10 AGENTS")
    agent_col = "NOM_AGENT" if "NOM_AGENT" in df_f.columns else "INTERMEDIAIRE"
    top_agents_ca = None
    top_agents_cnt = None
    top_pool_ca = None
    top_pool_cnt = None
    if agent_col in df_f.columns:
        df_agents = df_f.copy()
        df_agents[agent_col] = df_agents[agent_col].astype("string").str.strip().replace("", pd.NA)
        df_agents = df_agents.dropna(subset=[agent_col])

        by_ca = (
            df_agents.groupby(agent_col)["PRIME_TTC"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
            .rename(columns={agent_col: "Agent", "PRIME_TTC": "Prime TTC (FCFA)"})
        )
        top_agents_ca = by_ca.copy()
        by_cnt = (
            df_agents.groupby(agent_col)
            .size()
            .sort_values(ascending=False)
            .head(10)
            .reset_index(name="Nombre de contrats")
            .rename(columns={agent_col: "Agent"})
        )
        top_agents_cnt = by_cnt.copy()

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Top 10 Agents — Prime TTC**")
            st.dataframe(
                by_ca.style.format({"Prime TTC (FCFA)": lambda x: fmt_space_int(x)}),
                use_container_width=True,
                hide_index=True,
            )
        with c2:
            st.markdown("**Top 10 Agents — Nombre de contrats**")
            st.dataframe(
                by_cnt.style.format({"Nombre de contrats": lambda x: fmt_space_int(x)}),
                use_container_width=True,
                hide_index=True,
            )

        # Top 10 Agents POOL TPV
        if "POOLS_TPV" in df_agents.columns:
            df_pool = df_agents[df_agents["POOLS_TPV"] == 1].copy()
            if not df_pool.empty:
                by_ca_pool = (
                    df_pool.groupby(agent_col)["PRIME_TTC"]
                    .sum()
                    .sort_values(ascending=False)
                    .head(10)
                    .reset_index()
                    .rename(columns={agent_col: "Agent", "PRIME_TTC": "Prime TTC (FCFA)"})
                )
                top_pool_ca = by_ca_pool.copy()
                by_cnt_pool = (
                    df_pool.groupby(agent_col)
                    .size()
                    .sort_values(ascending=False)
                    .head(10)
                    .reset_index(name="Nombre de contrats")
                    .rename(columns={agent_col: "Agent"})
                )
                top_pool_cnt = by_cnt_pool.copy()
                c3, c4 = st.columns(2)
                with c3:
                    st.markdown("**Top 10 Agents POOL TPV — Prime TTC**")
                    st.dataframe(
                        by_ca_pool.style.format({"Prime TTC (FCFA)": lambda x: fmt_space_int(x)}),
                        use_container_width=True,
                        hide_index=True,
                    )
                with c4:
                    st.markdown("**Top 10 Agents POOL TPV — Nombre de contrats**")
                    st.dataframe(
                        by_cnt_pool.style.format({"Nombre de contrats": lambda x: fmt_space_int(x)}),
                        use_container_width=True,
                        hide_index=True,
                    )
    else:
        st.info("Aucune colonne Agent disponible (NOM_AGENT / INTERMEDIAIRE).")

    # ================= COMPARATIF SEMAINE =================
    st.subheader("📆 COMPARATIF SEMAINE COURANTE VS SEMAINE PRÉCÉDENTE")

    comp_week_df = None
    week_n_label = ""
    week_n1_label = ""
    if "DATE_CREATE" in df_f.columns:
        # Semaine basée sur la période filtrée (fenêtre glissante sur 7 jours)
        w_end = end_d
        w_start = max(start_d, end_d - timedelta(days=6))
        prev_end = w_start - timedelta(days=1)
        prev_start = max(start_d, prev_end - timedelta(days=6))

        df_cur = df_f[
            (df_f["DATE_CREATE"].dt.date >= w_start) & (df_f["DATE_CREATE"].dt.date <= w_end)
        ].copy()
        df_prev = df_f[
            (df_f["DATE_CREATE"].dt.date >= prev_start) & (df_f["DATE_CREATE"].dt.date <= prev_end)
        ].copy()

        k_cur = compute_kpis(df_cur, objectif_annuel)
        k_prev = compute_kpis(df_prev, objectif_annuel)

        metrics = [
            ("Nombre de contrats", k_cur["souscriptions"], k_prev["souscriptions"]),
            ("Renouvellements", k_cur["renouvellements"], k_prev["renouvellements"]),
            ("Nouvelles affaires", k_cur["nouvelles_affaires"], k_prev["nouvelles_affaires"]),
            ("Chiffre d'affaires", k_cur["ca"], k_prev["ca"]),
        ]

        rows = []
        for label, cur, prev in metrics:
            delta = cur - prev
            pct = (delta / prev) if prev else 0.0
            rows.append({
                "Indicateur": label,
                "Semaine courante": cur,
                "Semaine précédente": prev,
                "Écart": delta,
                "Variation %": pct,
            })

        comp_df = pd.DataFrame(rows)
        comp_week_df = comp_df.copy()
        week_n_label = f"{w_start.strftime('%d/%m/%Y')} au {w_end.strftime('%d/%m/%Y')}"
        week_n1_label = f"{prev_start.strftime('%d/%m/%Y')} au {prev_end.strftime('%d/%m/%Y')}"
        st.dataframe(
            comp_df.style.format({
                "Semaine courante": lambda x: fmt_space_int(x),
                "Semaine précédente": lambda x: fmt_space_int(x),
                "Écart": lambda x: fmt_space_int(x),
                "Variation %": lambda x: f"{x*100:.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
        )

        fig_week = go.Figure()
        chart_metrics = metrics[:3]
        fig_week.add_trace(go.Bar(
            x=[m[0] for m in chart_metrics],
            y=[m[1] for m in chart_metrics],
            name="Semaine courante",
        ))
        fig_week.add_trace(go.Bar(
            x=[m[0] for m in chart_metrics],
            y=[m[2] for m in chart_metrics],
            name="Semaine précédente",
        ))
        w_label = f"{w_start.strftime('%d/%m/%Y')} au {w_end.strftime('%d/%m/%Y')}"
        p_label = f"{prev_start.strftime('%d/%m/%Y')} au {prev_end.strftime('%d/%m/%Y')}"
        fig_week.update_layout(
            barmode="group",
            title=f"Comparatif semaines ({w_label} vs {p_label})",
            height=420,
            margin=dict(l=40, r=30, t=70, b=40),
        )
        st.plotly_chart(fig_week, use_container_width=True)
    else:
        st.info("Colonne DATE_CREATE manquante pour le comparatif hebdomadaire.")

    st.divider()

    # ================= COMPARATIF SEMAINES DU MOIS =================
    st.subheader("🗓️ Comparatif semaines du mois (vs mois précédent)")

    df_compare_base = _apply_segment_filter(df, segment_choice)
    if inter_choice != "Tous":
        df_compare_base = df_compare_base[df_compare_base["INTERMEDIAIRE"] == inter_choice].copy()
    df_compare_base["DATE_CREATE"] = pd.to_datetime(df_compare_base["DATE_CREATE"], errors="coerce")
    df_compare_base = df_compare_base.dropna(subset=["DATE_CREATE"]).copy()
    df_compare_base["DATE_CREATE_DATE"] = df_compare_base["DATE_CREATE"].dt.date

    cmp_month_week = None
    month_week_label = ""
    cmp_month_week_y = None
    month_week_label_y = ""
    comp_period_month = None
    comp_period_label = ""
    comp_pool_period_year = None
    comp_pool_period_label = ""
    if "DATE_CREATE" in df_compare_base.columns and not df_compare_base.empty:
        cur_month_start = end_d.replace(day=1)
        cur_month_end = (pd.Timestamp(cur_month_start) + pd.offsets.MonthEnd(0)).date()
        prev_month_end = cur_month_start - timedelta(days=1)
        prev_month_start = prev_month_end.replace(day=1)

        week_ranges = []
        for start_day in [1, 8, 15, 22, 29]:
            # Evite ValueError sur les mois courts (ex: fevrier sans 29/30/31)
            if start_day > cur_month_end.day:
                continue
            start_date = cur_month_start.replace(day=start_day)
            end_day = min(start_day + 6, cur_month_end.day)
            end_date = cur_month_start.replace(day=end_day)
            week_ranges.append((start_date, end_date))

        rows = []
        for w_start, w_end in week_ranges:
            p_start = prev_month_start.replace(day=min(w_start.day, prev_month_end.day))
            p_end = prev_month_start.replace(day=min(w_end.day, prev_month_end.day))

            df_cur = df_compare_base[
                (df_compare_base["DATE_CREATE_DATE"] >= w_start)
                & (df_compare_base["DATE_CREATE_DATE"] <= w_end)
            ]
            df_prev = df_compare_base[
                (df_compare_base["DATE_CREATE_DATE"] >= p_start)
                & (df_compare_base["DATE_CREATE_DATE"] <= p_end)
            ]

            k_cur = compute_kpis(df_cur, objectif_annuel)
            k_prev = compute_kpis(df_prev, objectif_annuel)

            rows.append({
                "Semaine": f"{w_start.strftime('%d/%m')}–{w_end.strftime('%d/%m')}",
                "Contrats M": k_cur["souscriptions"],
                "Contrats M-1": k_prev["souscriptions"],
                "Δ Contrats": k_cur["souscriptions"] - k_prev["souscriptions"],
                "Renouv. M": k_cur["renouvellements"],
                "Renouv. M-1": k_prev["renouvellements"],
                "Δ Renouv.": k_cur["renouvellements"] - k_prev["renouvellements"],
                "Nouv. aff. M": k_cur["nouvelles_affaires"],
                "Nouv. aff. M-1": k_prev["nouvelles_affaires"],
                "Δ Nouv.": k_cur["nouvelles_affaires"] - k_prev["nouvelles_affaires"],
                "CA M": k_cur["ca"],
                "CA M-1": k_prev["ca"],
                "Δ CA": k_cur["ca"] - k_prev["ca"],
            })

        cmp_month_week = pd.DataFrame(rows)
        num_cols = [c for c in cmp_month_week.columns if c != "Semaine"]
        total_row = {"Semaine": "Total"}
        total_row.update(cmp_month_week[num_cols].sum().to_dict())
        cmp_month_week = pd.concat([cmp_month_week, pd.DataFrame([total_row])], ignore_index=True)
        month_week_label = (
            f"Mois courant : {cur_month_start.strftime('%m/%Y')} | "
            f"Mois précédent : {prev_month_start.strftime('%m/%Y')}"
        )
        st.caption(month_week_label)
        total_idx = len(cmp_month_week) - 1
        def _style_cmp_row(row):
            styles = []
            is_total = str(row.get("Semaine", "")) == "Total"
            for col in row.index:
                if is_total:
                    styles.append("font-weight:700; color:#000; background-color:#e9ecef")
                elif str(col).startswith("Δ"):
                    styles.append("background-color:#fff3cd")
                else:
                    styles.append("")
            return styles

        st.dataframe(
            cmp_month_week.style.format({
                "Contrats M": fmt_space_int,
                "Contrats M-1": fmt_space_int,
                "Δ Contrats": fmt_space_int,
                "Renouv. M": fmt_space_int,
                "Renouv. M-1": fmt_space_int,
                "Δ Renouv.": fmt_space_int,
                "Nouv. aff. M": fmt_space_int,
                "Nouv. aff. M-1": fmt_space_int,
                "Δ Nouv.": fmt_space_int,
                "CA M": fmt_space_int,
                "CA M-1": fmt_space_int,
                "Δ CA": fmt_space_int,
            }).set_table_styles([
                {"selector": "thead th", "props": [
                    ("background-color", "#F4D35E"),
                    ("color", "#1b2a41"),
                    ("font-weight", "700"),
                    ("text-align", "center"),
                ]},
                {"selector": "tbody td", "props": [
                    ("border-bottom", "1px solid #e6e8ec"),
                    ("padding", "8px"),
                    ("font-size", "13px"),
                ]},
                {"selector": "tbody td:not(:first-child)", "props": [
                    ("text-align", "right"),
                ]},
            ]).apply(_style_cmp_row, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        c1, c2 = st.columns(2)
        with c1:
            fig_contracts = go.Figure()
            fig_contracts.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["Contrats M"],
                name="Contrats M",
            ))
            fig_contracts.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["Contrats M-1"],
                name="Contrats M-1",
            ))
            fig_contracts.update_layout(
                barmode="group",
                title="Contrats par semaine (M vs M-1)",
                height=360,
                margin=dict(l=40, r=20, t=60, b=40),
            )
            st.plotly_chart(fig_contracts, use_container_width=True)

        with c2:
            fig_ca = go.Figure()
            fig_ca.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["CA M"],
                name="CA M",
            ))
            fig_ca.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["CA M-1"],
                name="CA M-1",
            ))
            fig_ca.update_layout(
                barmode="group",
                title="Chiffre d'affaires par semaine (M vs M-1)",
                height=360,
                margin=dict(l=40, r=20, t=60, b=40),
            )
            st.plotly_chart(fig_ca, use_container_width=True)

        # ================= COMPARATIF SEMAINES DU MOIS (ANNEE PRECEDENTE) =================
        st.subheader("🗓️ Comparatif semaines du mois (vs même mois année précédente)")
        prev_year_start = cur_month_start.replace(year=cur_month_start.year - 1)
        prev_year_end = (pd.Timestamp(prev_year_start) + pd.offsets.MonthEnd(0)).date()

        rows_y = []
        for w_start, w_end in week_ranges:
            py_start = prev_year_start.replace(day=min(w_start.day, prev_year_end.day))
            py_end = prev_year_start.replace(day=min(w_end.day, prev_year_end.day))

            df_cur = df_compare_base[
                (df_compare_base["DATE_CREATE_DATE"] >= w_start)
                & (df_compare_base["DATE_CREATE_DATE"] <= w_end)
            ]
            df_prev_y = df_compare_base[
                (df_compare_base["DATE_CREATE_DATE"] >= py_start)
                & (df_compare_base["DATE_CREATE_DATE"] <= py_end)
            ]

            k_cur = compute_kpis(df_cur, objectif_annuel)
            k_prev_y = compute_kpis(df_prev_y, objectif_annuel)

            rows_y.append({
                "Semaine": f"{w_start.strftime('%d/%m')}–{w_end.strftime('%d/%m')}",
                "Contrats M": k_cur["souscriptions"],
                "Contrats M-1": k_prev_y["souscriptions"],
                "Δ Contrats": k_cur["souscriptions"] - k_prev_y["souscriptions"],
                "Renouv. M": k_cur["renouvellements"],
                "Renouv. M-1": k_prev_y["renouvellements"],
                "Δ Renouv.": k_cur["renouvellements"] - k_prev_y["renouvellements"],
                "Nouv. aff. M": k_cur["nouvelles_affaires"],
                "Nouv. aff. M-1": k_prev_y["nouvelles_affaires"],
                "Δ Nouv.": k_cur["nouvelles_affaires"] - k_prev_y["nouvelles_affaires"],
                "CA M": k_cur["ca"],
                "CA M-1": k_prev_y["ca"],
                "Δ CA": k_cur["ca"] - k_prev_y["ca"],
            })

        cmp_month_week_y = pd.DataFrame(rows_y)
        num_cols_y = [c for c in cmp_month_week_y.columns if c != "Semaine"]
        total_row_y = {"Semaine": "Total"}
        total_row_y.update(cmp_month_week_y[num_cols_y].sum().to_dict())
        cmp_month_week_y = pd.concat([cmp_month_week_y, pd.DataFrame([total_row_y])], ignore_index=True)
        total_idx_y = len(cmp_month_week_y) - 1
        month_week_label_y = (
            f"Mois courant : {cur_month_start.strftime('%m/%Y')} | "
            f"Année précédente : {prev_year_start.strftime('%m/%Y')}"
        )
        st.caption(month_week_label_y)
        st.dataframe(
            cmp_month_week_y.style.format({
                "Contrats M": fmt_space_int,
                "Contrats M-1": fmt_space_int,
                "Δ Contrats": fmt_space_int,
                "Renouv. M": fmt_space_int,
                "Renouv. M-1": fmt_space_int,
                "Δ Renouv.": fmt_space_int,
                "Nouv. aff. M": fmt_space_int,
                "Nouv. aff. M-1": fmt_space_int,
                "Δ Nouv.": fmt_space_int,
                "CA M": fmt_space_int,
                "CA M-1": fmt_space_int,
                "Δ CA": fmt_space_int,
            }).set_table_styles([
                {"selector": "thead th", "props": [
                    ("background-color", "#F4D35E"),
                    ("color", "#1b2a41"),
                    ("font-weight", "700"),
                    ("text-align", "center"),
                ]},
                {"selector": "tbody td", "props": [
                    ("border-bottom", "1px solid #e6e8ec"),
                    ("padding", "8px"),
                    ("font-size", "13px"),
                ]},
                {"selector": "tbody td:not(:first-child)", "props": [
                    ("text-align", "right"),
                ]},
            ]).apply(_style_cmp_row, axis=1),
            use_container_width=True,
            hide_index=True,
        )
        c3, c4 = st.columns(2)
        with c3:
            fig_ren = go.Figure()
            fig_ren.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["Renouv. M"],
                name="Renouv. M",
            ))
            fig_ren.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["Renouv. M-1"],
                name="Renouv. M-1",
            ))
            fig_ren.update_layout(
                barmode="group",
                title="Renouvellements par semaine (M vs M-1)",
                height=360,
                margin=dict(l=40, r=20, t=60, b=40),
            )
            st.plotly_chart(fig_ren, use_container_width=True)

        with c4:
            fig_new = go.Figure()
            fig_new.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["Nouv. aff. M"],
                name="Nouv. aff. M",
            ))
            fig_new.add_trace(go.Bar(
                x=cmp_month_week["Semaine"],
                y=cmp_month_week["Nouv. aff. M-1"],
                name="Nouv. aff. M-1",
            ))
            fig_new.update_layout(
                barmode="group",
                title="Nouvelles affaires par semaine (M vs M-1)",
                height=360,
                margin=dict(l=40, r=20, t=60, b=40),
            )
            st.plotly_chart(fig_new, use_container_width=True)

        # ================= COMPARATIF PERIODE DU MOIS =================
        st.subheader("📌 Comparatif période du mois (même période année passée)")

        cur_start = start_d
        cur_end = end_d
        prev_y_start = (pd.Timestamp(cur_start) - pd.DateOffset(years=1)).date()
        prev_y_end = (pd.Timestamp(cur_end) - pd.DateOffset(years=1)).date()

        def _slice_kpis(_start, _end):
            _df = df_compare_base[
                (df_compare_base["DATE_CREATE_DATE"] >= _start)
                & (df_compare_base["DATE_CREATE_DATE"] <= _end)
            ]
            return compute_kpis(_df, objectif_annuel)

        k_m = _slice_kpis(cur_start, cur_end)
        k_py = _slice_kpis(prev_y_start, prev_y_end)

        metrics_period = [
            ("Nombre de contrats", "souscriptions"),
            ("Renouvellements", "renouvellements"),
            ("Nouvelles affaires", "nouvelles_affaires"),
            ("Chiffre d'affaires", "ca"),
        ]

        rows_period = []
        for label, key in metrics_period:
            m_val = float(k_m.get(key, 0))
            py_val = float(k_py.get(key, 0))
            rows_period.append({
                "Indicateur": label,
                "Période M": m_val,
                "Période N-1": py_val,
                "Δ M vs N-1": m_val - py_val,
            })

        comp_period_month = pd.DataFrame(rows_period)
        comp_period_label = (
            f"M: {cur_start.strftime('%d/%m/%Y')}–{cur_end.strftime('%d/%m/%Y')} | "
            f"N-1: {prev_y_start.strftime('%d/%m/%Y')}–{prev_y_end.strftime('%d/%m/%Y')}"
        )
        st.caption(comp_period_label)

        st.dataframe(
            comp_period_month.style.format({
                "Période M": fmt_space_int,
                "Période N-1": fmt_space_int,
                "Δ M vs N-1": fmt_space_int,
            }).set_table_styles([
                {"selector": "thead th", "props": [
                    ("background-color", "#F4D35E"),
                    ("color", "#1b2a41"),
                    ("font-weight", "700"),
                    ("text-align", "center"),
                ]},
                {"selector": "tbody td", "props": [
                    ("border-bottom", "1px solid #e6e8ec"),
                    ("padding", "8px"),
                    ("font-size", "13px"),
                ]},
                {"selector": "tbody td:not(:first-child)", "props": [
                    ("text-align", "right"),
                ]},
            ]).apply(
                lambda row: [
                    "background-color:#fff3cd" if str(col).startswith("Δ") else ""
                    for col in row.index
                ],
                axis=1,
            ),
            use_container_width=True,
            hide_index=True,
        )

        # ================= COMPARATIF POOL vs HORS POOL (N-1) =================
        st.subheader("🧮 Comparatif POOL vs Hors pool (même période année passée)")
        pool_col_cmp = _find_col(df_compare_base, ["POOLS_TPV", "POOL_TPV", "POOL", "CESSION_POOL_TPV_40"])
        if pool_col_cmp and pool_col_cmp in df_compare_base.columns:
            pool_raw_cmp = df_compare_base[pool_col_cmp]
            pool_num_cmp = _to_num(pool_raw_cmp).fillna(0)
            pool_txt_cmp = pool_raw_cmp.astype("string").str.upper().str.strip()
            mask_pool_cmp = (pool_num_cmp > 0) | pool_txt_cmp.isin(["TRUE", "VRAI", "OUI", "YES"])
        else:
            mask_pool_cmp = pd.Series([False] * len(df_compare_base), index=df_compare_base.index)

        def _slice_pool_stats(mask_kind: pd.Series, _start, _end):
            _df = df_compare_base[
                mask_kind
                & (df_compare_base["DATE_CREATE_DATE"] >= _start)
                & (df_compare_base["DATE_CREATE_DATE"] <= _end)
            ].copy()
            return compute_kpis(_df, objectif_annuel)

        k_pool_m = _slice_pool_stats(mask_pool_cmp, cur_start, cur_end)
        k_pool_py = _slice_pool_stats(mask_pool_cmp, prev_y_start, prev_y_end)
        k_hors_m = _slice_pool_stats(~mask_pool_cmp, cur_start, cur_end)
        k_hors_py = _slice_pool_stats(~mask_pool_cmp, prev_y_start, prev_y_end)

        comp_pool_period_year = pd.DataFrame(
            [
                {
                    "Type": "POOL",
                    "Contrats M": k_pool_m["souscriptions"],
                    "Contrats N-1": k_pool_py["souscriptions"],
                    "Δ Contrats": k_pool_m["souscriptions"] - k_pool_py["souscriptions"],
                    "Renouv. M": k_pool_m["renouvellements"],
                    "Renouv. N-1": k_pool_py["renouvellements"],
                    "Δ Renouv.": k_pool_m["renouvellements"] - k_pool_py["renouvellements"],
                    "Nouv. aff. M": k_pool_m["nouvelles_affaires"],
                    "Nouv. aff. N-1": k_pool_py["nouvelles_affaires"],
                    "Δ Nouv.": k_pool_m["nouvelles_affaires"] - k_pool_py["nouvelles_affaires"],
                    "CA M": k_pool_m["ca"],
                    "CA N-1": k_pool_py["ca"],
                    "Δ CA": k_pool_m["ca"] - k_pool_py["ca"],
                },
                {
                    "Type": "Hors pool",
                    "Contrats M": k_hors_m["souscriptions"],
                    "Contrats N-1": k_hors_py["souscriptions"],
                    "Δ Contrats": k_hors_m["souscriptions"] - k_hors_py["souscriptions"],
                    "Renouv. M": k_hors_m["renouvellements"],
                    "Renouv. N-1": k_hors_py["renouvellements"],
                    "Δ Renouv.": k_hors_m["renouvellements"] - k_hors_py["renouvellements"],
                    "Nouv. aff. M": k_hors_m["nouvelles_affaires"],
                    "Nouv. aff. N-1": k_hors_py["nouvelles_affaires"],
                    "Δ Nouv.": k_hors_m["nouvelles_affaires"] - k_hors_py["nouvelles_affaires"],
                    "CA M": k_hors_m["ca"],
                    "CA N-1": k_hors_py["ca"],
                    "Δ CA": k_hors_m["ca"] - k_hors_py["ca"],
                },
            ]
        )

        total_pool_row = {
            "Type": "Total",
            "Contrats M": int(comp_pool_period_year["Contrats M"].sum()),
            "Contrats N-1": int(comp_pool_period_year["Contrats N-1"].sum()),
            "Δ Contrats": int(comp_pool_period_year["Δ Contrats"].sum()),
            "Renouv. M": int(comp_pool_period_year["Renouv. M"].sum()),
            "Renouv. N-1": int(comp_pool_period_year["Renouv. N-1"].sum()),
            "Δ Renouv.": int(comp_pool_period_year["Δ Renouv."].sum()),
            "Nouv. aff. M": int(comp_pool_period_year["Nouv. aff. M"].sum()),
            "Nouv. aff. N-1": int(comp_pool_period_year["Nouv. aff. N-1"].sum()),
            "Δ Nouv.": int(comp_pool_period_year["Δ Nouv."].sum()),
            "CA M": float(comp_pool_period_year["CA M"].sum()),
            "CA N-1": float(comp_pool_period_year["CA N-1"].sum()),
            "Δ CA": float(comp_pool_period_year["Δ CA"].sum()),
        }
        comp_pool_period_year = pd.concat([comp_pool_period_year, pd.DataFrame([total_pool_row])], ignore_index=True)
        comp_pool_period_label = comp_period_label

        def _style_pool_row(row):
            styles = []
            is_total = str(row.get("Type", "")).strip().lower() == "total"
            for col in row.index:
                if is_total:
                    styles.append("font-weight:700; color:#000; background-color:#e9ecef")
                elif str(col).startswith("Δ"):
                    styles.append("background-color:#fff3cd")
                else:
                    styles.append("")
            return styles

        st.caption(comp_pool_period_label)
        st.dataframe(
            comp_pool_period_year.style.format(
                {
                    "Contrats M": fmt_space_int,
                    "Contrats N-1": fmt_space_int,
                    "Δ Contrats": fmt_space_int,
                    "Renouv. M": fmt_space_int,
                    "Renouv. N-1": fmt_space_int,
                    "Δ Renouv.": fmt_space_int,
                    "Nouv. aff. M": fmt_space_int,
                    "Nouv. aff. N-1": fmt_space_int,
                    "Δ Nouv.": fmt_space_int,
                    "CA M": fmt_space_int,
                    "CA N-1": fmt_space_int,
                    "Δ CA": fmt_space_int,
                }
            ).set_table_styles(
                [
                    {"selector": "thead th", "props": [
                        ("background-color", "#F4D35E"),
                        ("color", "#1b2a41"),
                        ("font-weight", "700"),
                        ("text-align", "center"),
                    ]},
                    {"selector": "tbody td", "props": [
                        ("border-bottom", "1px solid #e6e8ec"),
                        ("padding", "8px"),
                        ("font-size", "13px"),
                    ]},
                    {"selector": "tbody td:not(:first-child)", "props": [
                        ("text-align", "right"),
                    ]},
                ]
            ).apply(_style_pool_row, axis=1),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Colonne DATE_CREATE manquante pour le comparatif mensuel.")

    st.divider()

    # ================= REPARTITION PAR PRODUIT =================
    st.subheader("📊 RÉPARTITION PAR PRODUIT")

    tab_prod = tableau_produit(df_f)

    tabp = tab_prod.copy()
    tabp["Nombre de contrats"] = pd.to_numeric(tabp["Nombre de contrats"], errors="coerce").fillna(0)
    tabp["Prime TTC (FCFA)"] = pd.to_numeric(tabp["Prime TTC (FCFA)"], errors="coerce").fillna(0)
    tabp["% Contrats"] = pd.to_numeric(tabp["% Contrats"], errors="coerce").fillna(0)
    tabp["% Prime TTC"] = pd.to_numeric(tabp["% Prime TTC"], errors="coerce").fillna(0)
    tabp = tabp[["PRODUIT", "Nombre de contrats", "% Contrats", "Prime TTC (FCFA)", "% Prime TTC"]]

    def _style_prod_rows(row):
        is_total = str(row["PRODUIT"]).strip().lower() == "total"
        if is_total:
            return ["background-color:#1F4E79;color:white;font-weight:bold;"] * len(row)
        return ["background-color:#FFF9E8;"] * len(row)

    styler_prod = (
        tabp.style
        .format({
            "Nombre de contrats": lambda x: fmt_space_int(x),
            "Prime TTC (FCFA)": lambda x: fmt_space_int(x),
            "% Contrats": lambda x: f"{x*100:.1f}%",
            "% Prime TTC": lambda x: f"{x*100:.1f}%",
        })
        .apply(_style_prod_rows, axis=1)
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", "#F4D35E"),
                ("color", "black"),
                ("font-weight", "bold"),
                ("text-align", "center"),
                ("border", "1px solid #cfcfcf"),
                ("padding", "10px"),
            ]},
            {"selector": "tbody td", "props": [
                ("border", "1px solid #d9d9d9"),
                ("padding", "10px"),
                ("text-align", "center"),
                ("font-size", "14px"),
            ]},
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "left"),
                ("font-weight", "bold"),
            ]},
        ])
    )

    st.dataframe(styler_prod, use_container_width=True, hide_index=True)

    # Graphique produit interactif
    df_plot = tabp[tabp["PRODUIT"].astype(str).str.lower() != "total"].copy()
    df_plot["Prime_M"] = df_plot["Prime TTC (FCFA)"] / 1_000_000
    df_plot["PctPrime"] = df_plot["% Prime TTC"] * 100

    fig_prod = px.bar(
        df_plot,
        x="PRODUIT",
        y="Prime_M",
        text=df_plot["PctPrime"].map(lambda v: f"{v:.1f}%"),
        labels={"Prime_M": "Chiffre d'affaires (Millions FCFA)", "PRODUIT": "Produit"},
        title="Répartition du Chiffre d’Affaires par Produit",
    )
    fig_prod.update_traces(textposition="outside")
    fig_prod.update_layout(height=420, yaxis=dict(ticksuffix=" M"), margin=dict(l=40, r=30, t=70, b=40))
    st.plotly_chart(fig_prod, use_container_width=True)

    st.divider()

    # ================= CA MENSUEL =================
    st.subheader("📅 CHIFFRE D’AFFAIRES MENSUEL")

    tab_month = monthly_ca_table(df_f, objectif_annuel)
    order_fr = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
    month_map = {i + 1: m for i, m in enumerate(order_fr)}
    months_present = []
    if "DATE_CREATE" in df_f.columns:
        months_present = (
            df_f["DATE_CREATE"]
            .dropna()
            .dt.month
            .unique()
            .tolist()
        )
        months_present = [month_map[m] for m in sorted(months_present) if m in month_map]

    if months_present:
        tab_month = tab_month[
            tab_month["Mois"].astype(str).isin(months_present)
            | tab_month["Mois"].astype(str).str.lower().str.startswith("total")
        ].copy()

    tabm = tab_month.copy()
    tabm.insert(0, " ", range(len(tabm)))

    def _style_month_row(row):
        mois = str(row["Mois"]).strip().lower()
        is_total = mois.startswith("total")
        styles = []
        for col in row.index:
            if col == " ":
                styles.append("background-color:#F4D35E;font-weight:bold;")
            elif is_total:
                styles.append("background-color:#FFF1C6;font-weight:bold;")
            else:
                styles.append("")
        return styles

    styler_month = (
        tabm.style
        .format({
            "CA par Mois": lambda x: fmt_space_int(x),
            "Objectif Mensuel": lambda x: fmt_space_int(x),
            "Ratio Objectif Mensuel": lambda x: fmt_pct_1(x),
        })
        .apply(_style_month_row, axis=1)
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", "#F4D35E"),
                ("color", "black"),
                ("font-weight", "bold"),
                ("text-align", "center"),
                ("border", "1px solid #cfcfcf"),
                ("padding", "10px"),
            ]},
            {"selector": "tbody td", "props": [
                ("border", "1px solid #e0e0e0"),
                ("padding", "10px"),
                ("text-align", "center"),
                ("font-size", "14px"),
            ]},
            {"selector": "tbody td:nth-child(2)", "props": [("text-align", "left"), ("font-weight", "bold")]},
        ])
    )
    st.dataframe(styler_month, use_container_width=True, hide_index=True)

    # --- Données graphes (sans Total) ---
    order_fr_used = [m for m in order_fr if m in months_present] if months_present else order_fr
    dfm = tab_month.copy()
    dfm = dfm[~dfm["Mois"].astype(str).str.lower().str.startswith("total")].copy()

    dfm["CA par Mois"] = pd.to_numeric(dfm["CA par Mois"], errors="coerce").fillna(0)
    dfm["Objectif Mensuel"] = pd.to_numeric(dfm["Objectif Mensuel"], errors="coerce").fillna(0)
    dfm["Ratio Objectif Mensuel"] = pd.to_numeric(dfm["Ratio Objectif Mensuel"], errors="coerce").fillna(0)

    dfm["CA_M"] = dfm["CA par Mois"] / 1_000_000
    dfm["OBJ_M"] = dfm["Objectif Mensuel"] / 1_000_000
    dfm["RATIO_PCT"] = dfm["Ratio Objectif Mensuel"] * 100

    st.markdown("#### ✅ Graphique 1 — CA mensuel vs Objectif + Ratio (%)")
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        x=dfm["Mois"], y=dfm["CA_M"], name="CA mensuel",
        hovertemplate="<b>%{x}</b><br>CA: %{customdata} FCFA<br><extra></extra>",
        customdata=dfm["CA par Mois"].map(lambda v: fmt_space_int(v)),
    ))
    fig1.add_trace(go.Scatter(
        x=dfm["Mois"], y=dfm["OBJ_M"], mode="lines+markers", name="Objectif mensuel",
        hovertemplate="<b>%{x}</b><br>Objectif: %{customdata} FCFA<br><extra></extra>",
        customdata=dfm["Objectif Mensuel"].map(lambda v: fmt_space_int(v)),
    ))
    fig1.add_trace(go.Scatter(
        x=dfm["Mois"], y=dfm["CA_M"], mode="text",
        text=dfm["RATIO_PCT"].map(lambda v: f"{v:.1f}%"),
        textposition="top center", showlegend=False
    ))
    title_months = (
        f"{order_fr_used[0][:3]} → {order_fr_used[-1][:3]}"
        if order_fr_used else "Mois disponibles"
    )
    fig1.update_layout(
        title=f"Chiffre d’affaires mensuel — {title_months}",
        xaxis=dict(categoryorder="array", categoryarray=order_fr_used),
        yaxis_title="Montant (Millions FCFA)",
        height=460,
        margin=dict(l=40, r=30, t=70, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig1, use_container_width=True)

    st.markdown("#### ✅ Graphique 2 — Évolution mensuelle du CA (valeurs FCFA sur les barres)")
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=dfm["Mois"], y=dfm["CA_M"], name="Chiffre d'affaires",
        text=dfm["CA par Mois"].map(lambda v: fmt_space_int(v)),
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>CA: %{customdata} FCFA<br><extra></extra>",
        customdata=dfm["CA par Mois"].map(lambda v: fmt_space_int(v)),
    ))
    fig2.update_layout(
        title=f"Évolution mensuelle du chiffre d’affaires (Janvier au {end_date_str})",
        xaxis=dict(categoryorder="array", categoryarray=order_fr_used),
        yaxis_title="Chiffre d'affaires (Millions FCFA)",
        height=460,
        margin=dict(l=40, r=30, t=70, b=40),
    )
    st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # ================= CA MOYEN JOURNALIER =================
    st.subheader("📆 CA MOYEN JOURNALIER PAR JOUR DE LA SEMAINE")

    tab_week = weekly_avg_table(df_f)

    order_days = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"]
    cols_week = ["Mois"] + order_days + ["Total"]

    tabw = tab_week.copy()
    for c in cols_week:
        if c not in tabw.columns:
            tabw[c] = 0
    tabw = tabw[cols_week].copy()

    for c in order_days + ["Total"]:
        tabw[c] = pd.to_numeric(tabw[c], errors="coerce").fillna(0)

    def _style_week_rows(row):
        is_total = str(row["Mois"]).strip().lower() == "total"
        if is_total:
            return ["background-color:#1F4E79;color:white;font-weight:bold;"] * len(row)
        return [""] * len(row)

    styler_week = (
        tabw.style
        .format({c: (lambda x: fmt_space_int(x)) for c in order_days + ["Total"]})
        .apply(_style_week_rows, axis=1)
        .background_gradient(
            subset=pd.IndexSlice[tabw["Mois"].astype(str).str.lower() != "total", order_days + ["Total"]],
            cmap="Blues",
            axis=None
        )
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", "#F4D35E"),
                ("color", "black"),
                ("font-weight", "bold"),
                ("text-align", "center"),
                ("border", "1px solid #cfcfcf"),
                ("padding", "10px"),
                ("font-size", "14px"),
            ]},
            {"selector": "tbody td", "props": [
                ("border", "1px solid #e0e0e0"),
                ("padding", "10px"),
                ("text-align", "right"),
                ("font-size", "14px"),
            ]},
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "left"),
                ("font-weight", "bold"),
            ]},
        ])
    )

    st.dataframe(styler_week, use_container_width=True, hide_index=True)

    # Graphique interactif (ligne)
    dfw2 = tabw.copy()
    dfw2 = dfw2[~dfw2["Mois"].astype(str).str.strip().str.lower().eq("total")].copy()

    df_long = dfw2.melt(
        id_vars=["Mois"],
        value_vars=[d for d in order_days if d in dfw2.columns],
        var_name="Jour",
        value_name="CA"
    )
    df_long["CA"] = pd.to_numeric(df_long["CA"], errors="coerce").fillna(0)
    df_long["Jour"] = pd.Categorical(df_long["Jour"], categories=order_days, ordered=True)
    df_long = df_long.sort_values(["Jour", "Mois"])

    fig_week = px.line(
        df_long,
        x="Jour",
        y="CA",
        color="Mois",
        markers=True,
        title="CA moyen journalier par jour de la semaine (comparaison mensuelle)"
    )
    fig_week.update_layout(
        height=420,
        yaxis_title="CA moyen journalier (FCFA)",
        xaxis_title="Jour de la semaine",
        legend_title_text="Mois",
    )
    st.plotly_chart(fig_week, use_container_width=True)

    st.divider()

    # ================= EXPORT PDF =================
    st.sidebar.header("⬇️ Rapport PDF")

    filters_key = f"{segment_label}|{inter_label}|{start_date_str}|{end_date_str}"

    if st.sidebar.button("🧾 Générer PDF", use_container_width=True, key="btn_pdf"):
        if st.session_state.get("pdf_key") == filters_key and st.session_state.get("pdf_bytes") is not None:
            st.sidebar.success("PDF déjà généré pour ces filtres ✅")
        else:
            with st.spinner("⏳ Génération du PDF en cours..."):
                k_pdf = dict(k)
                k_pdf.update(
                    {
                        "nb_contrats_sans_moto": nb_contrats_sans_moto,
                        "ren_non_moto": ren_non_moto,
                        "new_non_moto": new_non_moto,
                        "pct_ren_non_moto": pct_ren_non_moto,
                        "pct_new_non_moto": pct_new_non_moto,
                        "montant_reverser_pool": montant_reverser_pool,
                        "pct_reverser_pool_ca": pct_reverser_pool_ca,
                        "pct_1_month_non_moto": pct_1_month_non_moto if "pct_1_month_non_moto" in locals() else 0.0,
                        "ca_evol_text": ca_evol_pdf,
                        "ca_evol_up": ca_evol_up,
                        "ca_evol_down": ca_evol_down,
                        "pool_pct_text": pool_evol_pdf,
                        "pool_evol_up": pool_evol_up,
                        "pool_evol_down": pool_evol_down,
                        "hors_pool_pct_text": (
                            f"{fmt_pct_1(pct_hors_vs_ca)} | {fmt_pct_1(pct_hors_share)} | {hors_evol_pdf}"
                        ),
                        "hors_evol_up": hors_evol_up,
                        "hors_evol_down": hors_evol_down,
                        "one_month_count_pct_text": (
                            f"{fmt_pct_1(pct_1_month_non_moto)} | {fmt_pct_1(pct_1_month_all)} | {one_month_count_evol_pdf}"
                        ),
                        "one_month_count_evol_up": one_month_count_evol > 0,
                        "one_month_count_evol_down": one_month_count_evol < 0,
                        "one_month_pct_text": (
                            f"{fmt_pct_1(pct_amt_1m_vs_ca)}\n| {one_month_amt_evol_pdf}"
                        ),
                        "one_month_evol_up": one_month_amt_evol > 0,
                        "one_month_evol_down": one_month_amt_evol < 0,
                    }
                )

                st.session_state.pdf_bytes = export_pdf_dashboard(
                    k=k_pdf,
                    segment=segment_label,
                    intermediaire=inter_label,
                    start_date=start_date_str,
                    end_date=end_date_str,
                    objectif_annuel=objectif_annuel,
                    tab_produit=tab_prod,
                    tab_month=tab_month,
                    tab_week=tab_week,  # ✅ pour page 5 dans le PDF
                    avg_contracts_daily=avg_contracts,
                    avg_ca_daily=avg_ca,
                    ca_1_month=ca_1_month,
                    nb_1_month=nb_1_month,
                    comp_week=comp_week_df,
                    week_n_label=week_n_label,
                    week_n1_label=week_n1_label,
                    comp_month_week=cmp_month_week,
                    month_week_label=month_week_label,
                    comp_month_week_y=cmp_month_week_y,
                    month_week_label_y=month_week_label_y,
                    comp_period_month=comp_period_month,
                    comp_period_label=comp_period_label,
                    comp_pool_period_year=comp_pool_period_year,
                    comp_pool_period_label=comp_pool_period_label,
                    one_month_feb_table=tb_1m_feb,
                    one_month_feb_label=f"Contrats 1 mois arrivant à échéance en février {target_year_feb}",
                    top_agents_ca=top_agents_ca,
                    top_agents_cnt=top_agents_cnt,
                    top_pool_ca=top_pool_ca,
                    top_pool_cnt=top_pool_cnt,
                )
                st.session_state.pdf_key = filters_key

    if st.session_state.get("pdf_bytes") is not None and st.session_state.get("pdf_key") == filters_key:
        pdf_name = (
            f"RAPPORT_SUIVI_AUTO_{segment_label}_{inter_label}_{start_date_str}_au_{end_date_str}.pdf"
            .replace(" ", "_")
        )
        st.sidebar.download_button(
            "📄 Télécharger le rapport PDF",
            data=st.session_state.pdf_bytes,
            file_name=pdf_name,
            mime="application/pdf",
            use_container_width=True,
            key="dl_pdf",
        )
    else:
        st.sidebar.caption("Clique sur « Générer PDF » pour produire le rapport avec les filtres actuels.")
