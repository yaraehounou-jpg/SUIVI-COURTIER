# SUIVI_AUTO/tasks/suivi_hebdomadaire_auto.py
from __future__ import annotations

from datetime import datetime, timedelta, date
import pandas as pd
import polars as pl
import streamlit as st
import plotly.express as px

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

OBJECTIF_AGENT = 4_826_339_000
OBJECTIF_COURTIER = 2_000_000_000


def fmt_space_int(x) -> str:
    return fmt_int(x)


def fmt_pct_1(x) -> str:
    try:
        return f"{float(x) * 100:.1f}%"
    except Exception:
        return "0.0%"


def get_objectif(segment: str) -> int:
    if segment == "Agent":
        return OBJECTIF_AGENT
    if segment == "Courtier":
        return OBJECTIF_COURTIER
    return OBJECTIF_AGENT + OBJECTIF_COURTIER


def reset_exports():
    st.session_state.pdf_bytes = None
    st.session_state.pdf_key = None


def kpi_card(value: str, label: str, pct: str | None = None, kind: str = "default") -> str:
    pct_html = f"<div class='kpi_pct'>{pct}</div>" if pct else ""
    return f"""
    <div class="kpi_card {kind}">
      <div class="kpi_value">{value}</div>
      <div class="kpi_label">{label}</div>
      {pct_html}
    </div>
    """


def ratio_card(ratio: str) -> str:
    return f"""
    <div class="kpi_card ratio">
      <div class="kpi_ratio_value">{ratio}</div>
      <div class="kpi_ratio_label">Ratio Objectif Annuel</div>
    </div>
    """


# ---------------- Comparatif helpers ----------------
def _safe_div(a: float, b: float) -> float:
    try:
        a = float(a)
        b = float(b)
        return a / b if b != 0 else 0.0
    except Exception:
        return 0.0


def _make_comp_df(k_cur: dict, k_ref: dict, label_cur: str, label_ref: str) -> pd.DataFrame:
    rows = []

    def add_row(name: str, cur_key: str, ref_key: str | None = None):
        ref_key = ref_key or cur_key
        cur = float(k_cur.get(cur_key, 0) or 0)
        ref = float(k_ref.get(ref_key, 0) or 0)
        delta = cur - ref
        pct = _safe_div(delta, ref)
        rows.append([name, cur, ref, delta, pct])

    add_row("CA (Prime TTC)", "ca")
    add_row("Souscriptions", "souscriptions")
    add_row("Renouvellements", "renouvellements")
    add_row("Nouvelles affaires", "nouvelles_affaires")
    add_row("POOL (CA)", "montant_pool")
    add_row("Hors POOL (CA)", "montant_hors_pool")

    df = pd.DataFrame(rows, columns=["Indicateur", label_cur, label_ref, "Δ (valeur)", "Δ (%)"])
    df = df.rename(columns={label_cur: "Période actuelle", label_ref: "Période référence"})
    return df


def _filter_like_current(df_base: pd.DataFrame, segment_choice: str, inter_choice: str) -> pd.DataFrame:
    out = df_base.copy()
    if segment_choice != "Tous":
        out = out[out["SEGMENT"] == segment_choice].copy()
    if inter_choice != "Tous":
        out = out[out["INTERMEDIAIRE"] == inter_choice].copy()
    return out


def _slice_dates(df_in: pd.DataFrame, start_d: date, end_d: date) -> pd.DataFrame:
    if df_in is None or df_in.empty:
        return df_in
    if "DATE_CREATE" not in df_in.columns:
        return df_in.iloc[0:0].copy()

    d = df_in.copy()
    d["DATE_CREATE"] = pd.to_datetime(d["DATE_CREATE"], errors="coerce")
    d = d.dropna(subset=["DATE_CREATE"]).copy()
    if d.empty:
        return d

    return d[(d["DATE_CREATE"].dt.date >= start_d) & (d["DATE_CREATE"].dt.date <= end_d)].copy()


def _prev_same_duration(start_d: date, end_d: date) -> tuple[date, date]:
    n = (end_d - start_d).days + 1
    prev_end = start_d - timedelta(days=1)
    prev_start = prev_end - timedelta(days=max(n - 1, 0))
    return prev_start, prev_end


def _shift_month(d: date, months: int) -> date:
    ts = pd.Timestamp(d)
    return (ts + pd.DateOffset(months=months)).date()


def _shift_year(d: date, years: int) -> date:
    ts = pd.Timestamp(d)
    return (ts + pd.DateOffset(years=years)).date()


def _month_start(d: date) -> date:
    return date(d.year, d.month, 1)


# ---------------- TOP 10 helpers (POOL/HORS POOL/TOTAL) ----------------
def _policy_col(df: pd.DataFrame) -> str | None:
    for c in ["POLICYNO", "POLICY_NO", "POLICYNUMBER", "POLICY_NUMBER", "POLICY"]:
        if c in df.columns:
            return c
    return None


def top_intermediaires_table(df_in: pd.DataFrame, n: int = 10, pool_value: int | None = None) -> pd.DataFrame:
    """
    Top N par CA (PRIME_TTC), groupé par INTERMEDIAIRE.
    - pool_value=None => Total
    - pool_value=1 => POOL (POOLS_TPV = 1)
    - pool_value=0 => Hors POOL (POOLS_TPV = 0)
    """
    cols_out = ["Rang", "Intermédiaire", "CA (FCFA)", "Souscriptions", "% CA"]
    if df_in is None or df_in.empty:
        return pd.DataFrame(columns=cols_out)

    df = df_in.copy()

    # sécurité colonnes
    if "PRIME_TTC" not in df.columns:
        df["PRIME_TTC"] = 0
    df["PRIME_TTC"] = pd.to_numeric(df["PRIME_TTC"], errors="coerce").fillna(0)

    if "POOLS_TPV" not in df.columns:
        df["POOLS_TPV"] = 0
    df["POOLS_TPV"] = pd.to_numeric(df["POOLS_TPV"], errors="coerce").fillna(0).astype(int)

    if "INTERMEDIAIRE" not in df.columns:
        return pd.DataFrame(columns=cols_out)

    # filtre pool/hors pool
    if pool_value in (0, 1):
        df = df[df["POOLS_TPV"] == int(pool_value)].copy()

    df = df.dropna(subset=["INTERMEDIAIRE"]).copy()
    df["INTERMEDIAIRE"] = df["INTERMEDIAIRE"].astype(str).str.strip()
    df = df[df["INTERMEDIAIRE"] != ""].copy()

    if df.empty:
        return pd.DataFrame(columns=cols_out)

    pol = _policy_col(df)

    if pol:
        g = df.groupby("INTERMEDIAIRE", as_index=False).agg(
            CA_FCFA=("PRIME_TTC", "sum"),
            Souscriptions=(pol, "nunique"),
        )
    else:
        g = df.groupby("INTERMEDIAIRE", as_index=False).agg(
            CA_FCFA=("PRIME_TTC", "sum"),
            Souscriptions=("PRIME_TTC", "size"),
        )

    g = g.sort_values("CA_FCFA", ascending=False).head(n).reset_index(drop=True)

    total_ca = float(g["CA_FCFA"].sum()) if not g.empty else 0.0
    g["% CA"] = (g["CA_FCFA"] / total_ca) if total_ca else 0.0

    g.insert(0, "Rang", range(1, len(g) + 1))
    g = g.rename(columns={"INTERMEDIAIRE": "Intermédiaire", "CA_FCFA": "CA (FCFA)"})
    return g[cols_out]


def _maybe_styler(df_top: pd.DataFrame):
    """
    Streamlit récent: OK avec Styler
    Streamlit ancien: parfois bug -> on renvoie df normal
    """
    if df_top is None or df_top.empty:
        return df_top
    try:
        return (
            df_top.style.format(
                {
                    "CA (FCFA)": lambda x: fmt_space_int(x),
                    "% CA": lambda x: f"{float(x) * 100:.1f}%",
                }
            )
        )
    except Exception:
        return df_top


def top10_bar_chart(df_top: pd.DataFrame, title: str):
    if df_top is None or df_top.empty:
        st.info("Pas assez de données pour le Top 10.")
        return

    d = df_top.copy()
    d["CA_M"] = pd.to_numeric(d["CA (FCFA)"], errors="coerce").fillna(0) / 1_000_000

    fig = px.bar(
        d,
        x="Intermédiaire",
        y="CA_M",
        text=d["CA (FCFA)"].map(fmt_space_int),
        labels={"CA_M": "CA (Millions FCFA)", "Intermédiaire": "Intermédiaire"},
        title=title,
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(height=460, yaxis=dict(ticksuffix=" M"), margin=dict(l=40, r=30, t=70, b=90))
    fig.update_xaxes(tickangle=25)
    st.plotly_chart(fig, use_container_width=True)


def run() -> None:
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
            df0 = normalize(df_raw)
            df0 = add_segment(df0)
            df0 = add_intermediaire(df0)
            st.session_state.df_base = df0

    df = st.session_state.df_base
    if df is None or df.empty:
        st.info("Charge tes fichiers Excel dans la sidebar.")
        st.stop()

    if "DATE_CREATE" not in df.columns:
        st.error("Colonne DATE_CREATE introuvable après normalisation.")
        st.stop()

    # force datetime (sécurité)
    df = df.copy()
    df["DATE_CREATE"] = pd.to_datetime(df["DATE_CREATE"], errors="coerce")
    st.session_state.df_base = df

    df_dates = df.dropna(subset=["DATE_CREATE"]).copy()
    if df_dates.empty:
        st.error("Aucune date valide dans DATE_CREATE.")
        st.stop()

    # ---------------- Sidebar: Filters ----------------
    st.sidebar.header("🎛️ Filtres")

    min_date = df_dates["DATE_CREATE"].min().date()
    max_date = df_dates["DATE_CREATE"].max().date()

    segment_choice = st.sidebar.selectbox(
        "Segment",
        ["Tous", "Agent", "Courtier"],
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

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_d, end_d = date_range
    else:
        start_d, end_d = min_date, max_date

    df_seg = df.copy()
    if segment_choice != "Tous":
        df_seg = df_seg[df_seg["SEGMENT"] == segment_choice].copy()

    inter_list = sorted(df_seg["INTERMEDIAIRE"].dropna().astype(str).unique().tolist())
    inter_choice = st.sidebar.selectbox(
        "Intermédiaire",
        ["Tous"] + inter_list,
        index=0,
        key="inter_choice",
        on_change=reset_exports,
    )

    st.sidebar.button("🔄 Rafraîchir", on_click=st.rerun, use_container_width=True)

    # ---------------- Apply filters ----------------
    df_f = df.copy()

    if segment_choice != "Tous":
        df_f = df_f[df_f["SEGMENT"] == segment_choice].copy()

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

    # ---------------- Debug léger (utile si rien ne s’affiche) ----------------
    with st.expander("🔎 Debug (dates / pool)", expanded=False):
        st.write("DATE_CREATE dtype :", df_f["DATE_CREATE"].dtype)
        if "POOLS_TPV" in df_f.columns:
            u = sorted(pd.to_numeric(df_f["POOLS_TPV"], errors="coerce").fillna(0).astype(int).unique().tolist())
            st.write("POOLS_TPV uniques :", u)
            st.write("Nb POOL(1) :", int((pd.to_numeric(df_f["POOLS_TPV"], errors="coerce").fillna(0).astype(int) == 1).sum()))
            st.write("Nb Hors POOL(0) :", int((pd.to_numeric(df_f["POOLS_TPV"], errors="coerce").fillna(0).astype(int) == 0).sum()))
        else:
            st.warning("Colonne POOLS_TPV absente après normalize().")

    objectif_annuel = get_objectif(segment_choice)

    segment_label = segment_choice
    inter_label = inter_choice if inter_choice != "Tous" else (
        "Tous les agents" if segment_choice == "Agent"
        else "Tous les courtiers" if segment_choice == "Courtier"
        else "Tous les intermédiaires"
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

    row1 = st.columns(4)
    row2 = st.columns(3)
    row3 = st.columns(3)

    with row1[0]:
        st.markdown(kpi_card(fmt_space_int(k["taxes"]), "Impôts"), unsafe_allow_html=True)
    with row1[1]:
        st.markdown(kpi_card(fmt_space_int(k["ca"]), "Chiffre d'affaires", kind="highlight"), unsafe_allow_html=True)
    with row1[2]:
        st.markdown(kpi_card(fmt_space_int(k["nb_agents_fideles"]), "Agent fidèle"), unsafe_allow_html=True)
    with row1[3]:
        st.markdown(kpi_card(fmt_space_int(k["nb_agents_actifs"]), "Nom Agents Actifs"), unsafe_allow_html=True)

    with row2[0]:
        st.markdown(
            kpi_card(fmt_space_int(k["montant_pool"]), "Montant POOL TPV", pct=fmt_pct_1(k["pct_pool"]), kind="highlight"),
            unsafe_allow_html=True,
        )
    with row2[1]:
        st.markdown(ratio_card(fmt_pct_1(k["ratio_obj"])), unsafe_allow_html=True)
    with row2[2]:
        st.markdown(
            kpi_card(fmt_space_int(k["montant_hors_pool"]), "Montant hors pool", pct=fmt_pct_1(k["pct_hors_pool"]), kind="highlight"),
            unsafe_allow_html=True,
        )

    with row3[0]:
        st.markdown(
            kpi_card(fmt_space_int(k["renouvellements"]), "Renouvellement", pct=fmt_pct_1(k["pct_ren"]), kind="dark"),
            unsafe_allow_html=True,
        )
    with row3[1]:
        st.markdown(kpi_card(fmt_space_int(k["souscriptions"]), "Abonnements", kind="dark"), unsafe_allow_html=True)
    with row3[2]:
        st.markdown(
            kpi_card(fmt_space_int(k["nouvelles_affaires"]), "Nouvelles affaires", pct=fmt_pct_1(k["pct_new"]), kind="dark"),
            unsafe_allow_html=True,
        )

    st.markdown(
        f"<div class='objectif_footer'>Objectif annuel : <b>{fmt_space_int(objectif_annuel)}</b> FCFA</div>",
        unsafe_allow_html=True,
    )

    st.divider()

    # ================= TOP 10 (TOTAL / POOL / HORS POOL) =================
    st.subheader("🏆 TOP 10 INTERMÉDIAIRES — Total / POOL / Hors POOL")

    top_total = top_intermediaires_table(df_f, n=10, pool_value=None)
    top_pool = top_intermediaires_table(df_f, n=10, pool_value=1)
    top_hors = top_intermediaires_table(df_f, n=10, pool_value=0)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### ✅ Top 10 Total")
        st.dataframe(_maybe_styler(top_total), use_container_width=True, hide_index=True)
    with c2:
        st.markdown("### ✅ Top 10 POOL (POOLS_TPV = 1)")
        st.dataframe(_maybe_styler(top_pool), use_container_width=True, hide_index=True)
    with c3:
        st.markdown("### ✅ Top 10 Hors POOL (POOLS_TPV = 0)")
        st.dataframe(_maybe_styler(top_hors), use_container_width=True, hide_index=True)

    st.markdown("### 📊 Graphique — CA Top 10 (Total)")
    top10_bar_chart(top_total, "Top 10 Total — Chiffre d’affaires (Prime TTC)")

    st.divider()

    # ================= REPARTITION PAR PRODUIT =================
    st.subheader("📊 RÉPARTITION PAR PRODUIT")
    tab_prod = tableau_produit(df_f)
    st.dataframe(tab_prod, use_container_width=True, hide_index=True)

    if not tab_prod.empty and {"PRODUIT", "Prime TTC (FCFA)"}.issubset(tab_prod.columns):
        df_plot = tab_prod[tab_prod["PRODUIT"].astype(str).str.lower() != "total"].copy()
        df_plot["Prime TTC (FCFA)"] = pd.to_numeric(df_plot["Prime TTC (FCFA)"], errors="coerce").fillna(0)
        df_plot["Prime_M"] = df_plot["Prime TTC (FCFA)"] / 1_000_000

        fig_prod = px.bar(
            df_plot,
            x="PRODUIT",
            y="Prime_M",
            title="Répartition du CA par Produit",
            labels={"Prime_M": "CA (Millions FCFA)", "PRODUIT": "Produit"},
        )
        fig_prod.update_layout(height=420, yaxis=dict(ticksuffix=" M"))
        fig_prod.update_xaxes(tickangle=25)
        st.plotly_chart(fig_prod, use_container_width=True)

    st.divider()

    # ================= CA MENSUEL =================
    st.subheader("📅 CHIFFRE D’AFFAIRES MENSUEL")
    tab_month = monthly_ca_table(df_f, objectif_annuel)
    st.dataframe(tab_month, use_container_width=True, hide_index=True)

    st.divider()

    # ================= CA MOYEN JOURNALIER =================
    st.subheader("📆 CA MOYEN JOURNALIER PAR JOUR DE LA SEMAINE")
    tab_week = weekly_avg_table(df_f)
    st.dataframe(tab_week, use_container_width=True, hide_index=True)

    st.divider()

    # ================= ANALYSE COMPARATIVE =================
    st.subheader("🆚 ANALYSE COMPARATIVE")

    df_base_same_filters = _filter_like_current(df, segment_choice, inter_choice)

    prev_start, prev_end = _prev_same_duration(start_d, end_d)

    df_cur = _slice_dates(df_base_same_filters, start_d, end_d)
    df_prev = _slice_dates(df_base_same_filters, prev_start, prev_end)

    k_cur = compute_kpis(df_cur, objectif_annuel)
    k_prev = compute_kpis(df_prev, objectif_annuel)

    comp_week_df = _make_comp_df(
        k_cur, k_prev,
        label_cur=f"{start_d} → {end_d}",
        label_ref=f"{prev_start} → {prev_end}",
    )

    st.markdown(f"**1) Semaine N vs Semaine N-1**  \nN : {start_d} → {end_d} | N-1 : {prev_start} → {prev_end}")
    st.dataframe(comp_week_df, use_container_width=True, hide_index=True)

    mode_ref = st.selectbox(
        "📌 Choisir la référence",
        [
            "Période précédente (même durée)",
            "Même période mois précédent",
            "Même période année précédente",
        ],
        index=0,
        key="mode_ref_comp",
    )

    if mode_ref == "Période précédente (même durée)":
        ref_start, ref_end = prev_start, prev_end
    elif mode_ref == "Même période mois précédent":
        ref_start = _shift_month(start_d, -1)
        ref_end = _shift_month(end_d, -1)
    else:
        ref_start = _shift_year(start_d, -1)
        ref_end = _shift_year(end_d, -1)

    df_ref = _slice_dates(df_base_same_filters, ref_start, ref_end)
    k_ref = compute_kpis(df_ref, objectif_annuel)

    comp_ref_df = _make_comp_df(
        k_cur, k_ref,
        label_cur=f"{start_d} → {end_d}",
        label_ref=f"{ref_start} → {ref_end}",
    )

    st.markdown(f"**2) Période vs Référence** ({mode_ref})  \nRéf : {ref_start} → {ref_end}")
    st.dataframe(comp_ref_df, use_container_width=True, hide_index=True)

    # 3) MTD vs moyenne 3 derniers mois
    cur_m_start = _month_start(end_d)
    df_mtd = _slice_dates(df_base_same_filters, cur_m_start, end_d)
    k_mtd = compute_kpis(df_mtd, objectif_annuel)

    first_cur_month = pd.Timestamp(cur_m_start)
    m1_start = (first_cur_month - pd.DateOffset(months=1)).date().replace(day=1)
    m2_start = (first_cur_month - pd.DateOffset(months=2)).date().replace(day=1)
    m3_start = (first_cur_month - pd.DateOffset(months=3)).date().replace(day=1)

    def month_range(m_start: date) -> tuple[date, date]:
        ts = pd.Timestamp(m_start)
        end = (ts + pd.DateOffset(months=1) - pd.DateOffset(days=1)).date()
        return m_start, end

    m1_s, m1_e = month_range(m1_start)
    m2_s, m2_e = month_range(m2_start)
    m3_s, m3_e = month_range(m3_start)

    k_m1 = compute_kpis(_slice_dates(df_base_same_filters, m1_s, m1_e), objectif_annuel)
    k_m2 = compute_kpis(_slice_dates(df_base_same_filters, m2_s, m2_e), objectif_annuel)
    k_m3 = compute_kpis(_slice_dates(df_base_same_filters, m3_s, m3_e), objectif_annuel)

    keys = ["ca", "souscriptions", "renouvellements", "nouvelles_affaires", "montant_pool", "montant_hors_pool"]
    k_avg3 = {kk: (float(k_m1.get(kk, 0)) + float(k_m2.get(kk, 0)) + float(k_m3.get(kk, 0))) / 3 for kk in keys}

    comp_m3_df = _make_comp_df(
        k_mtd, k_avg3,
        label_cur=f"MTD {cur_m_start} → {end_d}",
        label_ref="Moyenne 3 derniers mois",
    )

    m3_title = "Mois en cours (MTD) vs Moyenne des 3 derniers mois"
    st.markdown(f"**3) {m3_title}**")
    st.dataframe(comp_m3_df, use_container_width=True, hide_index=True)

    st.divider()

    # ================= EXPORT PDF =================
    st.sidebar.header("⬇️ Rapport PDF")
    filters_key = f"{segment_label}|{inter_label}|{start_date_str}|{end_date_str}"

    if st.sidebar.button("🧾 Générer PDF", use_container_width=True, key="btn_pdf"):
        if st.session_state.get("pdf_key") == filters_key and st.session_state.get("pdf_bytes") is not None:
            st.sidebar.success("PDF déjà généré pour ces filtres ✅")
        else:
            with st.spinner("⏳ Génération du PDF en cours..."):
                st.session_state.pdf_bytes = export_pdf_dashboard(
                    k=k,
                    segment=segment_label,
                    intermediaire=inter_label,
                    start_date=start_date_str,
                    end_date=end_date_str,
                    objectif_annuel=objectif_annuel,
                    tab_produit=tab_prod,
                    tab_month=tab_month,
                    tab_week=tab_week,

                    comp_week=comp_week_df,
                    week_n_label=f"{start_d} → {end_d}",
                    week_n1_label=f"{prev_start} → {prev_end}",

                    comp_ref=comp_ref_df,
                    ref_mode_label=mode_ref,
                    ref_label=f"{ref_start} → {ref_end}",

                    comp_m3=comp_m3_df,
                    m3_title=m3_title,
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
