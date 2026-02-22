import streamlit as st
import pandas as pd
import numpy as np
from datetime import timedelta

OBJECTIF_ANNUEL = 4_826_339_000

# -------------------- Helpers --------------------
def normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")

    # Harmoniser OPERATEUR si le fichier utilise OPERATOR
    if "OPERATOR" in df.columns and "OPERATEUR" not in df.columns:
        df["OPERATEUR"] = df["OPERATOR"]

    # Conversions
    if "DATE_CREATE" in df.columns:
        df["DATE_CREATE"] = pd.to_datetime(df["DATE_CREATE"], errors="coerce")

    if "PRIME_TTC" in df.columns:
        df["PRIME_TTC"] = pd.to_numeric(df["PRIME_TTC"], errors="coerce").fillna(0)

    if "POOLS_TPV" in df.columns:
        df["POOLS_TPV"] = pd.to_numeric(df["POOLS_TPV"], errors="coerce").fillna(0).astype(int)

    if "EST_RENOUVELLER" in df.columns:
        df["EST_RENOUVELLER"] = (
            df["EST_RENOUVELLER"].astype(str).str.upper()
            .isin(["TRUE", "1", "OUI", "YES"])
        )

    if "NOM_AGENT" in df.columns:
        df["NOM_AGENT"] = df["NOM_AGENT"].astype(str).str.strip().replace("", pd.NA)

    if "OPERATEUR" in df.columns:
        df["OPERATEUR"] = df["OPERATEUR"].astype(str).str.strip()

        # Segment: Courtier si BROKER, sinon Agent
        df["SEGMENT"] = np.where(
            df["OPERATEUR"].str.upper().str.contains("BROKER", na=False),
            "Courtier",
            "Agent"
        )
    else:
        df["SEGMENT"] = "Agent"  # fallback si OPERATEUR absent

    return df


def get_week_range(anchor_date):
    """Semaine ISO Lun->Dim pour une date donnée."""
    d = pd.to_datetime(anchor_date)
    start = d - timedelta(days=d.weekday())     # Monday
    end = start + timedelta(days=6)            # Sunday
    return start.normalize(), end.normalize()


def get_month_range(anchor_date):
    """Mois complet pour une date donnée."""
    d = pd.to_datetime(anchor_date)
    start = d.replace(day=1).normalize()
    end = (start + pd.offsets.MonthEnd(1)).normalize()
    return start, end


def apply_sidebar_filters(df: pd.DataFrame, analysis_type: str):
    # Vérif date
    if "DATE_CREATE" not in df.columns or df["DATE_CREATE"].isna().all():
        st.error("❌ Colonne DATE_CREATE manquante ou vide.")
        st.stop()

    st.sidebar.header("🔎 Filtres")

    # 1) Date unique (ancre)
    dmin, dmax = df["DATE_CREATE"].min(), df["DATE_CREATE"].max()
    anchor = st.sidebar.date_input("📅 Choisir une date", value=dmax.date())

    # 2) Segment Agent/Courtier via OPERATEUR
    choix_segment = st.sidebar.selectbox("👤 Type", ["Tous", "Agent", "Courtier"])

    df_f = df.copy()
    if choix_segment != "Tous":
        df_f = df_f[df_f["SEGMENT"] == choix_segment]

    # 3) Fenêtre temporelle selon le type d’analyse
    if analysis_type == "Rapport Hebdomadaire":
        start, end = get_week_range(anchor)
        label = "Semaine (Lun–Dim)"
    else:
        start, end = get_month_range(anchor)
        label = "Mois"

    df_f = df_f[
        (df_f["DATE_CREATE"] >= start) &
        (df_f["DATE_CREATE"] <= end)
    ].copy()

    st.sidebar.caption(f"🗓️ {label} : {start.date()} → {end.date()}")

    return df_f, start, end, choix_segment


def compute_kpis(df_f: pd.DataFrame):
    ca_total = df_f["PRIME_TTC"].sum() if "PRIME_TTC" in df_f.columns else 0
    souscriptions = len(df_f)

    renouvellements = int(df_f["EST_RENOUVELLER"].sum()) if "EST_RENOUVELLER" in df_f.columns else 0
    nouvelles_affaires = souscriptions - renouvellements

    montant_pool = df_f.loc[df_f.get("POOLS_TPV", 0) == 1, "PRIME_TTC"].sum() if "PRIME_TTC" in df_f.columns else 0
    montant_hors_pool = df_f.loc[df_f.get("POOLS_TPV", 0) == 0, "PRIME_TTC"].sum() if "PRIME_TTC" in df_f.columns else 0

    nb_agents = (
        df_f["NOM_AGENT"].dropna().nunique()
        if "NOM_AGENT" in df_f.columns else 0
    )

    ratio_objectif = ca_total / OBJECTIF_ANNUEL if OBJECTIF_ANNUEL else 0
    taux_renouv = renouvellements / souscriptions if souscriptions else 0

    return ca_total, souscriptions, renouvellements, nouvelles_affaires, montant_pool, montant_hors_pool, nb_agents, ratio_objectif, taux_renouv


# -------------------- Pages --------------------
def page_rapport_hebdo(df: pd.DataFrame):
    st.subheader("📅 Rapport Hebdomadaire")
    df_f, start, end, seg = apply_sidebar_filters(df, "Rapport Hebdomadaire")

    (ca_total, sousc, renouv, newbiz, pool, hors_pool, nb_agents, ratio_obj, taux_renouv) = compute_kpis(df_f)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💰 CA Total", f"{ca_total:,.0f}")
    c2.metric("📄 Souscriptions", f"{sousc:,}")
    c3.metric("🔁 Renouvellements", f"{renouv:,}")
    c4.metric("🆕 Nouvelles affaires", f"{newbiz:,}")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("🏦 Montant POOL", f"{pool:,.0f}")
    c6.metric("🚫 Montant Hors POOL", f"{hors_pool:,.0f}")
    c7.metric("👥 Agents distinct", f"{nb_agents:,}")
    c8.metric("🎯 Taux objectif", f"{ratio_obj:.2%}")

    st.caption(f"Périmètre: {seg} | {start.date()} → {end.date()}")

    # Tableau des détails (filtré)
    st.write("### 📋 Détails (filtrés)")
    st.dataframe(df_f, use_container_width=True)

    # Mini agrégat jour dans la semaine (utile)
    if "DATE_CREATE" in df_f.columns:
        df_f["JOUR"] = df_f["DATE_CREATE"].dt.date
        daily = (
            df_f.groupby("JOUR")
               .agg(CA=("PRIME_TTC", "sum"), Souscriptions=("JOUR", "count"))
               .reset_index()
               .sort_values("JOUR")
        )
        st.write("### 📈 CA journalier (dans la semaine)")
        st.line_chart(daily.set_index("JOUR")["CA"])


def page_taux_renouvellement(df: pd.DataFrame):
    st.subheader("🔁 Taux de renouvellement (mensuel)")
    df_f, start, end, seg = apply_sidebar_filters(df, "Taux de renouvellement")

    (ca_total, sousc, renouv, newbiz, pool, hors_pool, nb_agents, ratio_obj, taux_renouv) = compute_kpis(df_f)

    c1, c2, c3 = st.columns(3)
    c1.metric("📄 Souscriptions", f"{sousc:,}")
    c2.metric("🔁 Renouvellements", f"{renouv:,}")
    c3.metric("📈 Taux de renouvellement", f"{taux_renouv:.2%}")

    st.caption(f"Périmètre: {seg} | {start.date()} → {end.date()}")

    # Suivi par mois (sur tout df filtré segment uniquement, pas que le mois)
    # => Ici on repart du df segmenté (sans filtre mois) pour voir tendance
    df_seg = df.copy()
    if seg != "Tous":
        df_seg = df_seg[df_seg["SEGMENT"] == seg].copy()

    if "DATE_CREATE" in df_seg.columns:
        df_seg["MOIS"] = df_seg["DATE_CREATE"].dt.to_period("M").astype(str)
        m = (
            df_seg.groupby("MOIS")
                  .agg(
                      Souscriptions=("MOIS", "count"),
                      Renouvellements=("EST_RENOUVELLER", "sum") if "EST_RENOUVELLER" in df_seg.columns else ("MOIS", "count")
                  )
                  .reset_index()
                  .sort_values("MOIS")
        )
        m["Taux_Renouvellement"] = np.where(m["Souscriptions"] > 0, m["Renouvellements"] / m["Souscriptions"], 0)

        st.write("### 📅 Taux de renouvellement par mois (tendance)")
        st.dataframe(m, use_container_width=True)
        st.line_chart(m.set_index("MOIS")["Taux_Renouvellement"])


# -------------------- App --------------------
st.set_page_config(page_title="Suivi Production Auto", layout="wide")
st.title("📊 Suivi Production – Assurance Auto")

uploaded = st.file_uploader("📂 Charger un fichier (Excel/CSV)", type=["xlsx", "csv"])
if not uploaded:
    st.info("Charge un fichier pour démarrer.")
    st.stop()

df = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
df = normalize(df)

st.sidebar.header("🧭 Type d’analyse")
analyse = st.sidebar.selectbox("Choisir", ["Rapport Hebdomadaire", "Taux de renouvellement"])

if analyse == "Rapport Hebdomadaire":
    page_rapport_hebdo(df)
else:
    page_taux_renouvellement(df)
