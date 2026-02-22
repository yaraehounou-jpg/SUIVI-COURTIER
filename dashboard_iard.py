from __future__ import annotations

from datetime import datetime
from io import BytesIO
import unicodedata

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px


st.set_page_config(page_title="TABLEAU DE BORD DG - IARD", layout="wide")


# -----------------------
# Helpers
# -----------------------
def _norm(s: str) -> str:
    t = str(s).strip().replace("\u00A0", " ").upper().replace(" ", "_")
    t = "".join(ch for ch in unicodedata.normalize("NFD", t) if unicodedata.category(ch) != "Mn")
    t = "".join(ch for ch in t if ch.isalnum() or ch == "_")
    return t


def _find_col(df: pd.DataFrame, targets: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    lookup = {_norm(c): c for c in df.columns}
    for t in targets:
        k = _norm(t)
        if k in lookup:
            return lookup[k]
    return None


def _to_datetime(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors="coerce")
    dt2 = pd.to_datetime(series, errors="coerce", dayfirst=True)
    dt = dt.fillna(dt2)
    num = pd.to_numeric(series, errors="coerce")
    mask = dt.isna() & num.notna()
    if mask.any():
        dt.loc[mask] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask], unit="D")
    return dt


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
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _safe_read(uploaded) -> pd.DataFrame:
    if uploaded is None:
        return pd.DataFrame()
    if uploaded.name.lower().endswith(".csv"):
        try:
            return pd.read_csv(uploaded, sep=None, engine="python")
        except Exception:
            uploaded.seek(0)
            return pd.read_csv(uploaded, sep=";", encoding="utf-8")
    return pd.read_excel(uploaded)


def _fmt_int(x: float) -> str:
    return f"{int(round(float(x))):,}".replace(",", " ")


def _fmt_money(x: float) -> str:
    return f"{int(round(float(x))):,} FCFA".replace(",", " ")


def _kpi_card(title: str, value: str, sub: str = "") -> str:
    return f"""
    <div class='iard-kpi'>
      <div class='iard-kpi-title'>{title}</div>
      <div class='iard-kpi-value'>{value}</div>
      <div class='iard-kpi-sub'>{sub}</div>
    </div>
    """


def _kpi_effectif_card(value: str) -> str:
    return f"""
    <div class='iard-effectif'>
      <div class='iard-effectif-icon'>🔎👥</div>
      <div class='iard-effectif-box'>
        <div class='iard-effectif-title'>Effectif Total</div>
        <div class='iard-effectif-value'>{value}</div>
      </div>
    </div>
    """


MONTH_FR = {
    1: "JANVIER", 2: "FEVRIER", 3: "MARS", 4: "AVRIL", 5: "MAI", 6: "JUIN",
    7: "JUILLET", 8: "AOUT", 9: "SEPTEMBRE", 10: "OCTOBRE", 11: "NOVEMBRE", 12: "DECEMBRE"
}


# -----------------------
# UI + style
# -----------------------
st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;700;800&display=swap');
      html, body, [class*="css"] { font-family: 'Montserrat', sans-serif; }
      .block-container { padding-top: 0.6rem; padding-bottom: 1rem; max-width: 100%; }
      .iard-header {
        display:flex; align-items:center; justify-content:space-between;
        background: linear-gradient(90deg,#1f2630 0%,#2b2f35 45%,#4b3f2a 100%);
        color:#fff; border-radius:0; padding:10px 16px; margin:-10px -10px 10px -10px;
        border-bottom:2px solid #f2b705;
      }
      .iard-date { color:#ffd400; font-size:20px; font-weight:800; min-width:170px; }
      .iard-title { font-size:44px; font-weight:800; letter-spacing:1px; flex:1; text-align:center; }
      .iard-tag { font-size:18px; font-weight:700; opacity:.95; min-width:100px; text-align:right; }
      .iard-panel { background:#ececec; border:1px solid #d4d4d4; border-radius:10px; padding:10px; }
      .iard-section-title { color:#d56b1f; font-weight:800; font-size:28px; text-align:center; margin:0; }
      .iard-kpi {
        background:#efefef; border:3px solid #f2b705; border-radius:6px;
        padding:8px 10px; min-height:92px;
      }
      .iard-kpi-title { font-size:22px; font-weight:700; color:#404040; }
      .iard-kpi-value { font-size:42px; font-weight:800; color:#1f1f1f; line-height:1.1; }
      .iard-kpi-sub { font-size:12px; color:#626262; }
      .iard-subtitle { font-size:22px; font-style:italic; font-weight:700; margin:4px 0 8px 0; }
      .iard-rh-title { color:#d56b1f; text-align:center; font-size:44px; font-weight:800; }
      .iard-fin-title { color:#d56b1f; text-align:center; font-size:44px; font-weight:800; }
      .iard-effectif{
        display:flex; align-items:flex-end; gap:10px;
        background: linear-gradient(180deg, #f4f5f7 0%, #eceef2 100%);
        border: 1px solid #d7dce3; border-radius:8px;
        padding:8px 10px; margin-bottom:8px;
      }
      .iard-effectif-icon{
        width:68px; height:68px; border-radius:50%;
        border:3px solid #2f4e78; color:#2f4e78;
        display:flex; align-items:center; justify-content:center;
        font-size:26px; line-height:1; background:#f7f9fc;
      }
      .iard-effectif-box{
        flex:1;
        background:#ffffff; border:1px solid #d6dbe4;
        box-shadow:0 2px 6px rgba(0,0,0,0.10);
        padding:8px 14px; min-height:72px;
      }
      .iard-effectif-title{ font-size:46px; font-weight:700; color:#2a2a2a; line-height:1; }
      .iard-effectif-value{ font-size:58px; font-weight:800; color:#111; line-height:1.05; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.sidebar.header("Fichier IARD")
uploaded = st.sidebar.file_uploader("Charger la base (Excel/CSV)", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.markdown(
        f"""
        <div class='iard-header'>
          <div class='iard-date'>{datetime.now().strftime('%d/%m/%Y')}</div>
          <div class='iard-title'>TABLEAU DE BORD DG</div>
          <div class='iard-tag'>IARD</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("Charge ton fichier IARD pour alimenter le dashboard.")
    st.stop()

try:
    df = _safe_read(uploaded)
except Exception as exc:
    st.error(f"Erreur lecture fichier: {exc}")
    st.stop()

if df.empty:
    st.warning("Le fichier est vide.")
    st.stop()

# -----------------------
# Mapping colonnes
# -----------------------
def_col_date = _find_col(df, ["DATE DEMISSION", "DATE D'EMISSION", "DATE_EMISSION", "DATE", "DATE_CREATE"])
def_col_value = _find_col(df, ["TOTAL PRIME TTC", "TOTAL_PRIME_TTC", "PRIME TTC CALCULEE", "MONTANT", "AMOUNT"])
def_col_gare = _find_col(df, ["INTERMEDIAIRE", "INTERMEDIAIRE_", "GROUPE INT", "GROUPE_INT"])
def_col_site = _find_col(df, ["BRANCHE", "LINE OF BUSINESS", "LINE_OF_BUSINESS"])
def_col_mode = _find_col(df, ["CHANNELS", "MODE DE PAIEMENT", "MODE"])
def_col_type = _find_col(df, ["BRANCHE", "LINE OF BUSINESS", "LINE_OF_BUSINESS"])
def_col_lob = _find_col(df, ["LINE OF BUSINESS", "LINE_OF_BUSINESS", "BRANCHE"])
def_col_class = _find_col(df, ["CATEGORIE", "CATEGORIE VEHICULE", "TYPE VEHICULE", "TYPE_VEHICULE"])
def_col_inf = _find_col(df, ["STATUT", "STATUS"])
def_col_rh = _find_col(df, ["CHANNELS", "GROUPE INT", "GROUPE_INT", "INTERMEDIAIRE"])
def_col_policy = _find_col(df, ["N* POLICE", "N° POLICE", "N POLICE", "POLICE", "POLICY NO", "POLICY_NO"])

with st.sidebar.expander("Mappage colonnes", expanded=True):
    col_date = st.selectbox("Date", [None] + list(df.columns), index=(list(df.columns).index(def_col_date)+1) if def_col_date in df.columns else 0)
    col_value = st.selectbox("Chiffre d'affaires", [None] + list(df.columns), index=(list(df.columns).index(def_col_value)+1) if def_col_value in df.columns else 0)
    col_gare = st.selectbox("Gare", [None] + list(df.columns), index=(list(df.columns).index(def_col_gare)+1) if def_col_gare in df.columns else 0)
    col_site = st.selectbox("Site", [None] + list(df.columns), index=(list(df.columns).index(def_col_site)+1) if def_col_site in df.columns else 0)
    col_lob = st.selectbox("Line of Business", [None] + list(df.columns), index=(list(df.columns).index(def_col_lob)+1) if def_col_lob in df.columns else 0)
    col_mode = st.selectbox("Mode de paiement", [None] + list(df.columns), index=(list(df.columns).index(def_col_mode)+1) if def_col_mode in df.columns else 0)
    col_type = st.selectbox("Type (PEAGE/PESAGE)", [None] + list(df.columns), index=(list(df.columns).index(def_col_type)+1) if def_col_type in df.columns else 0)
    col_class = st.selectbox("Classe véhicule", [None] + list(df.columns), index=(list(df.columns).index(def_col_class)+1) if def_col_class in df.columns else 0)
    col_inf = st.selectbox("Type infraction", [None] + list(df.columns), index=(list(df.columns).index(def_col_inf)+1) if def_col_inf in df.columns else 0)
    col_rh = st.selectbox("Catégorie RH", [None] + list(df.columns), index=(list(df.columns).index(def_col_rh)+1) if def_col_rh in df.columns else 0)
    col_policy = st.selectbox("N° Police", [None] + list(df.columns), index=(list(df.columns).index(def_col_policy)+1) if def_col_policy in df.columns else 0)

# Finance (saisie utilisateur)
comm_rev_col = _find_col(df, ["COMMISSION A REVERSER", "COMMISSION_A_REVERSER"])
tax_ded_col = _find_col(df, ["TAXES A DEDUIRE", "TAXES_A_DEDUIRE"])
with st.sidebar.expander("Finance (saisie / auto)", expanded=True):
    budget_annuel = st.number_input("Budget annuel (FCFA)", min_value=0.0, value=10_000_000_000.0, step=100_000_000.0)
    auto_fin = st.checkbox("Calcul auto des taux finance", value=True)
    tx_mobil = st.slider("Taux de Mobilisation Budgétaire (%)", 0.0, 100.0, 7.87, 0.01, disabled=auto_fin)
    tx_regle = st.slider("Taux de Règlement Des Décomptes (%)", 0.0, 100.0, 65.25, 0.01, disabled=auto_fin)
    tx_att = st.slider("Taux de Décomptes en Attente (%)", 0.0, 100.0, 34.75, 0.01, disabled=auto_fin)

lob_objectif_df = pd.DataFrame(columns=["Line of Business", "Objectif annuel (FCFA)"])
if def_col_lob and def_col_lob in df.columns:
    lobs = (
        df[def_col_lob]
        .astype("string")
        .fillna("")
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )
    if lobs:
        default_obj = budget_annuel / max(1, len(lobs))
        lob_objectif_df = pd.DataFrame(
            {"Line of Business": lobs, "Objectif annuel (FCFA)": [default_obj] * len(lobs)}
        )
        with st.sidebar.expander("Objectif annuel par branche", expanded=False):
            lob_objectif_df = st.data_editor(
                lob_objectif_df,
                num_rows="fixed",
                use_container_width=True,
                key="iard_objectif_lob",
            )

# -----------------------
# Préparation temporelle
# -----------------------
dfx = df.copy()
if col_date is not None:
    dfx["_date"] = _to_datetime(dfx[col_date])
    dfx = dfx.dropna(subset=["_date"]).copy()
else:
    dfx["_date"] = pd.NaT

if col_value is not None:
    dfx["_val"] = _to_num(dfx[col_value])
else:
    dfx["_val"] = 1.0

if dfx.empty:
    st.warning("Aucune date valide. Vérifie la colonne Date dans le mappage.")
    st.stop()

dfx["_year"] = dfx["_date"].dt.year
dfx["_month"] = dfx["_date"].dt.month
dfx["_week"] = dfx["_date"].dt.isocalendar().week.astype(int)
dfx["_day"] = dfx["_date"].dt.date

years = sorted(dfx["_year"].dropna().unique().tolist())
year_sel = st.sidebar.selectbox("Année", years, index=len(years)-1 if years else 0)

months_av = sorted(dfx.loc[dfx["_year"] == year_sel, "_month"].dropna().unique().tolist())
month_labels = ["Tout"] + [MONTH_FR.get(m, str(m)) for m in months_av]
month_sel_label = st.sidebar.selectbox("Mois", month_labels, index=1 if len(month_labels) > 1 else 0)
month_sel = None if month_sel_label == "Tout" else {v: k for k, v in MONTH_FR.items()}.get(month_sel_label)

wk_df = dfx[dfx["_year"] == year_sel]
if month_sel is not None:
    wk_df = wk_df[wk_df["_month"] == month_sel]
weeks_av = sorted(wk_df["_week"].dropna().unique().tolist())
week_sel_label = st.sidebar.selectbox("Semaine", ["Tout"] + [str(w) for w in weeks_av], index=0)
week_sel = None if week_sel_label == "Tout" else int(week_sel_label)

days_df = wk_df.copy()
if week_sel is not None:
    days_df = days_df[days_df["_week"] == week_sel]
days_av = sorted(days_df["_day"].dropna().unique().tolist())
day_sel_label = st.sidebar.selectbox("Jour", ["Tout"] + [d.strftime("%d/%m/%Y") for d in days_av], index=0)
day_sel = None if day_sel_label == "Tout" else pd.to_datetime(day_sel_label, dayfirst=True).date()

f = dfx[dfx["_year"] == year_sel].copy()
if month_sel is not None:
    f = f[f["_month"] == month_sel]
if week_sel is not None:
    f = f[f["_week"] == week_sel]
if day_sel is not None:
    f = f[f["_day"] == day_sel]

if col_gare and col_gare in f.columns:
    gares = sorted(f[col_gare].astype("string").fillna("-").unique().tolist())
    gare_sel = st.sidebar.selectbox("Gares", ["Tout"] + gares, index=0)
    if gare_sel != "Tout":
        f = f[f[col_gare].astype("string") == gare_sel]

if col_site and col_site in f.columns:
    sites = sorted(f[col_site].astype("string").fillna("-").unique().tolist())
    site_sel = st.sidebar.selectbox("Sites", ["Tout"] + sites, index=0)
    if site_sel != "Tout":
        f = f[f[col_site].astype("string") == site_sel]

# Filtre de périmètre cohérent sur base complète (pour comparatif mois-1)
scope_mask = pd.Series([True] * len(dfx), index=dfx.index)
scope_mask &= dfx["_year"].eq(year_sel)
if col_gare and col_gare in dfx.columns and "gare_sel" in locals() and gare_sel != "Tout":
    scope_mask &= dfx[col_gare].astype("string").eq(gare_sel)
if col_site and col_site in dfx.columns and "site_sel" in locals() and site_sel != "Tout":
    scope_mask &= dfx[col_site].astype("string").eq(site_sel)

scope_df = dfx.loc[scope_mask].copy()

# -----------------------
# Calculs
# -----------------------
if f.empty:
    st.warning("Aucune donnée après filtres.")
    st.stop()

n_days = max(1, f["_day"].nunique())
tmj = f["_val"].sum() / n_days

peage_mask = pd.Series([True] * len(f), index=f.index)
pesage_mask = pd.Series([True] * len(f), index=f.index)
if col_type and col_type in f.columns:
    t = f[col_type].astype("string").str.upper().fillna("")
    peage_mask = t.str.contains("AUTO|MOTOR|AUTOMOBILE|1-MOTOR", regex=True)
    pesage_mask = ~peage_mask

tmj_peage = f.loc[peage_mask, "_val"].sum() / max(1, f.loc[peage_mask, "_day"].nunique()) if peage_mask.any() else 0
n_pesage = f.loc[pesage_mask, "_val"].sum() / max(1, f.loc[pesage_mask, "_day"].nunique()) if pesage_mask.any() else 0

if auto_fin:
    ca_total = float(f["_val"].sum())
    tx_mobil = min(100.0, (ca_total / budget_annuel * 100.0) if budget_annuel > 0 else 0.0)
    comm_due = float(_to_num(f[comm_rev_col]).sum()) if comm_rev_col and comm_rev_col in f.columns else 0.0
    tax_ded = float(_to_num(f[tax_ded_col]).sum()) if tax_ded_col and tax_ded_col in f.columns else 0.0
    tx_regle = min(100.0, max(0.0, ((comm_due - tax_ded) / comm_due * 100.0) if comm_due > 0 else 0.0))
    tx_att = max(0.0, 100.0 - tx_regle)

# RH
if col_rh and col_rh in f.columns:
    rh = f.groupby(f[col_rh].astype("string").fillna("Non défini"), dropna=False).size().reset_index(name="Effectif")
    rh.columns = ["Catégorie", "Effectif"]
else:
    rh = pd.DataFrame({"Catégorie": ["Péage", "Pesage", "Siège", "Aire station"], "Effectif": [1370, 336, 208, 13]})

effectif_total = int(rh["Effectif"].sum())

# -----------------------
# Header + filtre visible
# -----------------------
st.markdown(
    f"""
    <div class='iard-header'>
      <div class='iard-date'>{datetime.now().strftime('%d/%m/%Y')}</div>
      <div class='iard-title'>TABLEAU DE BORD DG</div>
      <div class='iard-tag'>IARD</div>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3, c4 = st.columns(4)
jour_opts = ["Tout"] + [d.strftime("%d/%m/%Y") for d in days_av]
sem_opts = ["Tout"] + [str(w) for w in weeks_av]

jour_idx = 0
if day_sel is not None:
    day_lbl = day_sel.strftime("%d/%m/%Y")
    if day_lbl in jour_opts:
        jour_idx = jour_opts.index(day_lbl)

sem_idx = 0
if week_sel is not None:
    sem_lbl = str(week_sel)
    if sem_lbl in sem_opts:
        sem_idx = sem_opts.index(sem_lbl)

c1.selectbox("Jour", jour_opts, index=jour_idx, key="ui_jour", disabled=True)
c2.selectbox("Semaine", sem_opts, index=sem_idx, key="ui_sem", disabled=True)
c3.selectbox("Mois", month_labels, index=month_labels.index(month_sel_label), key="ui_mois", disabled=True)
c4.selectbox("Année", [str(y) for y in years], index=[str(y) for y in years].index(str(year_sel)), key="ui_ann", disabled=True)

# -----------------------
# Layout principal
# -----------------------
left, mid, right = st.columns([1.55, 1.05, 0.8], gap="small")

with left:
    st.markdown("<p class='iard-section-title'>PEAGE</p>", unsafe_allow_html=True)
    st.markdown(_kpi_card("Production Moy. Journalière", _fmt_money(tmj_peage), f"CA Segment: {_fmt_money(f.loc[peage_mask, '_val'].sum())}"), unsafe_allow_html=True)

    st.markdown("<p class='iard-subtitle'>CA par branche vs objectif annuel</p>", unsafe_allow_html=True)
    # Utilisation stricte de la colonne Line of Business
    lob_key = def_col_lob if (def_col_lob and def_col_lob in scope_df.columns) else None
    if lob_key is not None:
        if month_sel is not None:
            anchor_y, anchor_m = year_sel, month_sel
        else:
            anchor_dt = scope_df["_date"].max()
            anchor_y, anchor_m = int(anchor_dt.year), int(anchor_dt.month)
        prev_y, prev_m = (anchor_y - 1, 12) if anchor_m == 1 else (anchor_y, anchor_m - 1)

        cur_m = scope_df[(scope_df["_year"] == anchor_y) & (scope_df["_month"] == anchor_m)].copy()
        prv_m = scope_df[(scope_df["_year"] == prev_y) & (scope_df["_month"] == prev_m)].copy()

        ca_cur = cur_m.groupby(lob_key, dropna=False)["_val"].sum().rename("ca_cur")
        ca_prv = prv_m.groupby(lob_key, dropna=False)["_val"].sum().rename("ca_prv")
        ytd_m = scope_df[scope_df["_date"] <= cur_m["_date"].max()] if not cur_m.empty else scope_df
        ca_ytd = ytd_m.groupby(lob_key, dropna=False)["_val"].sum().rename("ca_ytd")
        if col_policy and col_policy in cur_m.columns:
            cnt_cur = cur_m.groupby(lob_key, dropna=False)[col_policy].count().rename("nb_contrats")
        else:
            cnt_cur = cur_m.groupby(lob_key, dropna=False).size().rename("nb_contrats")

        g = pd.concat([ca_cur, ca_prv, ca_ytd, cnt_cur], axis=1).fillna(0).reset_index()
        if not lob_objectif_df.empty:
            obj_map = dict(zip(lob_objectif_df["Line of Business"].astype("string"), _to_num(lob_objectif_df["Objectif annuel (FCFA)"])))
            g["objectif"] = g[lob_key].astype("string").map(obj_map).fillna(0.0)
        else:
            g["objectif"] = 0.0
        g["evol_pct"] = np.where(g["ca_prv"] > 0, (g["ca_cur"] - g["ca_prv"]) / g["ca_prv"] * 100.0, 0.0)
        g["atteinte_obj_pct"] = np.where(g["objectif"] > 0, (g["ca_ytd"] / g["objectif"]) * 100.0, 0.0)
        g = g.sort_values("ca_ytd", ascending=False).reset_index(drop=True)
        g["lbl"] = g.apply(
            lambda r: f"{_fmt_money(r['ca_ytd'])} | {_fmt_int(r['nb_contrats'])} contrats",
            axis=1,
        )
        g["perf"] = np.where(g["atteinte_obj_pct"] >= 100, "Atteint", "En cours")
        fig_gare = px.bar(
            g,
            x="ca_ytd",
            y=lob_key,
            orientation="h",
            text="lbl",
            color_discrete_sequence=["#f2b705"],
        )
        fig_gare.update_traces(marker_line_color="#e5b100", marker_line_width=1.2)
        for _, r in g.iterrows():
            up = float(r["evol_pct"]) >= 0
            evo_txt = f"{'▲' if up else '▼'} {r['evol_pct']:+.1f}% | Obj: {r['atteinte_obj_pct']:.1f}%"
            x_pos = float(r["ca_ytd"]) * 1.005 if float(r["ca_ytd"]) > 0 else 0
            fig_gare.add_annotation(
                x=x_pos,
                y=r[lob_key],
                text=evo_txt,
                showarrow=False,
                xanchor="left",
                yanchor="middle",
                font=dict(size=13, color=("#0b8f31" if up else "#d7191c")),
                bgcolor="rgba(245,245,245,0.9)",
                bordercolor="rgba(200,200,200,0.9)",
                borderwidth=1,
            )
    else:
        st.warning("Colonne 'Line of Business' introuvable dans la base. Merci de vérifier l'en-tête.")
        g = pd.DataFrame({"Line of Business": [], "ca_ytd": []})
        fig_gare = px.bar(g, x="ca_ytd", y="Line of Business", orientation="h")
    fig_gare.update_traces(textposition="outside", cliponaxis=False)
    fig_gare.update_layout(
        height=max(220, 34 * len(g) + 120 if "g" in locals() else 380),
        margin=dict(l=10, r=90, t=10, b=10),
        showlegend=False,
        xaxis_title="",
        yaxis_title="",
        bargap=0.25,
    )
    if lob_key is not None and "g" in locals():
        fig_gare.update_yaxes(type="category", categoryorder="array", categoryarray=g[lob_key].tolist()[::-1])
        fig_gare.update_xaxes(showgrid=True, gridcolor="rgba(180,180,180,0.35)")
    st.plotly_chart(fig_gare, use_container_width=True)

    st.markdown("<p class='iard-subtitle'>Mode de Paiement</p>", unsafe_allow_html=True)
    if col_mode and col_mode in f.columns:
        m = f.groupby(col_mode, dropna=False)["_val"].sum().reset_index()
        fig_mode = px.pie(m, values="_val", names=col_mode, hole=0.0)
    else:
        m = pd.DataFrame({"Mode": ["Espèces", "TAG", "Réquisition"], "Valeur": [75.81, 22.82, 1.37]})
        fig_mode = px.pie(m, values="Valeur", names="Mode", hole=0.0)
    fig_mode.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_mode, use_container_width=True)

    st.markdown("<p class='iard-subtitle'>Production moyenne journalière par catégorie</p>", unsafe_allow_html=True)
    if col_class and col_class in f.columns:
        cl = f.groupby(col_class, dropna=False)["_val"].sum().sort_values(ascending=False).reset_index()
        cl["Part %"] = (cl["_val"] / cl["_val"].sum() * 100).round(1)
        cl = cl.rename(columns={col_class: "CLASSE", "_val": "Trafic_Moyen_Journalier"})
        st.dataframe(cl, use_container_width=True, hide_index=True)
    else:
        st.info("Mappe une colonne 'Classe véhicule' pour ce tableau.")

with mid:
    st.markdown("<p class='iard-section-title'>PESAGE</p>", unsafe_allow_html=True)
    st.markdown(_kpi_card("Production Moy. Journalière", _fmt_money(n_pesage), f"CA Segment: {_fmt_money(f.loc[pesage_mask, '_val'].sum())}"), unsafe_allow_html=True)

    st.markdown("<p class='iard-subtitle'>Production moyenne journalière par Branche</p>", unsafe_allow_html=True)
    if col_site and col_site in f.columns:
        s = f.groupby(col_site, dropna=False)["_val"].sum().sort_values(ascending=True).tail(10).reset_index()
        fig_site = px.bar(s, x="_val", y=col_site, orientation="h", text="_val")
    else:
        s = f.groupby("_day")["_val"].sum().reset_index().tail(10)
        fig_site = px.bar(s, x="_val", y="_day", orientation="h", text="_val")
    fig_site.update_traces(marker_color="#f2b705", texttemplate="%{text:,.0f}")
    fig_site.update_layout(height=440, margin=dict(l=10, r=10, t=10, b=10), showlegend=False, xaxis_title="", yaxis_title="")
    st.plotly_chart(fig_site, use_container_width=True)

    st.markdown("<p class='iard-subtitle'>Répartition par Statut</p>", unsafe_allow_html=True)
    if col_inf and col_inf in f.columns:
        i = f.groupby(col_inf, dropna=False)["_val"].sum().reset_index()
        fig_inf = px.pie(i, values="_val", names=col_inf, hole=0.55)
    else:
        i = pd.DataFrame({"Type": ["NORME", "SURES", "SURFO", "SUREX"], "Valeur": [55.4, 25.01, 19.55, 0.04]})
        fig_inf = px.pie(i, values="Valeur", names="Type", hole=0.55)
    fig_inf.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_inf, use_container_width=True)

with right:
    st.markdown("<p class='iard-rh-title'>RESSOURCES HUMAINES</p>", unsafe_allow_html=True)
    st.markdown(_kpi_effectif_card(_fmt_int(effectif_total)), unsafe_allow_html=True)
    fig_rh = px.pie(rh, values="Effectif", names="Catégorie")
    fig_rh.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_rh, use_container_width=True)

    st.markdown("<p class='iard-fin-title'>FINANCE</p>", unsafe_allow_html=True)

    def _gauge(val: float, title: str, color: str):
        fig = px.pie(values=[val, 100-val], names=[title, "Reste"], hole=0.75)
        fig.update_traces(marker=dict(colors=[color, "#d9d9d9"]))
        fig.update_layout(height=200, showlegend=False, margin=dict(l=10, r=10, t=10, b=10), annotations=[dict(text=f"{val:.2f}%", x=0.5, y=0.5, showarrow=False, font=dict(size=34, color="#252525"))])
        st.markdown(f"<div style='font-size:20px;font-weight:700;text-align:center;margin-bottom:4px'>{title}</div>", unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)

    _gauge(tx_mobil, "Taux De Mobilisation Budgétaire", "#e27d34")
    _gauge(tx_regle, "Taux de Règlement Des Décomptes", "#f2b705")
    _gauge(tx_att, "Taux de Décomptes en Attente", "#4f7fd8")
