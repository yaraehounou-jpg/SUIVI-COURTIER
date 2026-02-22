import streamlit as st
import pandas as pd
import numpy as np
import polars as pl
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import date
import hashlib
import re
import unicodedata

st.set_page_config(page_title="Contrôle Production - Lazy", layout="wide")

st.markdown(
    """
    <style>
    section[data-testid="stSidebar"]{
      background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
      color: #e2e8f0;
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span{
      color: #f8fafc !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="radio"] div{
      color: #f8fafc !important;
      font-weight: 600;
    }
    section[data-testid="stSidebar"] input, 
    section[data-testid="stSidebar"] textarea{
      color: #0f172a !important;
      background: #f8fafc !important;
    }
    .sidebar-card{
      background: rgba(255,255,255,0.06);
      border: 1px solid rgba(255,255,255,0.12);
      border-radius: 14px;
      padding: 12px 14px;
      margin-bottom: 10px;
    }
    .kpi-grid{
      display: grid;
      grid-template-columns: repeat(6, minmax(140px, 1fr));
      gap: 12px;
      margin-bottom: 14px;
    }
    .kpi-card{
      background: linear-gradient(135deg, #0f172a, #1e293b);
      border: 1px solid rgba(255,255,255,0.12);
      border-radius: 16px;
      padding: 12px 14px;
      box-shadow: 0 6px 18px rgba(0,0,0,0.18);
    }
    .kpi-label{
      color: #cbd5e1;
      font-size: 12px;
      letter-spacing: 0.3px;
      text-transform: uppercase;
      margin-bottom: 6px;
    }
    .kpi-value{
      color: #f8fafc;
      font-size: 22px;
      font-weight: 800;
    }
    .kpi-grid-3{
      display: grid;
      grid-template-columns: repeat(3, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }
    .kpi-feature{
      background:#ffffff;
      border:3px solid #d46a6a;
      border-radius:22px;
      padding:12px 16px;
      min-height:120px;
      display:flex;
      flex-direction:column;
      justify-content:space-between;
    }
    .kpi-feature-title{
      color:#1f2937;
      font-size:14px;
      font-weight:700;
      margin-bottom:4px;
    }
    .kpi-feature-value{
      color:#d46a6a;
      font-size:34px;
      font-weight:900;
      line-height:1.05;
    }
    .kpi-feature-sub{
      color:#1f2937;
      font-size:12px;
      font-weight:700;
      margin-top:4px;
    }
    .kpi-feature-evol{
      font-size:22px;
      font-weight:900;
      margin-left:8px;
      margin-right:8px;
    }
    .kpi-feature-delta{
      color:#1f2937;
      font-size:18px;
      font-weight:800;
    }
    .kpi-feature-note{
      color:#374151;
      font-size:12px;
      margin-top:4px;
    }
    @media (max-width: 1100px){
      .kpi-grid{ grid-template-columns: repeat(3, 1fr); }
      .kpi-grid-3{ grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 700px){
      .kpi-grid{ grid-template-columns: repeat(2, 1fr); }
      .kpi-grid-3{ grid-template-columns: 1fr; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("DOSE EYES")

# =========================
# 1) Mapping Branche (selon ton tableau)
# =========================
BRANCHE_MAPPING = {
    "AUTOMOBILE": "AUTOMOBILE",
    "AUTRES DOMMAGES AUX BIENS": "INCENDIE & MULTIRISQUES",
    "CAUTION": "CAUTION",
    "DOMMAGES CORPORELS": "DOMMAGES CORPORELS",
    "INCENDIE & MULTIRISQUES": "INCENDIE & MULTIRISQUES",
    "RESPONSABILITE CIVILE": "RESPONSABILITE CIVILE",
    "SANTE": "SANTE",
    "TRANSPORT": "TRANSPORT",
    "GLOBALE DE BANQUE": "INCENDIE & MULTIRISQUES",
    "MULTIRISQUE PROFESSIONNELLE": "INCENDIE & MULTIRISQUES",
    "CAUTION D'AGREMENT": "CAUTION",
    "RC ENTREPRISES INDUST & COMMERCIALES": "RESPONSABILITE CIVILE",
    "FACULTES MARITIMES": "TRANSPORT",
    "RC STAGE": "RESPONSABILITE CIVILE",
    "CAUTION DE MARCHE": "CAUTION",
    "TOUS RISQUES CHANTIERS": "AUTRES DOMMAGES AUX BIENS",
    "PROMENADE & AFFAIRES": "AUTOMOBILE",
    "MULTIRISQUE HABITATION": "INCENDIE & MULTIRISQUES",
    "CAUTION D'AVANCE DE DEMARRAGE": "CAUTION",
    "CAUTION DE SOUMISSION": "CAUTION",
    "INDIVIDUELLE ACCIDENT GROUPE": "DOMMAGES CORPORELS",
    "ASSURANCE VOYAGE": "SANTE",
    "FLOTTE AUTOMOBILE": "AUTOMOBILE",
    "RC SCOLAIRE": "RESPONSABILITE CIVILE",
    "INCENDIE & MULTIRISQUE": "INCENDIE & MULTIRISQUES",
    "TOUS DOMMAGES SAUF": "AUTRES DOMMAGES AUX BIENS",
    "VIOLENCE POLITIQUE ET TERRORISME": "AUTRES DOMMAGES AUX BIENS",
    "MONO AUTOMOBILE": "AUTOMOBILE",
    "INDIVIDUEL ACCIDENTS GROUPE": "DOMMAGES CORPORELS",
    "INCENDIE & MULTIRISQUES": "INCENDIE & MULTIRISQUES",
}

# Budget annuel IARD par branche (FCFA)
BUDGET_PAR_BRANCHE = {
    "AUTOMOBILE": 5_717_334_000,
    "SANTE": 2_000_000_000,
    "INCENDIE & MULTIRISQUES": 1_500_000_000,
    "CAUTION": 800_000_000,
    "RESPONSABILITE CIVILE": 100_000_000,
    "TRANSPORT": 100_000_000,
}


def budget_from_lob(lob_value: str) -> float:
    s = str(lob_value or "").strip().upper()
    if "AUTOMOBILE" in s or "MOTOR" in s:
        return float(BUDGET_PAR_BRANCHE["AUTOMOBILE"])
    if "SANTE" in s or "HEALTH" in s:
        return float(BUDGET_PAR_BRANCHE["SANTE"])
    if "INCENDIE" in s or "MULTIRISQUE" in s or "PROPERTY" in s:
        return float(BUDGET_PAR_BRANCHE["INCENDIE & MULTIRISQUES"])
    if "CAUTION" in s or "BOND" in s:
        return float(BUDGET_PAR_BRANCHE["CAUTION"])
    if "RESPONSABILITE" in s or "LIABILITY" in s or s == "RC":
        return float(BUDGET_PAR_BRANCHE["RESPONSABILITE CIVILE"])
    if "TRANSPORT" in s:
        return float(BUDGET_PAR_BRANCHE["TRANSPORT"])
    return 0.0


def budget_lob_group(lob_value: str) -> str:
    s = str(lob_value or "").strip()
    su = s.upper()
    if "5-PERSONAL INJURIES" in su or "PERSONAL INJURIES" in su:
        return "2-PROPERTY & OTHERS DAMAGES"
    if "8-TRAVEL" in su or su == "TRAVEL" or "VOYAGE" in su:
        return "2-PROPERTY & OTHERS DAMAGES"
    return s if s else "INCONNU"

# =========================
# 2) Utils (lecture stable + exports)
# =========================
def normalize_col(c: str) -> str:
    s = str(c).strip().lower()
    # Retire les accents pour comparer Date d'émission == Date demission
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    # Uniformise ponctuation/underscore/apostrophes/espace
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())

def pick_col(df, candidates):
    norm_map = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col(cand)
        if key in norm_map:
            return norm_map[key]
    return None

def read_file(uploaded_file, name: str | None = None) -> pd.DataFrame:
    name = (name or getattr(uploaded_file, "name", "")).lower()
    if name.endswith(".csv"):
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";", encoding="utf-8")
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=",", encoding="utf-8")
    if name.endswith(".xlsx") or name.endswith(".xls"):
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
    raise ValueError("Format non supporté. Utilise CSV ou Excel.")

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Export") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


@st.cache_data(show_spinner=False)
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

# =========================
# 3) Polars helpers
# =========================
def pl_parse_date(col: pl.Expr) -> pl.Expr:
    s = col.cast(pl.Utf8, strict=False).str.strip_chars()
    s = s.str.replace_all("'", "")
    parsed = (
        s.str.strptime(pl.Date, "%d/%m/%Y", strict=False)
         .fill_null(s.str.strptime(pl.Date, "%d-%m-%Y", strict=False))
         .fill_null(s.str.strptime(pl.Date, "%Y-%m-%d", strict=False))
         .fill_null(s.str.strptime(pl.Date, "%Y-%m-%d %H:%M:%S", strict=False))
         .fill_null(s.str.strptime(pl.Date, "%d/%m/%Y %H:%M:%S", strict=False))
         .fill_null(s.str.strptime(pl.Date, "%d-%m-%Y %H:%M:%S", strict=False))
         .fill_null(
             s.str.strptime(pl.Datetime, "%Y-%m-%d %H:%M:%S%.f", strict=False)
              .dt.date()
         )
         .fill_null(
             s.str.strptime(pl.Datetime, "%Y-%m-%d %H:%M:%S%.fZ", strict=False)
              .dt.date()
         )
    )
    # Excel date serial numbers (days since 1899-12-30)
    excel_serial = pl.lit(date(1899, 12, 30)) + pl.duration(days=col.cast(pl.Float64, strict=False))
    return parsed.fill_null(excel_serial)

def pl_parse_float(col: pl.Expr) -> pl.Expr:
    return (
        col.cast(pl.Utf8, strict=False)
           .str.replace_all(" ", "")
           .str.replace_all("\u00A0", "")
           .str.replace_all(",", ".")
           .cast(pl.Float64, strict=False)
    )

def pl_parse_pct(col: pl.Expr) -> pl.Expr:
    return (
        col.cast(pl.Utf8, strict=False)
           .str.replace_all(" ", "")
           .str.replace_all("\u00A0", "")
           .str.replace_all(",", ".")
           .str.replace_all("%", "")
           .cast(pl.Float64, strict=False)
    )


def parse_dates_any(series: pd.Series) -> pd.Series:
    """
    Convertit une colonne en datetime (ISO/texte/timestamps/Excel),
    puis renvoie un datetime (NaT si impossible).
    """
    s = series.copy()
    s = s.astype("string").str.replace("'", "", regex=False)
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    num = pd.to_numeric(s, errors="coerce")
    mask_num = dt.isna() & num.notna()
    if mask_num.any():
        dt.loc[mask_num] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask_num], unit="D")
    return dt


def to_ddmmyyyy(series: pd.Series) -> pd.Series:
    """Renvoie un texte 'dd/mm/yyyy'."""
    dt = parse_dates_any(series)
    return dt.dt.strftime("%d/%m/%Y")


def compute_summary(df: pl.DataFrame) -> dict:
    return df.select([
        pl.len().alias("total"),
        (pl.col("Branche_rectifiee") == "AUTOMOBILE").sum().alias("nb_auto"),
        (pl.col("Controle (Branche AUTO & Durée contrat)") == "Revoir date effet & expiration").sum().alias("nb_auto_bad"),
        (pl.col("Controle (Date d'effet / Date d'expiration)") == "Date invalide").sum().alias("nb_date_invalid"),
        (pl.col("Controle (Date d'effet / Date d'expiration)") == "Inversion dates").sum().alias("nb_date_inv"),
        ((pl.col("_EFFET_PARSED").is_not_null()) & (pl.col("_EXP_PARSED").is_not_null()) & (pl.col("_EFFET_PARSED") < pl.col("_EXP_PARSED"))).sum().alias("nb_effet_inf_echeance"),
        pl.col("_TTC_NUM").fill_null(0).sum().alias("sum_ttc"),
        pl.col("_PRIME_NUM").fill_null(0).sum().alias("sum_prime"),
        (pl.col("_PRIME_NUM").fill_null(0) + pl.col("_ACCESS_NUM").fill_null(0)).sum().alias("sum_ca"),
    ]).to_dicts()[0]

# =========================
# 4) UI
# =========================
with st.sidebar:
    st.markdown('<div class="sidebar-card"><h2>👁️ DOSE EYES</h2><div>Contrôle production</div></div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "📤 Télécharge ton fichier (Excel ou CSV)",
        type=["xlsx", "xls", "csv"],
    )
    nav = st.radio(
        "Navigation",
        ["Dashboard IARD", "Vue d'ensemble", "AUTO Durée", "Dates Effet/Expiration", "Formats", "Exports"],
        index=0,
    )
if not uploaded:
    st.info("Charge un fichier pour lancer les contrôles.")
    st.stop()

file_bytes = uploaded.getvalue()
file_hash = hashlib.md5(file_bytes).hexdigest()

if st.session_state.get("file_hash") != file_hash:
    try:
        df_pd = read_file(BytesIO(file_bytes), uploaded.name)
    except Exception as e:
        st.error(f"Erreur lecture fichier : {e}")
        st.stop()

    df_pd = df_pd.copy()
    # Polars via PyArrow n'aime pas les colonnes "object" avec types mixtes (int/str).
    # On normalise ces colonnes en chaînes pour sécuriser la conversion.
    for col in df_pd.columns:
        if df_pd[col].dtype == "object":
            df_pd[col] = df_pd[col].astype("string")

    st.session_state.file_hash = file_hash
    st.session_state.df_pd = df_pd
else:
    df_pd = st.session_state.df_pd

st.success(f"Fichier chargé : {uploaded.name} — {df_pd.shape[0]:,} lignes, {df_pd.shape[1]} colonnes")

with st.expander("🛠️ Debug (colonnes / aperçu)", expanded=False):
    st.write("**Colonnes détectées :**")
    st.write(list(df_pd.columns))
    st.write("**Aperçu (10 premières lignes) :**")
    st.dataframe(df_pd.head(10), use_container_width=True)
    st.write("**Types :**")
    st.write(df_pd.dtypes)

if df_pd.shape[1] == 1:
    st.warning("⚠️ Ton fichier a été lu avec 1 seule colonne (souvent un problème de séparateur CSV).")
    st.info("➡️ Essaie de fournir un Excel, ou réexporter en CSV séparé par ';'.")
    st.stop()

# =========================
# 5) Détection colonnes requises
# =========================
col_branche = pick_col(df_pd, ["Branche", "branche"])
col_effet = pick_col(
    df_pd,
    [
        "Date d'effet",
        "date d'effet",
        "date effet",
        "date_deffet",
        "date deffet",
        "date d effet",
    ],
)
col_exp = pick_col(
    df_pd,
    [
        "Date d'expiration",
        "date d'expiration",
        "date expiration",
        "date_dexpiration",
        "date dexpiration",
        "date d expiration",
    ],
)
col_prime = pick_col(df_pd, ["Prime nette", "Prime nette ", "prime nette", "prime nette "])
col_ttc = pick_col(df_pd, ["Total prime TTC", "Total Prime TTC", "Total prime ttc", "Total prime TTC "])

missing = []
if col_branche is None: missing.append("Branche")
if col_effet is None: missing.append("Date d'effet")
if col_exp is None: missing.append("Date d'expiration")
if col_prime is None: missing.append("Prime nette")

if missing:
    st.error(f"Colonnes introuvables : {missing}")
    st.stop()

# Optionnelle mais utile pour la comparaison Prime nette vs TTC
if col_ttc is None:
    st.warning("Colonne 'Total prime TTC' introuvable : contrôle TTC disponible en mode 'VIDE'.")

# Optionnelles
col_police = pick_col(df_pd, ["N* Police", "N° Police", "N°Police", "N*Police", "Num Police", "Police"])
col_client = pick_col(df_pd, ["Nom du client", "Nom client", "Client"])
col_inter = pick_col(df_pd, ["Intermediaire", "Intermédiaire", "Intermediaire "])
col_channel = pick_col(df_pd, ["Channels", "Channel", "Canal"])
col_emission = pick_col(df_pd, ["Date d'emission", "Date d emission", "Date emission", "Date_emission", "Date demission"])
col_capitaux = pick_col(df_pd, ["Capitaux", "Capital", "Capitaux assurés", "Capitaux assures"])
col_accessoires = pick_col(df_pd, ["Accessoires", "Accessory", "Accessories"])
col_cession_pool = pick_col(df_pd, ["Cession Pool TPV 40%", "Cession Pool TPV 40", "Cession Pool TPV"])
col_taxes = pick_col(df_pd, ["Taxes", "Taxe"])
col_comm_brute = pick_col(df_pd, ["Com brute", "Commission brute", "Com. brute", "Com Brute"])
col_taux_comm = pick_col(df_pd, ["Taux de Commission", "Taux commission", "Taux_commission"])
col_taux_taxes = pick_col(df_pd, ["Taux de taxes", "Taux taxes", "Taux_taxes"])
col_gros = pick_col(df_pd, ["Gros contrats", "Gros contrat", "Gros_Contrats", "Gros Contrats"])
col_lob = "Line of Business" if "Line of Business" in df_pd.columns else None
col_statut = pick_col(df_pd, ["Statut", "STATUT"])

# =========================
# 6) Polars Lazy pipeline
# =========================
today_year = pd.Timestamp.today().year
lf = pl.from_pandas(df_pd).lazy()

branche_norm_expr = (
    pl.col(col_branche)
      .cast(pl.Utf8, strict=False)
      .str.strip_chars()
      .str.replace_all(r"\s+", " ")
      .str.to_uppercase()
)

lf2 = (
    lf.with_columns([
        branche_norm_expr.alias("_BRANCHE_NORM"),
        branche_norm_expr.replace(BRANCHE_MAPPING, default=branche_norm_expr).alias("Branche_rectifiee"),
        pl.col(col_effet).alias("_EFFET_RAW"),
        pl.col(col_exp).alias("_EXP_RAW"),
        (pl.col(col_emission) if col_emission else pl.lit(None)).alias("_EMISSION_RAW"),
        pl_parse_date(pl.col(col_effet)).alias("_EFFET_PARSED"),
        pl_parse_date(pl.col(col_exp)).alias("_EXP_PARSED"),
        (pl_parse_date(pl.col(col_emission)) if col_emission else pl.lit(None)).alias("_EMISSION_PARSED"),
        pl_parse_float(pl.col(col_prime)).alias("_PRIME_NUM"),
        (pl_parse_float(pl.col(col_ttc)) if col_ttc else pl.lit(None)).alias("_TTC_NUM"),
        (pl_parse_float(pl.col(col_accessoires)) if col_accessoires else pl.lit(None)).alias("_ACCESS_NUM"),
        (pl_parse_float(pl.col(col_cession_pool)) if col_cession_pool else pl.lit(None)).alias("_CESSION_POOL_NUM"),
        (pl_parse_float(pl.col(col_comm_brute)) if col_comm_brute else pl.lit(None)).alias("_COM_BRUTE_NUM"),
    ])
    .with_columns([
        (pl.col("_EXP_PARSED") - pl.col("_EFFET_PARSED")).dt.total_days().alias("Nombre de jour")
    ])
)

# =========================
# 7) Contrôle 1 : AUTO Durée
# =========================
branche_auto = pl.col("Branche_rectifiee") == "AUTOMOBILE"
cond_auto_ok = branche_auto & (pl.col("Nombre de jour") < 366)
cond_auto_bad = branche_auto & (pl.col("Nombre de jour") >= 366)
invalid_dates_for_auto = pl.col("_EFFET_PARSED").is_null() | pl.col("_EXP_PARSED").is_null() | pl.col("Nombre de jour").is_null()

lf2 = lf2.with_columns([
    pl.when(invalid_dates_for_auto).then(pl.lit("AUTRE BRANCHE"))
     .when(cond_auto_ok).then(pl.lit("Correcte"))
     .when(cond_auto_bad).then(pl.lit("Revoir date effet & expiration"))
     .otherwise(pl.lit("AUTRE BRANCHE"))
     .alias("Controle (Branche AUTO & Durée contrat)")
])

# =========================
# 8) Contrôle 2 : Dates Effet/Expiration
# =========================
date_effet_invalide = pl.col("_EFFET_PARSED").is_null() & pl.col("_EFFET_RAW").is_not_null()
date_exp_invalide = pl.col("_EXP_PARSED").is_null() & pl.col("_EXP_RAW").is_not_null()
cond_date_invalide = date_effet_invalide | date_exp_invalide
cond_inversion = (~cond_date_invalide) & (pl.col("_EXP_PARSED") <= pl.col("_EFFET_PARSED"))

lf2 = lf2.with_columns([
    pl.when(cond_date_invalide).then(pl.lit("Date invalide"))
     .when(cond_inversion).then(pl.lit("Inversion dates"))
     .otherwise(pl.lit("Correcte"))
     .alias("Controle (Date d'effet / Date d'expiration)")
])

# =========================
# 8b) Contrôle 3 : Présence Prime nette + comparaison TTC
# =========================
lf2 = lf2.with_columns([
    pl.when(pl.col("_PRIME_NUM").is_null())
      .then(pl.lit("Prime non rensignée"))
      .otherwise(pl.lit("Bon"))
      .alias("Controle (Présence / absence)"),
    pl.when(pl.col("_TTC_NUM").is_not_null() & pl.col("_PRIME_NUM").is_not_null() & (pl.col("_TTC_NUM") > pl.col("_PRIME_NUM")))
      .then(pl.lit("Bon"))
    .when(pl.col("_TTC_NUM").is_not_null() & pl.col("_PRIME_NUM").is_not_null() & (pl.col("_TTC_NUM") == pl.col("_PRIME_NUM")))
      .then(pl.lit("P nette = P TTC"))
    .when(pl.col("_TTC_NUM").is_not_null() & pl.col("_PRIME_NUM").is_not_null() & (pl.col("_TTC_NUM") < pl.col("_PRIME_NUM")))
      .then(pl.lit("P nette > P TTC"))
    .otherwise(pl.lit("VIDE"))
      .alias("Controle (Prime nette vs TTC)"),
    pl.when(pl.col("_PRIME_NUM").is_not_null() & pl.col("_TTC_NUM").is_not_null() & (pl.col("_PRIME_NUM") < pl.col("_TTC_NUM")))
      .then(pl.lit("Bon"))
      .when(pl.col("_PRIME_NUM").is_not_null() & pl.col("_TTC_NUM").is_not_null() & (pl.col("_PRIME_NUM") >= pl.col("_TTC_NUM")))
      .then(pl.lit("Anomalie"))
      .otherwise(pl.lit("VIDE"))
      .alias("Controle (Prime nette < Prime TTC)"),
])

# =========================
# 8c) Contrôle 4 : Formats (dates / montants / taux)
# =========================
date_cols = [c for c in [col_emission, col_effet, col_exp] if c]
amount_cols = [c for c in [col_capitaux, col_prime, col_accessoires, col_taxes, col_ttc, col_comm_brute] if c]
rate_cols = [c for c in [col_taux_comm, col_taux_taxes] if c]

fmt_checks = []
for c in date_cols:
    fmt_checks.append(
        (pl_parse_date(pl.col(c)).is_null() & pl.col(c).is_not_null()).alias(f"_bad_date_{c}")
    )
for c in amount_cols:
    fmt_checks.append(
        (pl_parse_float(pl.col(c)).is_null() & pl.col(c).is_not_null()).alias(f"_bad_amount_{c}")
    )
for c in rate_cols:
    fmt_checks.append(
        (pl_parse_pct(pl.col(c)).is_null() & pl.col(c).is_not_null()).alias(f"_bad_rate_{c}")
    )

if fmt_checks:
    lf2 = lf2.with_columns(fmt_checks)

if st.session_state.get("file_hash") != file_hash or "df_pl" not in st.session_state:
    st.session_state.df_pl = lf2.collect()
df_pl = st.session_state.df_pl

# =========================
# 8d) Filtre Date d'émission (sidebar)
# =========================
df_dash = df_pl
start_em = None
end_em = None
if col_emission:
    min_em = df_pl.select(pl.col("_EMISSION_PARSED").min()).item()
    max_em = df_pl.select(pl.col("_EMISSION_PARSED").max()).item()
    with st.sidebar:
        st.markdown('<div class="sidebar-card"><h3>📅 Filtre Date d’émission</h3></div>', unsafe_allow_html=True)
    if min_em and max_em:
        with st.sidebar:
            date_range = st.date_input(
                "Période",
                value=(min_em, max_em),
                min_value=min_em,
                max_value=max_em,
                key="emission_range",
            )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_em, end_em = date_range
            df_dash = df_pl.filter(
                (pl.col("_EMISSION_PARSED").is_not_null()) &
                (pl.col("_EMISSION_PARSED") >= pl.lit(start_em)) &
                (pl.col("_EMISSION_PARSED") <= pl.lit(end_em))
            )
    else:
        with st.sidebar:
            st.info("Date d’émission non détectée ou valeurs non valides.")
else:
    with st.sidebar:
        st.info("Colonne Date d’émission introuvable.")

# Date d'évaluation (B4) utilisée pour les contrôles "arrivés à échéance"
with st.sidebar:
    st.markdown('<div class="sidebar-card"><h3>🗓️ Date d’évaluation (B4)</h3></div>', unsafe_allow_html=True)
    eval_date_b4 = st.date_input(
        "Date d’évaluation",
        value=date.today(),
        key="eval_date_b4",
    )

if nav == "Dashboard IARD":
    st.markdown(
        """
        <style>
          .iard-head{background:linear-gradient(90deg,#1f345a,#294679);padding:10px 16px;border-radius:10px;margin-bottom:10px;}
          .iard-head h1{margin:0;text-align:center;color:#fff;font-size:44px;font-weight:900;letter-spacing:.8px;}
          .iard-filter-wrap{background:#eef1f6;border:1px solid #d8dee8;border-radius:10px;padding:10px;margin-bottom:10px;}
          .iard-kpi-grid{display:grid;grid-template-columns:repeat(7,minmax(200px,1fr));gap:10px;margin-bottom:12px;}
          .iard-kpi{background:#fff;border:1px solid #d3d9e5;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,.08);}
          .iard-kpi-top{background:linear-gradient(90deg,#1f4b83,#284f8e);color:#fff;font-weight:900;padding:8px 10px;font-size:17px;text-align:center;min-height:38px;display:flex;align-items:center;justify-content:center;}
          .iard-kpi-body{padding:10px 10px 12px 10px;}
          .iard-kpi-val-row{display:flex;align-items:center;justify-content:center;gap:8px;margin-top:4px;min-height:74px;}
          .iard-kpi-val{
            font-size:clamp(34px,2.1vw,50px);
            font-weight:900;
            color:#0f2344;
            line-height:1;
            white-space:nowrap;
            letter-spacing:0.2px;
          }
          .iard-kpi-arr{font-size:clamp(28px,1.6vw,40px);line-height:1;font-weight:900;white-space:nowrap;}
          .iard-kpi-sub{font-size:clamp(13px,0.85vw,15px);font-weight:900;margin-top:6px;text-align:center;white-space:nowrap;}
          .iard-up{color:#14843b;} .iard-down{color:#ce2e2e;} .iard-mid{color:#374151;}
          .iard-pan{background:#fff;border:1px solid #d3d9e5;border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,.08);}
          .iard-pan h4{margin:0;background:linear-gradient(90deg,#1f4b83,#284f8e);color:#fff;padding:7px 10px;border-radius:8px 8px 0 0;font-size:16px;font-weight:800;text-align:center;}
          .iard-pan-body{padding:8px;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<div class='iard-head'><h1>TABLEAU DE BORD IARD</h1></div>", unsafe_allow_html=True)

    dfi = df_dash.to_pandas()
    if dfi.empty:
        st.info("Aucune donnée après filtre.")
        st.stop()

    dfi["_EMISSION_PARSED"] = pd.to_datetime(dfi.get("_EMISSION_PARSED"), errors="coerce")
    dfi["_PRIME_NUM"] = pd.to_numeric(dfi.get("_PRIME_NUM"), errors="coerce").fillna(0)
    dfi["_ACCESS_NUM"] = pd.to_numeric(dfi.get("_ACCESS_NUM"), errors="coerce").fillna(0)
    dfi["_TTC_NUM"] = pd.to_numeric(dfi.get("_TTC_NUM"), errors="coerce").fillna(0)
    dfi["_COM_BRUTE_NUM"] = pd.to_numeric(dfi.get("_COM_BRUTE_NUM"), errors="coerce").fillna(0)
    dfi["_CA"] = dfi["_PRIME_NUM"] + dfi["_ACCESS_NUM"]

    f1, f2, f3, f4, f5 = st.columns(5)
    with f1:
        months = sorted(dfi["_EMISSION_PARSED"].dropna().dt.to_period("M").astype(str).unique().tolist())
        month_sel = st.selectbox("Période", ["Tout"] + months, index=0, key="iard_m")
    with f2:
        b_col = "Branche_rectifiee" if "Branche_rectifiee" in dfi.columns else col_branche
        b_vals = sorted(dfi[b_col].astype("string").fillna("INCONNU").unique().tolist()) if b_col else []
        b_sel = st.selectbox("Branche", ["Tous"] + b_vals, index=0, key="iard_b")
    with f3:
        ch_vals = sorted(dfi[col_channel].astype("string").fillna("INCONNU").unique().tolist()) if col_channel and col_channel in dfi.columns else []
        ch_sel = st.selectbox("Canal", ["Tous"] + ch_vals, index=0, key="iard_ch")
    with f4:
        inter_vals = sorted(dfi[col_inter].astype("string").fillna("INCONNU").unique().tolist()) if col_inter and col_inter in dfi.columns else []
        inter_sel = st.selectbox("Intermédiaire", ["Tous"] + inter_vals, index=0, key="iard_inter")
    with f5:
        st_vals = sorted(dfi[col_statut].astype("string").fillna("INCONNU").unique().tolist()) if col_statut and col_statut in dfi.columns else []
        st_sel = st.selectbox("Statut", ["Tous"] + st_vals, index=0, key="iard_st")

    if b_sel != "Tous" and b_col:
        dfi = dfi[dfi[b_col].astype("string") == b_sel].copy()
    if ch_sel != "Tous" and col_channel and col_channel in dfi.columns:
        dfi = dfi[dfi[col_channel].astype("string") == ch_sel].copy()
    if inter_sel != "Tous" and col_inter and col_inter in dfi.columns:
        dfi = dfi[dfi[col_inter].astype("string") == inter_sel].copy()
    if st_sel != "Tous" and col_statut and col_statut in dfi.columns:
        dfi = dfi[dfi[col_statut].astype("string") == st_sel].copy()

    # Base de référence (avant filtre de mois) pour calculer M vs M-1
    dfi_scope_pre_month = dfi.copy()
    if month_sel != "Tout":
        dfi = dfi[dfi["_EMISSION_PARSED"].dt.to_period("M").astype(str) == month_sel].copy()

    if dfi.empty:
        st.warning("Aucune donnée sur ce périmètre.")
        st.stop()

    if month_sel != "Tout":
        m_cur = pd.Period(month_sel, freq="M")
        m_prev = m_cur - 12
        dprev = df_dash.to_pandas()
        dprev["_EMISSION_PARSED"] = pd.to_datetime(dprev.get("_EMISSION_PARSED"), errors="coerce")
        dprev["_PRIME_NUM"] = pd.to_numeric(dprev.get("_PRIME_NUM"), errors="coerce").fillna(0)
        dprev["_ACCESS_NUM"] = pd.to_numeric(dprev.get("_ACCESS_NUM"), errors="coerce").fillna(0)
        dprev["_CA"] = dprev["_PRIME_NUM"] + dprev["_ACCESS_NUM"]
        dprev = dprev[dprev["_EMISSION_PARSED"].dt.to_period("M") == m_prev].copy()
    else:
        dprev = pd.DataFrame(columns=dfi.columns)

    pol_col = col_police if col_police and col_police in dfi.columns else None
    n_cur = int(dfi[pol_col].astype("string").replace("", pd.NA).dropna().nunique()) if pol_col else int(len(dfi))
    n_prev = int(dprev[pol_col].astype("string").replace("", pd.NA).dropna().nunique()) if pol_col and not dprev.empty else 0
    ca_cur = float(dfi["_CA"].sum())
    ca_prev = float(dprev["_CA"].sum()) if not dprev.empty else 0.0
    prime_moy = (ca_cur / n_cur) if n_cur else 0.0
    prime_moy_prev = (ca_prev / n_prev) if n_prev else 0.0

    if col_statut and col_statut in dfi.columns:
        ren_cur = int(dfi[col_statut].astype("string").str.upper().str.contains("RENOUV", na=False).sum())
        ren_prev = int(dprev[col_statut].astype("string").str.upper().str.contains("RENOUV", na=False).sum()) if not dprev.empty and col_statut in dprev.columns else 0
    else:
        ren_cur, ren_prev = 0, 0
    tx_ren_cur = (ren_cur / n_cur) if n_cur else 0.0
    tx_ren_prev = (ren_prev / n_prev) if n_prev else 0.0

    sin_col = pick_col(df_pd, ["SINISTRES_PAYES", "SINISTRES_PAYE", "MONTANT_SINISTRE", "SINISTRES"])
    if sin_col and sin_col in dfi.columns:
        sin_cur = float(pd.to_numeric(dfi[sin_col], errors="coerce").fillna(0).sum())
        sin_prev = float(pd.to_numeric(dprev[sin_col], errors="coerce").fillna(0).sum()) if not dprev.empty and sin_col in dprev.columns else 0.0
    else:
        sin_cur = float(dfi["_TTC_NUM"].sum() * 0.12)
        sin_prev = float(dprev["_TTC_NUM"].sum() * 0.12) if not dprev.empty else 0.0

    loss_cur = (sin_cur / ca_cur) if ca_cur else 0.0
    loss_prev = (sin_prev / ca_prev) if ca_prev else 0.0
    exp_cur = (float(dfi["_COM_BRUTE_NUM"].sum()) / ca_cur) if ca_cur else 0.0
    exp_prev = (float(dprev["_COM_BRUTE_NUM"].sum()) / ca_prev) if ca_prev else 0.0
    comb_cur = loss_cur + exp_cur
    comb_prev = loss_prev + exp_prev

    def _delta(v, p):
        if p and p != 0:
            d = (v - p) / p
        else:
            d = 1.0 if v > 0 else 0.0
        return d

    def _fmt_int(v):
        return f"{int(round(v)):,}".replace(",", " ")

    def _fmt_m(v):
        return f"{(float(v)/1_000_000):,.1f}".replace(",", " ").replace(".", ",") + " M"

    def _kpi_html(title: str, value_txt: str, delta_ratio: float) -> str:
        is_up = delta_ratio >= 0
        arr = "▲" if is_up else "▼"
        arr_cls = "iard-up" if is_up else "iard-down"
        sub = f"{arr} {delta_ratio*100:+.1f}% vs N-1"
        return (
            "<div class='iard-kpi'>"
            f"<div class='iard-kpi-top'>{title}</div>"
            "<div class='iard-kpi-body'>"
            f"<div class='iard-kpi-val-row'><div class='iard-kpi-val'>{value_txt}</div><div class='iard-kpi-arr {arr_cls}'>{arr}</div></div>"
            f"<div class='iard-kpi-sub {arr_cls}'>{sub}</div>"
            "</div></div>"
        )

    kpis = [
        _kpi_html("Chiffre d’Affaires", _fmt_m(ca_cur), _delta(ca_cur, ca_prev)),
        _kpi_html("Nombre de Contrats", _fmt_int(n_cur), _delta(n_cur, n_prev)),
        _kpi_html("Prime Moyenne", _fmt_int(prime_moy), _delta(prime_moy, prime_moy_prev)),
        _kpi_html("% de Renouvellement", f"{tx_ren_cur*100:.1f}%".replace(".", ","), _delta(tx_ren_cur, tx_ren_prev)),
        _kpi_html("Sinistres Payés", _fmt_m(sin_cur), _delta(sin_cur, sin_prev)),
        _kpi_html("Loss Ratio", f"{loss_cur*100:.1f}%".replace(".", ","), _delta(loss_cur, loss_prev)),
        _kpi_html("Combined Ratio", f"{comb_cur*100:.1f}%".replace(".", ","), _delta(comb_cur, comb_prev)),
    ]
    cards_html = "".join(kpis)
    st.markdown(f"<div class='iard-kpi-grid'>{cards_html}</div>", unsafe_allow_html=True)

    p1, p2, p3 = st.columns([1.15, 1.0, 1.0])
    with p1:
        st.markdown("<div class='iard-pan'><h4>Évolution CA & Contrats (M FCFA)</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
        if col_lob and col_lob in dfi_scope_pre_month.columns:
            evo_src = dfi_scope_pre_month[[col_lob, "_EMISSION_PARSED", "_CA"]].copy()
            evo_src["_EMISSION_PARSED"] = pd.to_datetime(evo_src["_EMISSION_PARSED"], errors="coerce")
            evo_src = evo_src.dropna(subset=["_EMISSION_PARSED"]).copy()
            if not evo_src.empty:
                evo_src["_YM"] = evo_src["_EMISSION_PARSED"].dt.to_period("M")
                evo_src["_LOB_GRP"] = evo_src[col_lob].astype("string").fillna("INCONNU").map(budget_lob_group)

                monthly = (
                    evo_src.groupby("_YM", dropna=False)
                    .agg(
                        CA=("_CA", "sum"),
                        Contrats=("_CA", "size"),
                    )
                    .reset_index()
                )
                monthly["CA_M"] = monthly["CA"] / 1_000_000.0

                lob_month = evo_src[["_YM", "_LOB_GRP"]].drop_duplicates()
                lob_month["BudgetAnnuel"] = lob_month["_LOB_GRP"].map(budget_from_lob).fillna(0.0)
                obj_month = (
                    lob_month.groupby("_YM", dropna=False)["BudgetAnnuel"]
                    .sum()
                    .reset_index(name="ObjAnnuel")
                )
                obj_month["Objectif_M"] = (obj_month["ObjAnnuel"] / 12.0) / 1_000_000.0

                monthly = monthly.merge(obj_month[["_YM", "Objectif_M"]], on="_YM", how="left").fillna(0.0)
                monthly = monthly.sort_values("_YM")
                monthly["Mois"] = monthly["_YM"].astype(str)

                fig_evo = go.Figure()
                fig_evo.add_bar(
                    x=monthly["Mois"],
                    y=monthly["CA_M"],
                    name="CA réalisé (M FCFA)",
                    marker=dict(color="#1f4b83"),
                )
                fig_evo.add_scatter(
                    x=monthly["Mois"],
                    y=monthly["Objectif_M"],
                    mode="lines+markers",
                    name="Objectif (M FCFA)",
                    line=dict(color="#f2b705", width=3),
                )
                fig_evo.add_scatter(
                    x=monthly["Mois"],
                    y=monthly["Contrats"],
                    mode="lines+markers",
                    name="Contrats",
                    line=dict(color="#2f8f46", width=2, dash="dot"),
                    yaxis="y2",
                )
                fig_evo.update_layout(
                    height=320,
                    margin=dict(l=10, r=10, t=5, b=10),
                    barmode="group",
                    legend=dict(orientation="h", y=1.12, x=0),
                    yaxis=dict(title="M FCFA"),
                    yaxis2=dict(title="Contrats", overlaying="y", side="right", showgrid=False),
                    xaxis=dict(title="Mois"),
                )
                st.plotly_chart(fig_evo, use_container_width=True)
            else:
                st.info("Aucune Date d’émission valide.")
        else:
            st.info("Colonne Line of Business introuvable.")
        st.markdown("</div></div>", unsafe_allow_html=True)
    with p2:
        st.markdown("<div class='iard-pan'><h4>Répartition du CA</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
        if col_lob and col_lob in dfi.columns:
            rep = dfi.groupby(col_lob, dropna=False)["_CA"].sum().reset_index().rename(columns={col_lob: "LOB"})
            rep = rep.sort_values("_CA", ascending=False)
            if len(rep) > 5:
                top = rep.head(4).copy()
                other = pd.DataFrame({"LOB": ["Autres"], "_CA": [rep.iloc[4:]["_CA"].sum()]})
                rep = pd.concat([top, other], ignore_index=True)
            fig_rep = px.pie(rep, values="_CA", names="LOB", hole=0.5, color_discrete_sequence=px.colors.sequential.Blues_r)
            fig_rep.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_rep, use_container_width=True)
        else:
            st.info("Colonne Line of Business introuvable.")
        st.markdown("</div></div>", unsafe_allow_html=True)
    with p3:
        st.markdown("<div class='iard-pan'><h4>Top Intermédiaires</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
        if col_inter and col_inter in dfi.columns:
            topi = dfi.groupby(col_inter, dropna=False).agg(CA=("_CA", "sum"), Contrats=("_CA", "size")).reset_index()
            topi = topi.sort_values("CA", ascending=False).head(8)
            fig_top = px.bar(topi.sort_values("CA"), x="CA", y=col_inter, orientation="h", color_discrete_sequence=["#1f4b83"])
            fig_top.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10), showlegend=False, yaxis_title="")
            st.plotly_chart(fig_top, use_container_width=True)
        else:
            st.info("Colonne Intermédiaire introuvable.")
        st.markdown("</div></div>", unsafe_allow_html=True)

    r1, r2, r3 = st.columns([1.0, 1.0, 1.0])
    with r1:
        st.markdown("<div class='iard-pan'><h4>Loss Ratio par Branche</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
        if col_lob and col_lob in dfi.columns:
            lr = dfi.groupby(col_lob, dropna=False).agg(CA=("_CA", "sum"), Sinistres=("_TTC_NUM", "sum"), Com=("_COM_BRUTE_NUM", "sum")).reset_index()
            lr = lr.rename(columns={col_lob: "LOB"}).sort_values("CA", ascending=False).head(8)
            fig_lr = px.bar(lr, x="LOB", y=["Sinistres", "Com"], barmode="stack", color_discrete_sequence=["#1f4b83", "#f2b705"])
            fig_lr.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), xaxis_tickangle=-25)
            st.plotly_chart(fig_lr, use_container_width=True)
        else:
            st.info("Colonne Line of Business introuvable.")
        st.markdown("</div></div>", unsafe_allow_html=True)
    with r2:
        st.markdown("<div class='iard-pan'><h4>Fréquence & Sévérité</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
        freq = (ren_cur / n_cur) if n_cur else 0.0
        sev = (sin_cur / max(1, ren_cur))
        g1, g2 = st.columns(2)
        with g1:
            fig_g1 = go.Figure(go.Indicator(mode="gauge+number", value=freq * 100, number={'suffix': "%"}, title={'text': "Fréquence"}, gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#1f4b83"}}))
            fig_g1.update_layout(height=240, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig_g1, use_container_width=True)
        with g2:
            fig_g2 = go.Figure(go.Indicator(mode="gauge+number", value=sev, title={'text': "Coût moyen"}, gauge={'axis': {'range': [0, max(1.0, sev * 1.4)]}, 'bar': {'color': "#f2b705"}}))
            fig_g2.update_layout(height=240, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig_g2, use_container_width=True)
        st.markdown("</div></div>", unsafe_allow_html=True)
    with r3:
        st.markdown("<div class='iard-pan'><h4>Qualité Sinistre</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
        q1, q2 = st.columns(2)
        with q1:
            qd = pd.DataFrame({"Type": ["Litiges/Rejets", "Qualifiés"], "N": [max(0, int(n_cur * 0.28)), max(1, int(n_cur * 0.72))]})
            fig_q1 = px.pie(qd, names="Type", values="N", hole=0.6, color_discrete_sequence=["#1f4b83", "#4f9d3a"])
            fig_q1.update_layout(height=240, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
            st.plotly_chart(fig_q1, use_container_width=True)
        with q2:
            fig_q2 = px.pie(qd, names="Type", values="N", color_discrete_sequence=["#f2b705", "#1f4b83"])
            fig_q2.update_layout(height=240, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
            st.plotly_chart(fig_q2, use_container_width=True)
        st.markdown("</div></div>", unsafe_allow_html=True)

    st.markdown("<div class='iard-pan'><h4>Alerte & Qualité des Données</h4><div class='iard-pan-body'>", unsafe_allow_html=True)
    alerts = []
    if "Controle (Date d'effet / Date d'expiration)" in dfi.columns:
        alerts.append(("Polices sans Date d'Effet", int((dfi["Controle (Date d'effet / Date d'expiration)"] == "Date invalide").sum())))
        alerts.append(("Inversion dates Effet/Expiration", int((dfi["Controle (Date d'effet / Date d'expiration)"] == "Inversion dates").sum())))
    if "Controle (Prime nette vs TTC)" in dfi.columns:
        alerts.append(("Prime nette > Prime TTC", int((dfi["Controle (Prime nette vs TTC)"] == "P nette > P TTC").sum())))
    if col_police and col_police in dfi.columns:
        dup_pol = int(dfi[col_police].astype("string").replace("", pd.NA).dropna().duplicated().sum())
        alerts.append(("Doublons N° Police", dup_pol))
    if not alerts:
        alerts = [("Aucune alerte", 0)]
    alert_df = pd.DataFrame(alerts, columns=["Alerte", "Nombre"]).sort_values("Nombre", ascending=False)
    st.dataframe(alert_df, use_container_width=True, hide_index=True)
    st.markdown("</div></div>", unsafe_allow_html=True)
    st.stop()

if nav == "Vue d'ensemble":
    summary = compute_summary(df_dash)
    eval_date = pd.Timestamp(eval_date_b4).strftime("%d/%m/%Y")

    st.subheader("📊 Résumé global")
    st.caption(f"Date d’évaluation : {eval_date}")
    st.markdown(
        f"""
        <div class="kpi-grid">
          <div class="kpi-card"><div class="kpi-label">Nombre Total de contrat</div><div class="kpi-value">{summary['total']:,}</div></div>
          <div class="kpi-card"><div class="kpi-label">Branche = Automobile</div><div class="kpi-value">{summary['nb_auto']:,}</div></div>
          <div class="kpi-card"><div class="kpi-label">AUTO - Incohérences</div><div class="kpi-value">{summary['nb_auto_bad']:,}</div></div>
          <div class="kpi-card"><div class="kpi-label">Dates invalides</div><div class="kpi-value">{summary['nb_date_invalid']:,}</div></div>
          <div class="kpi-card"><div class="kpi-label">Dates inversion</div><div class="kpi-value">{summary['nb_date_inv']:,}</div></div>
          <div class="kpi-card"><div class="kpi-label">Effet &lt; Échéance</div><div class="kpi-value">{summary['nb_effet_inf_echeance']:,}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("💰 Montants")
    ca_cur = float(summary.get("sum_ca", 0))
    ca_prev = 0.0
    if col_emission and start_em and end_em:
        prev_start = (pd.Timestamp(start_em) - pd.DateOffset(years=1)).date()
        prev_end = (pd.Timestamp(end_em) - pd.DateOffset(years=1)).date()
        prev_df = df_pl.filter(
            (pl.col("_EMISSION_PARSED").is_not_null())
            & (pl.col("_EMISSION_PARSED") >= pl.lit(prev_start))
            & (pl.col("_EMISSION_PARSED") <= pl.lit(prev_end))
        )
        ca_prev = float(compute_summary(prev_df).get("sum_ca", 0))
    if ca_prev > 0:
        ca_evol = (ca_cur / ca_prev) - 1
    else:
        ca_evol = 1.0 if ca_cur > 0 else 0.0
    ca_delta = ca_cur - ca_prev
    evol_arrow = "▲" if ca_evol >= 0 else "▼"
    evol_color = "#16a34a" if ca_evol >= 0 else "#dc2626"

    cession_cur = float(df_dash.select(pl.col("_CESSION_POOL_NUM").fill_null(0).sum()).item()) if "_CESSION_POOL_NUM" in df_dash.columns else 0.0
    cession_prev = 0.0
    if col_emission and start_em and end_em and "_CESSION_POOL_NUM" in df_pl.columns:
        prev_start = (pd.Timestamp(start_em) - pd.DateOffset(years=1)).date()
        prev_end = (pd.Timestamp(end_em) - pd.DateOffset(years=1)).date()
        prev_df_cession = df_pl.filter(
            (pl.col("_EMISSION_PARSED").is_not_null())
            & (pl.col("_EMISSION_PARSED") >= pl.lit(prev_start))
            & (pl.col("_EMISSION_PARSED") <= pl.lit(prev_end))
        )
        cession_prev = float(prev_df_cession.select(pl.col("_CESSION_POOL_NUM").fill_null(0).sum()).item())
    if cession_prev > 0:
        cession_evol = (cession_cur / cession_prev) - 1
    else:
        cession_evol = 1.0 if cession_cur > 0 else 0.0
    cession_delta = cession_cur - cession_prev
    cession_arrow = "▲" if cession_evol >= 0 else "▼"
    cession_color = "#16a34a" if cession_evol >= 0 else "#dc2626"

    com_cur = float(df_dash.select(pl.col("_COM_BRUTE_NUM").fill_null(0).sum()).item()) if "_COM_BRUTE_NUM" in df_dash.columns else 0.0
    com_prev = 0.0
    if col_emission and start_em and end_em and "_COM_BRUTE_NUM" in df_pl.columns:
        prev_start = (pd.Timestamp(start_em) - pd.DateOffset(years=1)).date()
        prev_end = (pd.Timestamp(end_em) - pd.DateOffset(years=1)).date()
        prev_df_com = df_pl.filter(
            (pl.col("_EMISSION_PARSED").is_not_null())
            & (pl.col("_EMISSION_PARSED") >= pl.lit(prev_start))
            & (pl.col("_EMISSION_PARSED") <= pl.lit(prev_end))
        )
        com_prev = float(prev_df_com.select(pl.col("_COM_BRUTE_NUM").fill_null(0).sum()).item())
    if com_prev > 0:
        com_evol = (com_cur / com_prev) - 1
    else:
        com_evol = 1.0 if com_cur > 0 else 0.0
    com_delta = com_cur - com_prev
    com_arrow = "▲" if com_evol >= 0 else "▼"
    com_color = "#16a34a" if com_evol >= 0 else "#dc2626"

    st.markdown(
        f"""
        <div class="kpi-grid-3">
          <div class="kpi-feature">
            <div class="kpi-feature-title">Cession Pool TPV 40%</div>
            <div class="kpi-feature-value">XOF {cession_cur:,.0f}</div>
            <div class="kpi-feature-sub">
              {cession_evol*100:,.1f}% Increase
              <span class="kpi-feature-evol" style="color:{cession_color};">{cession_arrow}</span>
              <span class="kpi-feature-delta">{cession_delta:,.0f}</span>
            </div>
            <div class="kpi-feature-note">Compared to same period N-1</div>
          </div>
          <div class="kpi-feature">
            <div class="kpi-feature-title">Gross Written Premium</div>
            <div class="kpi-feature-value">XOF {ca_cur:,.0f}</div>
            <div class="kpi-feature-sub">
              {ca_evol*100:,.1f}% Increase
              <span class="kpi-feature-evol" style="color:{evol_color};">{evol_arrow}</span>
              <span class="kpi-feature-delta">{ca_delta:,.0f}</span>
            </div>
            <div class="kpi-feature-note">Compared to same period N-1</div>
          </div>
          <div class="kpi-feature">
            <div class="kpi-feature-title">Com brute</div>
            <div class="kpi-feature-value">XOF {com_cur:,.0f}</div>
            <div class="kpi-feature-sub">
              {com_evol*100:,.1f}% Increase
              <span class="kpi-feature-evol" style="color:{com_color};">{com_arrow}</span>
              <span class="kpi-feature-delta">{com_delta:,.0f}</span>
            </div>
            <div class="kpi-feature-note">Compared to same period N-1</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    df_dash_pd = None
    if col_inter or col_channel or col_lob or col_statut:
        cols_needed = ["_TTC_NUM", "Branche_rectifiee"]
        for c in [col_inter, col_channel, col_lob, col_statut]:
            if c and c not in cols_needed:
                cols_needed.append(c)
        df_dash_pd = df_dash.select(cols_needed).to_pandas()
        df_dash_pd["_TTC_NUM"] = pd.to_numeric(df_dash_pd["_TTC_NUM"], errors="coerce").fillna(0)

    if col_lob and df_dash_pd is not None:
        lob_card = df_dash_pd[[col_lob, "_TTC_NUM"]].copy()
        lob_card[col_lob] = lob_card[col_lob].astype("string").fillna("INCONNU")
        agg_lob_card = (
            lob_card.groupby(col_lob, dropna=False)
            .agg(**{"Chiffre d'affaires (Prime TTC)": ("_TTC_NUM", "sum")})
            .reset_index()
            .rename(columns={col_lob: "Line of Business"})
        )
        agg_lob_card = agg_lob_card.sort_values("Chiffre d'affaires (Prime TTC)", ascending=True)
        agg_lob_card["CA_label"] = (
            agg_lob_card["Chiffre d'affaires (Prime TTC)"]
            .round(0)
            .astype(int)
            .map(lambda v: f"{v:,} FCFA".replace(",", " "))
        )
        c_left, c_mid, c_right = st.columns(3)

        with c_left:
            with st.container(border=True):
                st.markdown("#### 📈 CA par Line of Business")
                fig_lob_top = px.bar(
                    agg_lob_card,
                    y="Line of Business",
                    x="Chiffre d'affaires (Prime TTC)",
                    orientation="h",
                    text="CA_label",
                    color_discrete_sequence=["#F4C20D"],
                )
                fig_lob_top.update_traces(
                    textposition="outside",
                    cliponaxis=False,
                    marker_line_color="#C89D00",
                    marker_line_width=1,
                )
                fig_lob_top.update_layout(
                    yaxis_title="Line of Business",
                    xaxis_title="Chiffre d'affaires (Prime TTC)",
                    showlegend=False,
                    height=420,
                    margin=dict(l=20, r=90, t=10, b=20),
                )
                st.plotly_chart(fig_lob_top, use_container_width=True)

        with c_mid:
            with st.container(border=True):
                st.markdown("#### 📌 Intermédiaire : répartition")
                if col_inter:
                    inter_top = df_dash_pd[[col_inter, "_TTC_NUM"]].copy()
                    inter_series_top = inter_top[col_inter].astype("string").str.strip().str.upper()
                    categories_top = ["SIEGE", "PLATEFORME AUTO", "PLATEFORME VOYAGE"]
                    inter_top["Intermediaire"] = inter_series_top.where(inter_series_top.isin(categories_top), "AUTRES")
                    agg_inter_top = (
                        inter_top.groupby("Intermediaire", dropna=False)
                        .agg(**{"Nombre de contrats": ("Intermediaire", "size")})
                        .reset_index()
                    )
                    fig_inter_top = px.pie(
                        agg_inter_top,
                        names="Intermediaire",
                        values="Nombre de contrats",
                        color_discrete_sequence=px.colors.qualitative.Set2,
                        hole=0.25,
                    )
                    fig_inter_top.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
                    st.plotly_chart(fig_inter_top, use_container_width=True)
                else:
                    st.info("Colonne Intermédiaire introuvable.")

        with c_right:
            with st.container(border=True):
                st.markdown("#### 📌 Channels : répartition")
                if col_channel:
                    ch_top = df_dash_pd[[col_channel, "_TTC_NUM"]].copy()
                    ch_top[col_channel] = (
                        ch_top[col_channel]
                        .astype("string")
                        .str.strip()
                        .replace("", "INCONNU")
                        .fillna("INCONNU")
                    )
                    agg_ch_top = (
                        ch_top.groupby(col_channel, dropna=False)
                        .agg(**{"Nombre de contrats": (col_channel, "size")})
                        .reset_index()
                        .rename(columns={col_channel: "Channels"})
                    )
                    fig_ch_top = px.pie(
                        agg_ch_top,
                        names="Channels",
                        values="Nombre de contrats",
                        color_discrete_sequence=px.colors.qualitative.Pastel,
                        hole=0.25,
                    )
                    fig_ch_top.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
                    st.plotly_chart(fig_ch_top, use_container_width=True)
                else:
                    st.info("Colonne Channels introuvable.")

    if col_statut and df_dash_pd is not None:
        st.subheader("📌 Statut : répartition")
        st_pd = df_dash_pd[[col_statut, "_TTC_NUM"]].copy()
        st_pd[col_statut] = (
            st_pd[col_statut]
            .astype("string")
            .str.strip()
            .str.replace("RENOUVELLEMENTOUVELLEMENT", "RENOUVELLEMENT", regex=False)
            .replace("", "INCONNU")
            .fillna("INCONNU")
        )

        agg_st = (
            st_pd.groupby(col_statut, dropna=False)
            .agg(
                **{
                    "Nombre de contrats": (col_statut, "size"),
                    "Chiffre d'affaires (Prime TTC)": ("_TTC_NUM", "sum"),
                }
            )
            .reset_index()
            .rename(columns={col_statut: "Statut"})
            .sort_values("Nombre de contrats", ascending=False)
        )
        total_contrats_st = agg_st["Nombre de contrats"].sum()
        agg_st["Taux de repartition"] = np.where(
            total_contrats_st > 0, agg_st["Nombre de contrats"] / total_contrats_st, 0
        )

        display_st = agg_st.copy()
        display_st["Chiffre d'affaires (Prime TTC)"] = display_st["Chiffre d'affaires (Prime TTC)"].round(0).astype(int)
        display_st["Taux de repartition"] = (display_st["Taux de repartition"] * 100).round(1).astype(str) + "%"
        col_pie, col_budget = st.columns([0.9, 1.35])
        with col_pie:
            with st.container(border=True):
                st.markdown("#### Camembert Statut")
                fig_st = px.pie(
                    agg_st,
                    names="Statut",
                    values="Nombre de contrats",
                    color_discrete_sequence=px.colors.qualitative.Set3,
                    hole=0.25,
                )
                fig_st.update_layout(height=330, margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(fig_st, use_container_width=True)
            with st.container(border=True):
                st.markdown("#### Poids budget par Line of Business")
                if col_lob and col_lob in df_dash_pd.columns:
                    budget_share_src = df_dash_pd[[col_lob]].copy()
                    budget_share_src["_LOB_BUDGET"] = (
                        budget_share_src[col_lob]
                        .astype("string")
                        .fillna("INCONNU")
                        .map(budget_lob_group)
                    )
                    budget_share_df = (
                        budget_share_src.groupby("_LOB_BUDGET", dropna=False)
                        .size()
                        .reset_index(name="_N")
                        .rename(columns={"_LOB_BUDGET": "Line of Business"})
                    )
                    budget_share_df["Budget annuel"] = budget_share_df["Line of Business"].map(budget_from_lob).fillna(0.0)
                    budget_share_df = budget_share_df[budget_share_df["Budget annuel"] > 0].copy()
                    budget_share_df = budget_share_df.drop_duplicates(subset=["Line of Business", "Budget annuel"])
                    if not budget_share_df.empty:
                        fig_budget_share = px.pie(
                            budget_share_df,
                            names="Line of Business",
                            values="Budget annuel",
                            color_discrete_sequence=px.colors.qualitative.Set2,
                            hole=0.2,
                        )
                        fig_budget_share.update_layout(height=330, margin=dict(l=10, r=10, t=10, b=10))
                        st.plotly_chart(fig_budget_share, use_container_width=True)
                    else:
                        st.info("Aucun budget disponible pour calculer la répartition.")
                else:
                    st.info("Colonne Line of Business introuvable.")
        with col_budget:
            with st.container(border=True):
                st.markdown("#### Évolution du CA réalisé vs objectif annuel")
                if col_lob and col_lob in df_dash_pd.columns:
                    evo_src = df_dash_pd[[col_lob, "_TTC_NUM"]].copy()
                    evo_src["_LOB_BUDGET"] = (
                        evo_src[col_lob]
                        .astype("string")
                        .fillna("INCONNU")
                        .map(budget_lob_group)
                    )
                    evo_budget = (
                        evo_src.groupby("_LOB_BUDGET", dropna=False)
                        .agg(**{"CA réalisé (Prime TTC)": ("_TTC_NUM", "sum")})
                        .reset_index()
                        .rename(columns={"_LOB_BUDGET": "Line of Business"})
                    )
                    evo_budget["Line of Business"] = evo_budget["Line of Business"].astype("string").fillna("INCONNU")
                    evo_budget["Budget annuel"] = evo_budget["Line of Business"].map(budget_from_lob).fillna(0.0)
                    evo_budget["Taux de réalisation"] = np.where(
                        evo_budget["Budget annuel"] > 0,
                        evo_budget["CA réalisé (Prime TTC)"] / evo_budget["Budget annuel"],
                        0.0,
                    )
                    evo_budget = evo_budget.sort_values("Taux de réalisation", ascending=False)
                    col_left, col_right = st.columns(2)
                    n = len(evo_budget)
                    split = (n + 1) // 2
                    left_rows = evo_budget.iloc[:split]
                    right_rows = evo_budget.iloc[split:]

                    def _short_lob_name(v: str, max_len: int = 26) -> str:
                        s = str(v or "").strip()
                        return s if len(s) <= max_len else (s[: max_len - 1] + "…")

                    def _draw_budget_pies(rows_df, col_obj, key_prefix: str):
                        with col_obj:
                            for i, row in rows_df.iterrows():
                                with st.container(border=True):
                                    taux_pct = float(row["Taux de réalisation"] * 100)
                                    fill_pct = min(max(taux_pct, 0.0), 100.0)
                                    pie_color = "#16a34a" if taux_pct >= 100 else ("#f59e0b" if taux_pct >= 70 else "#ef4444")
                                    fig_evo_statut = go.Figure(
                                        data=[
                                            go.Pie(
                                                values=[fill_pct, 100 - fill_pct],
                                                hole=0.72,
                                                sort=False,
                                                marker=dict(colors=[pie_color, "#e5e7eb"]),
                                                textinfo="none",
                                                hoverinfo="skip",
                                                showlegend=False,
                                            )
                                        ]
                                    )
                                    fig_evo_statut.add_annotation(
                                        text=f"{taux_pct:.1f}%",
                                        x=0.5,
                                        y=0.5,
                                        showarrow=False,
                                        font=dict(size=17, color="#0f172a"),
                                    )
                                    fig_evo_statut.update_layout(
                                        title=dict(
                                            text=_short_lob_name(row["Line of Business"]),
                                            x=0.5,
                                            font=dict(size=12, color="#1e293b"),
                                        ),
                                        margin=dict(l=8, r=8, t=30, b=8),
                                        height=170,
                                        paper_bgcolor="white",
                                    )
                                    st.plotly_chart(fig_evo_statut, use_container_width=True, key=f"{key_prefix}_{i}")

                    _draw_budget_pies(left_rows, col_left, "evo_budget_left")
                    _draw_budget_pies(right_rows, col_right, "evo_budget_right")
                else:
                    st.info("Colonne Line of Business introuvable.")
        with st.container(border=True):
            st.markdown("#### Tableau Statut")
            st_display_small = display_st.copy()
            st_display_small = st_display_small[[
                "Statut",
                "Nombre de contrats",
                "Chiffre d'affaires (Prime TTC)",
                "Taux de repartition",
            ]]
            st.dataframe(st_display_small, use_container_width=True, hide_index=True, height=330)

    if col_inter:
        st.subheader("📌 Intermédiaire : répartition")
        inter_pd = df_dash_pd[[col_inter, "_TTC_NUM"]].copy()
        inter_series = inter_pd[col_inter].astype("string").str.strip().str.upper()
        categories = ["SIEGE", "PLATEFORME AUTO", "PLATEFORME VOYAGE"]
        inter_pd["GROUPE"] = inter_series.where(inter_series.isin(categories), "AUTRES")

        agg = (
            inter_pd.groupby("GROUPE", dropna=False)
            .agg(
                **{
                    "Nombre de contrats": ("GROUPE", "size"),
                    "Chiffre d'affaires (Prime TTC)": ("_TTC_NUM", "sum"),
                }
            )
            .reset_index()
        )
        total_contrats = agg["Nombre de contrats"].sum()
        agg["Taux de repartition"] = np.where(
            total_contrats > 0, agg["Nombre de contrats"] / total_contrats, 0
        )

        order = categories + ["AUTRES"]
        agg["GROUPE"] = pd.Categorical(agg["GROUPE"], categories=order, ordered=True)
        agg = agg.sort_values("GROUPE").rename(columns={"GROUPE": "Intermediaire"})

        display_df = agg.copy()
        display_df["Chiffre d'affaires (Prime TTC)"] = display_df["Chiffre d'affaires (Prime TTC)"].round(0).astype(int)
        display_df["Taux de repartition"] = (display_df["Taux de repartition"] * 100).round(1).astype(str) + "%"

        st.dataframe(display_df, use_container_width=True, hide_index=True)

    if col_channel:
        st.subheader("📌 Channels : répartition")
        ch_pd = df_dash_pd[[col_channel, "_TTC_NUM"]].copy()
        ch_pd[col_channel] = ch_pd[col_channel].astype("string").str.strip().replace("", "INCONNU").fillna("INCONNU")

        agg_ch = (
            ch_pd.groupby(col_channel, dropna=False)
            .agg(
                **{
                    "Nombre de contrats": (col_channel, "size"),
                    "Chiffre d'affaires (Prime TTC)": ("_TTC_NUM", "sum"),
                }
            )
            .reset_index()
            .rename(columns={col_channel: "Channels"})
            .sort_values("Nombre de contrats", ascending=False)
        )
        total_contrats_ch = agg_ch["Nombre de contrats"].sum()
        agg_ch["Taux de repartition"] = np.where(
            total_contrats_ch > 0, agg_ch["Nombre de contrats"] / total_contrats_ch, 0
        )

        display_ch = agg_ch.copy()
        display_ch["Chiffre d'affaires (Prime TTC)"] = display_ch["Chiffre d'affaires (Prime TTC)"].round(0).astype(int)
        display_ch["Taux de repartition"] = (display_ch["Taux de repartition"] * 100).round(1).astype(str) + "%"

        st.dataframe(display_ch, use_container_width=True, hide_index=True)
        fig_bar = px.bar(
            agg_ch,
            x="Channels",
            y="Chiffre d'affaires (Prime TTC)",
            title="Chiffre d'affaires par Channel",
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    if col_lob:
        st.subheader("📌 Secteurs d'activité : répartition")
        lob_pd = df_dash_pd[[col_lob, "_TTC_NUM", "Branche_rectifiee"]].copy()
        lob_pd[col_lob] = lob_pd[col_lob].astype("string").fillna("INCONNU")
        lob_pd["_TTC_NUM"] = pd.to_numeric(lob_pd["_TTC_NUM"], errors="coerce").fillna(0)

        agg_lob = (
            lob_pd.groupby(col_lob, dropna=False)
            .agg(
                **{
                    "Nombre de contrats": (col_lob, "size"),
                    "Chiffre d'affaires (Prime TTC)": ("_TTC_NUM", "sum"),
                }
            )
            .reset_index()
            .rename(columns={col_lob: "Secteur d'activité"})
            .sort_values("Nombre de contrats", ascending=False)
        )
        total_contrats_lob = agg_lob["Nombre de contrats"].sum()
        agg_lob["Taux de repartition"] = np.where(
            total_contrats_lob > 0, agg_lob["Nombre de contrats"] / total_contrats_lob, 0
        )

        display_lob = agg_lob.copy()
        display_lob["Chiffre d'affaires (Prime TTC)"] = display_lob["Chiffre d'affaires (Prime TTC)"].round(0).astype(int)
        display_lob["Taux de repartition"] = (display_lob["Taux de repartition"] * 100).round(1).astype(str) + "%"

        st.dataframe(display_lob, use_container_width=True, hide_index=True)

        # Evolution CA par Line of Business (M vs M-1) en donuts
        st.markdown("#### 📈 Evolution du chiffre d'affaires par Line of Business")
        lob_evo_src = df_pl.select([
            pl.col(col_lob).cast(pl.Utf8, strict=False).str.strip_chars().fill_null("INCONNU").alias("_LOB_EVO"),
            pl.col("_EMISSION_PARSED").alias("_EM"),
            pl.col("_TTC_NUM").fill_null(0).alias("_CA"),
        ]).to_pandas()
        lob_evo_src["_EM"] = pd.to_datetime(lob_evo_src["_EM"], errors="coerce")
        lob_evo_src = lob_evo_src.dropna(subset=["_EM"]).copy()

        if not lob_evo_src.empty:
            ref_month = lob_evo_src["_EM"].max().to_period("M")
            prev_month = ref_month - 1
            lob_evo_src["_YM"] = lob_evo_src["_EM"].dt.to_period("M")

            ca_m = (
                lob_evo_src[lob_evo_src["_YM"] == ref_month]
                .groupby("_LOB_EVO", dropna=False)["_CA"]
                .sum()
                .rename("CA_M")
            )
            ca_m1 = (
                lob_evo_src[lob_evo_src["_YM"] == prev_month]
                .groupby("_LOB_EVO", dropna=False)["_CA"]
                .sum()
                .rename("CA_M1")
            )
            evo_df = pd.concat([ca_m, ca_m1], axis=1).fillna(0).reset_index()
            evo_df["Evol"] = np.where(
                evo_df["CA_M1"] > 0,
                (evo_df["CA_M"] - evo_df["CA_M1"]) / evo_df["CA_M1"],
                np.where(evo_df["CA_M"] > 0, 1.0, 0.0),
            )
            evo_df = evo_df.sort_values("CA_M", ascending=False)

            with st.container(border=True):
                for i, row in evo_df.iterrows():
                    lob_name = str(row["_LOB_EVO"])
                    evol_pct = float(row["Evol"] * 100)
                    fill_pct = min(abs(evol_pct), 100.0)
                    ring_color = "#16a34a" if evol_pct >= 0 else "#dc2626"

                    c_name, c_ring = st.columns([1.55, 1])
                    with c_name:
                        st.markdown(f"**{lob_name}**")
                        st.caption(
                            f"CA M: {int(row['CA_M']):,} | CA M-1: {int(row['CA_M1']):,}".replace(",", " ")
                        )
                    with c_ring:
                        fig_ring = go.Figure(
                            data=[
                                go.Pie(
                                    values=[fill_pct, 100 - fill_pct],
                                    hole=0.0,
                                    sort=False,
                                    direction="clockwise",
                                    marker=dict(colors=[ring_color, "#e5e7eb"]),
                                    textinfo="none",
                                    hoverinfo="skip",
                                    showlegend=False,
                                )
                            ]
                        )
                        fig_ring.add_annotation(
                            text=f"{evol_pct:+.1f}%",
                            x=0.5,
                            y=0.5,
                            showarrow=False,
                            font=dict(size=18, color="#0f172a"),
                        )
                        fig_ring.update_layout(
                            margin=dict(l=8, r=8, t=8, b=8),
                            height=160,
                            paper_bgcolor="white",
                        )
                        st.plotly_chart(fig_ring, use_container_width=True, key=f"lob_evo_{i}")
                    if i < len(evo_df) - 1:
                        st.markdown("---")
            st.caption(f"Référence: {ref_month.strftime('%m/%Y')} vs {prev_month.strftime('%m/%Y')}")
        else:
            st.info("Aucune Date d’émission valide pour calculer l’évolution CA.")

        # Réalisation budget par branche (objectif annuel)
        st.markdown("### 🎯 Réalisation budget par branche")
        lob_budget_src = df_dash_pd.copy()
        lob_budget_src["_LOB_BUDGET"] = (
            lob_budget_src[col_lob]
            .astype("string")
            .fillna("INCONNU")
            .map(budget_lob_group)
        )

        budget_df = (
            lob_budget_src.groupby("_LOB_BUDGET", dropna=False)
            .agg(**{"CA réalisé (Prime TTC)": ("_TTC_NUM", "sum")})
            .reset_index()
            .rename(columns={"_LOB_BUDGET": "Line of Business"})
        )
        budget_df["Line of Business"] = budget_df["Line of Business"].astype("string").fillna("INCONNU")
        budget_df["Budget annuel"] = budget_df["Line of Business"].map(budget_from_lob).fillna(0.0)
        budget_df["Écart"] = budget_df["CA réalisé (Prime TTC)"] - budget_df["Budget annuel"]
        budget_df["Taux de réalisation"] = np.where(
            budget_df["Budget annuel"] > 0,
            budget_df["CA réalisé (Prime TTC)"] / budget_df["Budget annuel"],
            0.0,
        )
        budget_df["Statut"] = np.select(
            [
                budget_df["Taux de réalisation"] >= 1.0,
                budget_df["Taux de réalisation"] >= 0.7,
            ],
            ["Atteint", "En bonne voie"],
            default="À renforcer",
        )
        budget_df = budget_df.sort_values("Taux de réalisation", ascending=False)

        budget_show = budget_df.copy()
        budget_show["CA réalisé (Prime TTC)"] = budget_show["CA réalisé (Prime TTC)"].round(0).astype(int)
        budget_show["Budget annuel"] = budget_show["Budget annuel"].round(0).astype(int)
        budget_show["Écart"] = budget_show["Écart"].round(0).astype(int)
        budget_show["Taux de réalisation"] = (budget_show["Taux de réalisation"] * 100).round(1).astype(str) + "%"
        st.dataframe(
            budget_show[["Line of Business", "CA réalisé (Prime TTC)", "Budget annuel", "Écart", "Taux de réalisation", "Statut"]],
            use_container_width=True,
            hide_index=True,
        )

        with st.container(border=True):
            st.markdown("#### Nombre de polices par branche")
            if col_emission:
                em_all = df_pl.select(["Branche_rectifiee", "_EMISSION_PARSED"] + ([col_police] if col_police else [])).to_pandas()
                em_all["_EMISSION_PARSED"] = pd.to_datetime(em_all["_EMISSION_PARSED"], errors="coerce")
                em_all = em_all.dropna(subset=["_EMISSION_PARSED"]).copy()
                if not em_all.empty:
                    ref_month = em_all["_EMISSION_PARSED"].max().to_period("M")
                    prev_month = ref_month - 1
                    em_all["YM"] = em_all["_EMISSION_PARSED"].dt.to_period("M")

                    cur_m = em_all[em_all["YM"] == ref_month].copy()
                    prev_m = em_all[em_all["YM"] == prev_month].copy()

                    if col_police and col_police in em_all.columns:
                        cur_g = (
                            cur_m.groupby("Branche_rectifiee", dropna=False)[col_police]
                            .nunique(dropna=True)
                            .rename("Nombre de contrats")
                        )
                        prev_g = (
                            prev_m.groupby("Branche_rectifiee", dropna=False)[col_police]
                            .nunique(dropna=True)
                            .rename("Prev")
                        )
                    else:
                        cur_g = cur_m.groupby("Branche_rectifiee", dropna=False).size().rename("Nombre de contrats")
                        prev_g = prev_m.groupby("Branche_rectifiee", dropna=False).size().rename("Prev")

                    pol_df = pd.concat([cur_g, prev_g], axis=1).fillna(0).reset_index()
                    pol_df["Nombre de contrats"] = pol_df["Nombre de contrats"].astype(int)
                    pol_df["Prev"] = pol_df["Prev"].astype(int)
                    pol_df["Evol"] = np.where(
                        pol_df["Prev"] > 0,
                        (pol_df["Nombre de contrats"] / pol_df["Prev"]) - 1,
                        0.0,
                    )
                    pol_df = pol_df.sort_values("Nombre de contrats", ascending=False)
                    pol_df["Évolution"] = pol_df["Evol"].map(
                        lambda v: f"{'▲' if v > 0 else ('▼' if v < 0 else '•')} {v*100:.1f}%"
                    )
                    pol_view = pol_df.rename(columns={"Branche_rectifiee": "Branche"})[["Branche", "Nombre de contrats", "Évolution"]]
                    st.dataframe(pol_view, use_container_width=True, hide_index=True, height=380)
                    st.caption(f"Période de référence : {ref_month.strftime('%m/%Y')} vs {prev_month.strftime('%m/%Y')}")
                else:
                    st.info("Aucune Date d’émission valide pour calculer l’évolution.")
            else:
                st.info("Colonne Date d’émission introuvable.")

        st.markdown("### 💎 Gros contrats par secteur")
        with st.sidebar:
            seuil_ttc = st.number_input(
                "Seuil gros contrats (Prime TTC)",
                min_value=0.0,
                value=1_000_000.0,
                step=100_000.0,
                format="%.0f",
                key="seuil_gros_ttc",
            )

        total_portefeuille = float(df_dash_pd["_TTC_NUM"].sum()) if df_dash_pd is not None else 0.0
        gros_df = df_dash_pd[df_dash_pd["_TTC_NUM"] >= seuil_ttc].copy()
        if gros_df.empty:
            st.info("Aucun gros contrat avec ce seuil.")
        else:
            gros_agg = (
                gros_df.groupby(["Branche_rectifiee"], dropna=False)
                .agg(
                    **{
                        "Nombre de contrats": ("_TTC_NUM", "size"),
                        "Montant TTC": ("_TTC_NUM", "sum"),
                    }
                )
                .reset_index()
            )
            gros_agg["Part portefeuille"] = np.where(
                total_portefeuille > 0, gros_agg["Montant TTC"] / total_portefeuille, 0
            )
            gros_agg = gros_agg.sort_values("Montant TTC", ascending=False)
            display_gros = gros_agg.copy()
            display_gros["Montant TTC"] = display_gros["Montant TTC"].round(0).astype(int)
            display_gros["Part portefeuille"] = (display_gros["Part portefeuille"] * 100).round(1).astype(str) + "%"
            st.dataframe(display_gros, use_container_width=True, hide_index=True)

            st.markdown("#### Top 10 gros contrats")
            cols_top = [col_police, col_client, "Branche_rectifiee", col_lob, "_TTC_NUM"]
            cols_top = [c for c in cols_top if c]
            top10 = (
                df_dash.select(cols_top)
                .with_columns(pl.col("_TTC_NUM").fill_null(0))
                .filter(pl.col("_TTC_NUM") >= pl.lit(seuil_ttc))
                .sort("_TTC_NUM", descending=True)
                .head(10)
                .to_pandas()
            )
            if "_TTC_NUM" in top10.columns:
                top10 = top10.rename(columns={"_TTC_NUM": "Montant TTC"})
                top10["Montant TTC"] = pd.to_numeric(top10["Montant TTC"], errors="coerce").fillna(0).round(0).astype(int)
            st.dataframe(top10, use_container_width=True, hide_index=True)

    st.subheader("📌 Durée de contrat (jours) : répartition")
    dur_pd = df_dash.select(["_EFFET_PARSED", "_EXP_PARSED", "_TTC_NUM"]).to_pandas()
    dur_pd["_EFFET_PARSED"] = pd.to_datetime(dur_pd["_EFFET_PARSED"], errors="coerce")
    dur_pd["_EXP_PARSED"] = pd.to_datetime(dur_pd["_EXP_PARSED"], errors="coerce")
    dur_pd["DUREE_JOURS"] = (dur_pd["_EXP_PARSED"] - dur_pd["_EFFET_PARSED"]).dt.days
    dur_pd = dur_pd.dropna(subset=["DUREE_JOURS"])
    dur_pd["DUREE_JOURS"] = dur_pd["DUREE_JOURS"].astype(int)
    dur_pd["_TTC_NUM"] = pd.to_numeric(dur_pd["_TTC_NUM"], errors="coerce").fillna(0)

    if not dur_pd.empty:
        agg_dur = (
            dur_pd.groupby("DUREE_JOURS", dropna=False)
            .agg(
                **{
                    "Nombre de contrats": ("DUREE_JOURS", "size"),
                    "Chiffre d'affaires (Prime TTC)": ("_TTC_NUM", "sum"),
                }
            )
            .reset_index()
            .sort_values("DUREE_JOURS")
        )
        total_dur = agg_dur["Nombre de contrats"].sum()
        agg_dur["Taux de repartition"] = np.where(
            total_dur > 0, agg_dur["Nombre de contrats"] / total_dur, 0
        )

        display_dur = agg_dur.copy()
        display_dur["Chiffre d'affaires (Prime TTC)"] = display_dur["Chiffre d'affaires (Prime TTC)"].round(0).astype(int)
        display_dur["Taux de repartition"] = (display_dur["Taux de repartition"] * 100).round(1).astype(str) + "%"
        display_dur = display_dur.rename(columns={"DUREE_JOURS": "Durée (jours)"})

        st.dataframe(display_dur, use_container_width=True, hide_index=True)
        fig_dur = px.bar(
            agg_dur,
            x="DUREE_JOURS",
            y="Nombre de contrats",
            title="Nombre de contrats par durée (jours)",
        )
        st.plotly_chart(fig_dur, use_container_width=True)
        fig_dur_ca = px.bar(
            agg_dur,
            x="DUREE_JOURS",
            y="Chiffre d'affaires (Prime TTC)",
            title="Chiffre d'affaires par durée (jours)",
        )
        st.plotly_chart(fig_dur_ca, use_container_width=True)
    else:
        st.info("Aucune durée de contrat calculable (dates manquantes).")

    show_preview = st.checkbox("Afficher un aperçu (50 lignes)", value=False, key="preview_dash")
    if show_preview:
        st.dataframe(df_dash.head(50).to_pandas(), use_container_width=True)

    # =========================
    # 10) Branche : correction & comptage
    # =========================
    st.subheader("📌 Branche : correction & comptage")

    nb_corrige = df_dash.select(
        (pl.col("Branche_rectifiee") != pl.col("_BRANCHE_NORM")).sum().alias("nb_corrige")
    ).item()

    cA, cB = st.columns(2)
    cA.metric("Lignes corrigées (Branche)", f"{int(nb_corrige):,}")
    cB.metric("Lignes non modifiées", f"{int(summary['total'] - nb_corrige):,}")

    counts = (
        df_dash.group_by("Branche_rectifiee")
           .agg(pl.len().alias("nb_lignes"))
           .sort("nb_lignes", descending=True)
    )
    st.write("**Nombre de lignes par Branche rectifiée :**")
    st.dataframe(counts.to_pandas(), use_container_width=True)

    details = (
        df_dash.group_by(["_BRANCHE_NORM", "Branche_rectifiee"])
           .agg(pl.len().alias("nb_lignes"))
           .sort("nb_lignes", descending=True)
    )
    with st.expander("🔁 Détail des corrections (Branche d'origine → Branche rectifiée)"):
        st.dataframe(details.to_pandas(), use_container_width=True)

def add_excel_like_rownum(df_pl: pl.DataFrame) -> pd.DataFrame:
    return df_pl.with_row_index("ligne").with_columns((pl.col("ligne") + 2).alias("ligne")).to_pandas()

schema_cols = set(df_pl.columns)

# ---- Section : AUTO Durée ----
if nav == "AUTO Durée":
    st.subheader("⚠️ Lignes incohérentes (AUTO & durée ≥ 366 jours)")
    if col_lob:
        st.markdown("### 📌 Segments Line of Business")

        lob_col = pl.col(col_lob).cast(pl.Utf8, strict=False).str.strip_chars().fill_null("INCONNU")
        base_cols = ["Branche_rectifiee", "_TTC_NUM", "Controle (Branche AUTO & Durée contrat)"]
        if col_gros:
            base_cols.append(col_gros)

        df_auto = df_dash.select([lob_col.alias("_LOB")] + base_cols)

        if col_gros:
            gros_flag = (
                pl.col(col_gros)
                .cast(pl.Utf8, strict=False)
                .str.strip_chars()
                .str.to_uppercase()
                .is_in(["OUI", "YES", "Y", "TRUE", "1"])
                | (
                    pl.col(col_gros).is_not_null()
                    & (pl.col(col_gros).cast(pl.Utf8, strict=False).str.strip_chars() != "")
                )
            ).alias("_GROS_FLAG")
            df_auto = df_auto.with_columns(gros_flag)
        else:
            df_auto = df_auto.with_columns(pl.lit(False).alias("_GROS_FLAG"))

        metrics_lob = (
            df_auto.group_by("_LOB")
            .agg([
                pl.len().alias("total_lignes"),
                pl.col("_TTC_NUM").fill_null(0).sum().alias("ca_lob"),
                pl.col("_GROS_FLAG").sum().alias("nb_gros"),
                (pl.col("Controle (Branche AUTO & Durée contrat)") == "Revoir date effet & expiration").sum().alias("nb_incoh"),
            ])
            .sort("total_lignes", descending=True)
        ).to_pandas()

        synth_lob = (
            df_auto.group_by(["_LOB", "Branche_rectifiee"])
            .agg([
                pl.len().alias("Nombre de contrats"),
                pl.col("_TTC_NUM").fill_null(0).sum().alias("Chiffre d'affaires (Prime TTC)"),
                pl.col("_GROS_FLAG").sum().alias("Nb gros capitaux"),
            ])
        ).to_pandas()

        lob_vals = ["Tous"] + sorted(metrics_lob["_LOB"].astype("string").fillna("INCONNU").unique().tolist())
        tabs = st.tabs(lob_vals)

        for tab, lob in zip(tabs, lob_vals):
            with tab:
                if lob == "Tous":
                    m = metrics_lob.sum(numeric_only=True)
                    total_lignes = int(m.get("total_lignes", 0))
                    ca_lob = float(m.get("ca_lob", 0))
                    nb_gros = int(m.get("nb_gros", 0))
                    nb_incoh = int(m.get("nb_incoh", 0))
                    df_lob = df_dash
                    nb_echeance = df_lob.filter(
                        pl.col("_EXP_PARSED").is_not_null() &
                        (pl.col("_EXP_PARSED") <= pl.lit(eval_date_b4))
                    ).height
                    synth = synth_lob.groupby("Branche_rectifiee", dropna=False).sum(numeric_only=True).reset_index()
                else:
                    mrow = metrics_lob[metrics_lob["_LOB"] == lob]
                    total_lignes = int(mrow["total_lignes"].iloc[0]) if not mrow.empty else 0
                    ca_lob = float(mrow["ca_lob"].iloc[0]) if not mrow.empty else 0
                    nb_gros = int(mrow["nb_gros"].iloc[0]) if not mrow.empty else 0
                    nb_incoh = int(mrow["nb_incoh"].iloc[0]) if not mrow.empty else 0
                    df_lob = df_dash.filter(pl.col(col_lob) == lob)
                    nb_echeance = df_lob.filter(
                        pl.col("_EXP_PARSED").is_not_null() &
                        (pl.col("_EXP_PARSED") <= pl.lit(eval_date_b4))
                    ).height
                    synth = synth_lob[synth_lob["_LOB"] == lob].drop(columns=["_LOB"])

                st.markdown(
                    f"""
                    <div class="kpi-grid-3">
                      <div class="kpi-card"><div class="kpi-label">Nombre de contrats</div><div class="kpi-value">{total_lignes:,}</div></div>
                      <div class="kpi-card"><div class="kpi-label">CA (Prime TTC)</div><div class="kpi-value">{ca_lob:,.0f}</div></div>
                      <div class="kpi-card"><div class="kpi-label">Gros capitaux</div><div class="kpi-value">{nb_gros:,}</div></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"""
                    <div class="kpi-feature" style="border-color:#ffffff;background:#ffffff;">
                      <div class="kpi-feature-title">Branche = Automobile</div>
                      <div class="kpi-feature-value">{nb_echeance:,.0f}</div>
                      <div class="kpi-feature-sub">
                        <span class="kpi-feature-delta">Contrats arrivés à échéance (Date expiration &le; Date d'évaluation)</span>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown(f"**Incohérences AUTO (≥ 366 jours)** : {nb_incoh:,}")

                st.markdown("#### 💎 Gros contrats (par branche)")
                seuil_auto = st.number_input(
                    "Seuil gros contrats AUTO (Prime TTC)",
                    min_value=0.0,
                    value=1_000_000.0,
                    step=100_000.0,
                    format="%.0f",
                    key=f"seuil_gros_auto_{lob}",
                )

                total_portefeuille_auto = float(df_dash.select(pl.col("_TTC_NUM").fill_null(0).sum()).item())
                gros_auto = df_lob.select(["Branche_rectifiee", "_TTC_NUM"]).to_pandas()
                gros_auto["_TTC_NUM"] = pd.to_numeric(gros_auto["_TTC_NUM"], errors="coerce").fillna(0)
                gros_auto = gros_auto[gros_auto["_TTC_NUM"] >= seuil_auto]

                if gros_auto.empty:
                    st.info("Aucun gros contrat pour ce seuil.")
                else:
                    gros_auto_agg = (
                        gros_auto.groupby("Branche_rectifiee", dropna=False)
                        .agg(
                            **{
                                "Nombre de contrats": ("_TTC_NUM", "size"),
                                "Montant TTC": ("_TTC_NUM", "sum"),
                            }
                        )
                        .reset_index()
                    )
                    gros_auto_agg["Part portefeuille"] = np.where(
                        total_portefeuille_auto > 0, gros_auto_agg["Montant TTC"] / total_portefeuille_auto, 0
                    )
                    gros_auto_agg = gros_auto_agg.sort_values("Montant TTC", ascending=False)
                    display_auto = gros_auto_agg.copy()
                    display_auto["Montant TTC"] = display_auto["Montant TTC"].round(0).astype(int)
                    display_auto["Part portefeuille"] = (display_auto["Part portefeuille"] * 100).round(1).astype(str) + "%"
                    st.dataframe(display_auto, use_container_width=True, hide_index=True)

                    st.markdown("#### Top 10 gros contrats (AUTO)")
                    cols_top = [col_police, col_client, "Branche_rectifiee", col_lob, "_TTC_NUM"]
                    cols_top = [c for c in cols_top if c]
                    top10_auto = (
                        df_lob.select(cols_top)
                        .with_columns(pl.col("_TTC_NUM").fill_null(0))
                        .filter(pl.col("_TTC_NUM") >= pl.lit(seuil_auto))
                        .sort("_TTC_NUM", descending=True)
                        .head(10)
                        .to_pandas()
                    )
                    if "_TTC_NUM" in top10_auto.columns:
                        top10_auto = top10_auto.rename(columns={"_TTC_NUM": "Montant TTC"})
                        top10_auto["Montant TTC"] = pd.to_numeric(top10_auto["Montant TTC"], errors="coerce").fillna(0).round(0).astype(int)
                    st.dataframe(top10_auto, use_container_width=True, hide_index=True)

                    if st.checkbox(f"Exporter gros contrats {lob}", value=False, key=f"export_gros_{lob}"):
                        st.download_button(
                            f"⬇️ Gros contrats {lob} (Excel)",
                            data=to_excel_bytes(top10_auto, f"GROS_{str(lob)[:24]}"),
                            file_name=f"gros_contrats_{str(lob)}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                synth = synth.sort_values("Nombre de contrats", ascending=False)
                st.dataframe(synth, use_container_width=True, hide_index=True)

                if st.checkbox(f"Préparer export {lob}", value=False, key=f"prep_export_{lob}"):
                    df_lob = df_dash if lob == "Tous" else df_dash.filter(pl.col(col_lob) == lob)
                    st.download_button(
                        f"⬇️ Télécharger {lob} (Excel)",
                        data=to_excel_bytes(df_lob.to_pandas(), f"{str(lob)[:31]}"),
                        file_name=f"lob_{str(lob)}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    else:
        st.info("Colonne Line of Business introuvable : segments indisponibles.")

    incoh_auto_lf = df_dash.filter(pl.col("Controle (Branche AUTO & Durée contrat)") == "Revoir date effet & expiration")
    nb = incoh_auto_lf.height
    st.write(f"Nombre de lignes incohérentes : **{nb:,}**")

    if nb == 0:
        st.success("Aucune incohérence détectée ✅")
    else:
        cols_show = []
        for c in [col_police, col_client, col_inter, col_channel, col_branche, col_effet, col_exp,
                  "Nombre de jour", "Branche_rectifiee", "Controle (Branche AUTO & Durée contrat)"]:
            if c and c in schema_cols and c not in cols_show:
                cols_show.append(c)

        df_bad = incoh_auto_lf.select(cols_show)
        pdf_bad = add_excel_like_rownum(df_bad)
        cols_out = ["ligne"] + [c for c in cols_show if c in pdf_bad.columns]
        if col_effet in pdf_bad.columns:
            pdf_bad[col_effet] = to_ddmmyyyy(pdf_bad[col_effet])
        if col_exp in pdf_bad.columns:
            pdf_bad[col_exp] = to_ddmmyyyy(pdf_bad[col_exp])

        st.dataframe(pdf_bad[cols_out], use_container_width=True, height=520)
        st.download_button(
            "⬇️ Télécharger incohérences AUTO (Excel)",
            data=to_excel_bytes(pdf_bad[cols_out], "AUTO_Duree"),
            file_name="incoherences_AUTO_duree.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---- Section : Dates Effet/Expiration ----
if nav == "Dates Effet/Expiration":
    st.subheader("⚠️ Incohérences : Date d'effet / Date d'expiration")
    incoh_dates_lf = df_pl.filter(pl.col("Controle (Date d'effet / Date d'expiration)") != "Correcte")
    nb = incoh_dates_lf.height
    st.write(f"Nombre de lignes incohérentes : **{nb:,}**")

    if nb == 0:
        st.success("Aucune incohérence détectée ✅")
    else:
        df_bad2 = incoh_dates_lf.select([
            pl.col("_EFFET_RAW").alias("valeur_brute_effet"),
            pl.col("_EXP_RAW").alias("valeur_brute_expiration"),
            pl.col("_EFFET_PARSED").alias("date_effet_parsee"),
            pl.col("_EXP_PARSED").alias("date_exp_parsee"),
            pl.col("Controle (Date d'effet / Date d'expiration)").alias("motif"),
        ])
        pdf_bad2 = add_excel_like_rownum(df_bad2)
        pdf_bad2["date_effet_parsee"] = to_ddmmyyyy(pdf_bad2["valeur_brute_effet"])
        pdf_bad2["date_exp_parsee"] = to_ddmmyyyy(pdf_bad2["valeur_brute_expiration"])
        eff_fmt = pdf_bad2["date_effet_parsee"]
        exp_fmt = pdf_bad2["date_exp_parsee"]
        pdf_bad2["valeur_brute_effet"] = eff_fmt.fillna(pdf_bad2["valeur_brute_effet"].astype("string"))
        pdf_bad2["valeur_brute_expiration"] = exp_fmt.fillna(pdf_bad2["valeur_brute_expiration"].astype("string"))
        cols_out2 = ["ligne", "valeur_brute_effet", "valeur_brute_expiration", "date_effet_parsee", "date_exp_parsee", "motif"]

        st.dataframe(pdf_bad2[cols_out2], use_container_width=True, height=520)
        st.download_button(
            "⬇️ Télécharger anomalies Dates (Excel)",
            data=to_excel_bytes(pdf_bad2[cols_out2], "Dates_Effet_Exp"),
            file_name="anomalies_dates_effet_expiration.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---- Section : Formats ----
if nav == "Formats":
    if fmt_checks:
        bad_cols = [name for name in lf2.columns if name.startswith("_bad_")]
        if bad_cols:
            invalid_mask = None
            for c in bad_cols:
                invalid_mask = pl.col(c) if invalid_mask is None else (invalid_mask | pl.col(c))
            invalid_lf = df_pl.filter(invalid_mask)
            st.subheader("🧪 Formats invalides (dates / montants / taux)")
            if invalid_lf.height == 0:
                st.success("Aucun format invalide détecté.")
            else:
                show_cols = [col_branche, col_effet, col_exp, col_emission, col_prime, col_ttc, col_capitaux, col_accessoires, col_taxes, col_comm_brute, col_taux_comm, col_taux_taxes]
                show_cols = [c for c in show_cols if c]
                st.dataframe(invalid_lf.select(show_cols + bad_cols).to_pandas(), use_container_width=True)

# ---- Section : Exports ----
if nav == "Exports":
    st.subheader("⬇️ Exports")
    st.write("Télécharge le fichier complet avec Branche rectifiée + contrôles.")
    summary_all = compute_summary(df_pl)
    st.write(f"Nombre exact de lignes : **{summary_all['total']:,}**")

    with st.spinner("Préparation du fichier complet..."):
        df_full_pd = df_pl.to_pandas()
        try:
            full_xlsx = to_excel_bytes(df_full_pd, "Production")
        except Exception as e:
            full_xlsx = None
            st.error(f"Erreur export Excel : {e}")

    if full_xlsx is not None:
        st.sidebar.download_button(
            "⬇️ Fichier complet (Excel)",
            data=full_xlsx,
            file_name="production_avec_controles.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Export Excel indisponible, téléchargement CSV proposé.")
        st.sidebar.download_button(
            "⬇️ Fichier complet (CSV)",
            data=to_csv_bytes(df_full_pd),
            file_name="production_avec_controles.csv",
            mime="text/csv",
        )

    st.write("---")
    st.write("Export uniquement des lignes « Revoir date effet & expiration » (AUTO).")

    anom_auto = df_pl.filter(
        pl.col("Controle (Branche AUTO & Durée contrat)") == "Revoir date effet & expiration"
    )

    n1 = anom_auto.height
    if n1 == 0:
        st.info("Aucune anomalie AUTO à exporter.")
    else:
        anom_pd = add_excel_like_rownum(anom_auto)

        cols_anom = ["ligne"]
        for c in [col_police, col_client, col_inter, col_channel, col_branche, col_prime, col_effet, col_exp,
                  "Nombre de jour", "Branche_rectifiee",
                  "Controle (Branche AUTO & Durée contrat)"]:
            if c and c in anom_pd.columns and c not in cols_anom:
                cols_anom.append(c)

        st.dataframe(anom_pd[cols_anom], use_container_width=True, height=520)
        st.sidebar.download_button(
            "⬇️ Anomalies AUTO (Excel)",
            data=to_excel_bytes(anom_pd[cols_anom], "AUTO_Duree"),
            file_name="anomalies_auto_duree.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
