from __future__ import annotations

from io import BytesIO
from datetime import datetime
import json
import os

import numpy as np
import pandas as pd
import streamlit as st

HAS_POLARS = False
pl = None


def _find_col(df: pd.DataFrame, targets: list[str]) -> str | None:
    if df is None or df.empty:
        return None

    def _norm(s: str) -> str:
        import unicodedata
        out = str(s).strip().replace("\u00A0", " ").upper().replace(" ", "_")
        out = "".join(ch for ch in unicodedata.normalize("NFD", out) if unicodedata.category(ch) != "Mn")
        out = "".join(ch for ch in out if ch.isalnum() or ch == "_")
        return out

    map_up = {_norm(c): c for c in df.columns}
    for t in targets:
        key = _norm(t)
        if key in map_up:
            return map_up[key]
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
    return pd.to_numeric(s, errors="coerce").fillna(0)


@st.cache_data(show_spinner=False)
def _safe_read_cached(file_bytes: bytes, filename: str) -> pd.DataFrame:
    if not file_bytes:
        return pd.DataFrame()
    if filename.lower().endswith(".csv"):
        try:
            return pd.read_csv(BytesIO(file_bytes), sep=None, engine="python")
        except Exception:
            return pd.read_csv(BytesIO(file_bytes), sep=";", encoding="utf-8")
    return pd.read_excel(BytesIO(file_bytes))


def _safe_read(uploaded, label: str) -> pd.DataFrame:
    if not uploaded:
        return pd.DataFrame()
    try:
        data = uploaded.getvalue()
        return _safe_read_cached(data, uploaded.name)
    except Exception as exc:
        st.sidebar.error(f"Erreur lecture {label}: {exc}")
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def _build_metrics_cached(df: pd.DataFrame, prec_map: dict[str, float], open_map: dict[str, float]) -> pd.DataFrame:
    return _build_metrics(df, prec_map, open_map)


def _build_metrics(df: pd.DataFrame, prec_map: dict[str, float], open_map: dict[str, float]) -> pd.DataFrame:
    lob_col = _find_col(df, ["LINE_OF_BUSINESS", "LINE OF BUSINESS", "LOB", "BRANCHE"])
    cat_col = _find_col(df, ["CATEGORIE", "CATEGORIE_", "CATEGORIE ", "CATEGORIE_ ", "CATEGORY"])
    policy_col = _find_col(df, [
        "N* POLICE", "N° POLICE", "N POLICE",
        "N* POLICE INTERMEDIAIRE", "N° POLICE INTERMEDIAIRE", "POLICE INTERMEDIAIRE",
        "POLICY", "POLICY_NO", "POLICY NO"
    ])
    prime_nette_col = _find_col(df, ["PRIME_NETTE", "PRIME NETTE"])
    assist_col = _find_col(df, ["ASSISTANCE_PSYCHOLOGIQUE", "ASSISTANCE PSYCHOLOGIQUE"])
    acc_col = _find_col(df, ["ACCESSOIRES", "ACCESSORY", "ACCESSORIES"])

    if not lob_col or not policy_col:
        return pd.DataFrame()

    use_cols = [lob_col, policy_col]
    if cat_col:
        use_cols.append(cat_col)
    if prime_nette_col:
        use_cols.append(prime_nette_col)
    if assist_col:
        use_cols.append(assist_col)
    if acc_col:
        use_cols.append(acc_col)

    if HAS_POLARS:
        df_use = df[use_cols].copy()
        # Cast text columns to string to avoid pyarrow type errors
        for c in [lob_col, policy_col, cat_col]:
            if c and c in df_use.columns:
                df_use[c] = df_use[c].astype("string")
        # Ensure numeric columns are numeric
        for c in [prime_nette_col, assist_col, acc_col]:
            if c and c in df_use.columns:
                df_use[c] = _to_num(df_use[c])
        pl_df = pl.from_pandas(df_use)
        pl_df = pl_df.with_columns([
            pl.col(lob_col).cast(pl.Utf8).str.strip_chars().alias(lob_col),
            pl.col(policy_col).cast(pl.Utf8).str.strip_chars().alias(policy_col),
        ])
        pl_df = pl_df.filter(pl.col(lob_col).is_not_null() & (pl.col(lob_col) != "") & pl.col(policy_col).is_not_null() & (pl.col(policy_col) != ""))

        net_written = pl.lit(0.0)
        if prime_nette_col:
            net_written = net_written + pl.col(prime_nette_col).cast(pl.Float64, strict=False).fill_null(0.0)
        if assist_col:
            net_written = net_written + pl.col(assist_col).cast(pl.Float64, strict=False).fill_null(0.0)

        accessoires = pl.col(acc_col).cast(pl.Float64, strict=False).fill_null(0.0) if acc_col else pl.lit(0.0)

        is_sub = pl.col(lob_col).str.starts_with("└")
        var_cur = pl.col(lob_col).map_elements(lambda v: float(prec_map.get(v, 0.0)), return_dtype=pl.Float64)
        var_open = pl.col(lob_col).map_elements(lambda v: float(open_map.get(v, 0.0)), return_dtype=pl.Float64)
        variation = pl.when(is_sub).then(0.0).otherwise((var_cur - var_open))

        pl_df = pl_df.with_columns([
            net_written.alias("Net Written Premium"),
            accessoires.alias("Accessoires"),
            variation.alias("Variation PREC"),
        ])
        pl_df = pl_df.with_columns([
            (pl.col("Net Written Premium") + pl.col("Accessoires") - pl.col("Variation PREC")).alias("Net Earned Premium")
        ])

        out = (
            pl_df.group_by(lob_col)
            .agg([
                pl.col(policy_col).n_unique().alias("Policy Count"),
                pl.col("Net Written Premium").sum().alias("Gross Written Premium"),
                pl.col("Net Earned Premium").sum().alias("Gross Earned Premium"),
            ])
            .sort(lob_col)
            .to_pandas()
        )
    else:
        dfx = df[use_cols].copy()
        dfx[lob_col] = dfx[lob_col].astype("string").str.strip().replace("", pd.NA)
        # Fusionner 8-TRAVEL dans 2-PROPERTY & OTHERS DAMAGES
        def _merge_lob(v: str) -> str:
            u = str(v).strip().upper()
            if u in {"8-TRAVEL", "8 TRAVEL", "TRAVEL"}:
                return "2-PROPERTY & OTHERS DAMAGES"
            return str(v).strip()
        dfx[lob_col] = dfx[lob_col].map(_merge_lob)
        if policy_col:
            dfx[policy_col] = dfx[policy_col].astype("string").str.strip().replace("", pd.NA)
        dfx = dfx.dropna(subset=[lob_col]).copy()

        dfx["Net Written Premium"] = 0.0
        if prime_nette_col:
            dfx["Net Written Premium"] += _to_num(df.loc[dfx.index, prime_nette_col]) / 1_000_000.0
        if assist_col:
            dfx["Net Written Premium"] += _to_num(df.loc[dfx.index, assist_col]) / 1_000_000.0

        dfx["Accessoires"] = (_to_num(df.loc[dfx.index, acc_col]) / 1_000_000.0) if acc_col else 0.0
        dfx["CA Prime"] = (_to_num(df.loc[dfx.index, prime_nette_col]) / 1_000_000.0 if prime_nette_col else 0.0) + dfx["Accessoires"]

        out = dfx.groupby(lob_col).agg(
            **{
                "Policy Count": (lob_col, "size"),
                "Gross Written Premium": ("Net Written Premium", "sum"),
                "Accessoires Sum": ("Accessoires", "sum"),
                "CA Prime": ("CA Prime", "sum"),
            }
        )
        out.index.name = "Line of Business"
        out = out.reset_index()
        # Variation PREC doit etre appliquee une seule fois par branche (en M FCFA)
        out["Variation PREC"] = out["Line of Business"].map(prec_map).fillna(0.0) - out["Line of Business"].map(open_map).fillna(0.0)
        out["Gross Earned Premium"] = out["Gross Written Premium"] + out["Accessoires Sum"] - out["Variation PREC"]

    if cat_col:
        dfx[cat_col] = dfx[cat_col].astype("string").str.strip().replace("", pd.NA)
        lob_norm = dfx[lob_col].astype("string").str.strip().str.upper()
        is_motor = lob_norm.eq("1-MOTOR") | lob_norm.eq("MOTOR") | lob_norm.eq("1 MOTOR")
        is_health = lob_norm.eq("4-HEALTH") | lob_norm.eq("HEALTH") | lob_norm.eq("4 HEALTH")
        cat_norm = dfx[cat_col].astype("string").str.strip().str.upper()
        motor_allow = ["FLOTTE AUTOMOBILE", "MONO AUTOMOBILE"]
        health_allow = ["GROUPE", "INDIVIDUEL"]
        keep_motor = is_motor & cat_norm.isin(motor_allow)
        keep_health = is_health & cat_norm.isin(health_allow)
        dfx_cat = dfx[(keep_motor | keep_health) & dfx[cat_col].notna()].copy()
        if not dfx_cat.empty:
            keep_motor_cat = dfx_cat[lob_col].astype("string").str.strip().str.upper().isin(["1-MOTOR", "MOTOR", "1 MOTOR"])
            dfx_cat["__parent"] = np.where(keep_motor_cat, "1-MOTOR", "4-HEALTH")
            dfx_cat["Line of Business"] = "└ " + dfx_cat[cat_col].astype("string")
            sub = dfx_cat.groupby(["__parent", "Line of Business"]).agg(
                **{
                    "Policy Count": (lob_col, "size"),
                    "Gross Written Premium": ("Net Written Premium", "sum"),
                    "Accessoires Sum": ("Accessoires", "sum"),
                    "CA Prime": ("CA Prime", "sum"),
                }
            ).reset_index()
            sub["Variation PREC"] = 0.0
            sub["Gross Earned Premium"] = sub["Gross Written Premium"] + sub["Accessoires Sum"]
            sub = sub.drop(columns=["__parent"]).assign(**{"__is_sub": True})
            out = pd.concat([out.assign(**{"__is_sub": False}), sub], ignore_index=True)

    return out


st.set_page_config(page_title="TBR", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&display=swap');
    .tbr-title { font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 34px; margin-bottom: 6px; }
    .tbr-sub { color:#5f6b7a; margin-bottom: 18px; font-size: 14px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='tbr-title'>TBR</div>", unsafe_allow_html=True)
st.markdown("<div class='tbr-sub'>Tableau de bord trimestriel (Q3 vs Q4 2025).</div>", unsafe_allow_html=True)

st.sidebar.header("📂 Fichiers")
q3_file = st.sidebar.file_uploader("Données 2025Q3 (Excel/CSV) — optionnel", type=["xlsx", "xls", "csv"], key="tbr_q3")
q4_file = st.sidebar.file_uploader("Données 2025Q4 (Excel/CSV)", type=["xlsx", "xls", "csv"], key="tbr_q4")

df_q3 = _safe_read(q3_file, "2025Q3")
df_q4 = _safe_read(q4_file, "2025Q4")

if df_q4.empty:
    st.info("Charge le fichier Q4 pour afficher le tableau.")
    st.stop()

lob_col = _find_col(df_q4, ["LINE_OF_BUSINESS", "LINE OF BUSINESS", "LOB", "BRANCHE"]) or _find_col(df_q3, ["LINE_OF_BUSINESS", "LINE OF BUSINESS", "LOB", "BRANCHE"])
lob_list = []
if lob_col and not df_q4.empty:
    lob_list = df_q4[lob_col].astype("string").str.strip().replace("", pd.NA).dropna().unique().tolist()
elif lob_col and not df_q3.empty:
    lob_list = df_q3[lob_col].astype("string").str.strip().replace("", pd.NA).dropna().unique().tolist()

lob_list = [x for x in lob_list if isinstance(x, str)]
def _merge_lob_label(v: str) -> str:
    u = str(v).strip().upper()
    if u in {"8-TRAVEL", "8 TRAVEL", "TRAVEL"}:
        return "2-PROPERTY & OTHERS DAMAGES"
    return str(v).strip()
lob_list = sorted({_merge_lob_label(x) for x in lob_list})

# Ajouter les sous-catégories (uniquement MOTOR/HEALTH) pour les paramètres
cat_col = _find_col(df_q4, ["CATEGORIE", "CATEGORIE_", "CATEGORIE ", "CATEGORIE_ ", "CATEGORY"])
param_rows = list(lob_list)
if cat_col and lob_col and not df_q4.empty:
    dfx = df_q4[[lob_col, cat_col]].copy()
    lob_norm = dfx[lob_col].astype("string").str.strip().str.upper()
    cat_norm = dfx[cat_col].astype("string").str.strip().str.upper()
    motor_allow = ["FLOTTE AUTOMOBILE", "MONO AUTOMOBILE"]
    health_allow = ["GROUPE", "INDIVIDUEL"]
    is_motor = lob_norm.eq("1-MOTOR") | lob_norm.eq("MOTOR") | lob_norm.eq("1 MOTOR")
    is_health = lob_norm.eq("4-HEALTH") | lob_norm.eq("HEALTH") | lob_norm.eq("4 HEALTH")
    sub_motor = sorted(dfx[is_motor & cat_norm.isin(motor_allow)][cat_col].astype("string").str.strip().dropna().unique().tolist())
    sub_health = sorted(dfx[is_health & cat_norm.isin(health_allow)][cat_col].astype("string").str.strip().dropna().unique().tolist())

    def _insert_sub(parent_keys: list[str], subs: list[str], rows: list[str]) -> list[str]:
        if not subs:
            return rows
        for pk in parent_keys:
            if pk in rows:
                idx = rows.index(pk) + 1
                for s in subs:
                    rows.insert(idx, f"└ {s}")
                    idx += 1
                break
        return rows

    param_rows = _insert_sub(["1-MOTOR", "MOTOR", "1 MOTOR"], sub_motor, param_rows)
    param_rows = _insert_sub(["4-HEALTH", "HEALTH", "4 HEALTH"], sub_health, param_rows)

def _load_params(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_params(path: str, data: dict) -> dict:
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        payload = _load_params(path)
        history = payload.get("history", [])
        history.append({"_saved_at": ts, "data": data})
        payload["history"] = history[-50:]
        payload["_saved_at"] = ts
        payload["data"] = data
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return payload
    except Exception:
        return {}


def _as_records(obj) -> list[dict]:
    if obj is None:
        return []
    if isinstance(obj, pd.DataFrame):
        return obj.to_dict(orient="records")
    if isinstance(obj, list):
        if all(isinstance(x, dict) for x in obj):
            return obj
        try:
            return pd.DataFrame(obj).to_dict(orient="records")
        except Exception:
            return []
    if isinstance(obj, dict):
        if "records" in obj and isinstance(obj["records"], list):
            return obj["records"]
        if "data" in obj and isinstance(obj["data"], list):
            return obj["data"]
        try:
            return pd.DataFrame(obj).to_dict(orient="records")
        except Exception:
            return []
    return []


st.sidebar.header("🧮 Paramètres")
st.sidebar.markdown("Saisis les valeurs en **millions de FCFA**.")

params_path = os.path.join(os.getcwd(), "tbr_params.json")
saved = _load_params(params_path)

def _init_param_df(rows: list[str], key: str, col_name: str) -> pd.DataFrame:
    base = pd.DataFrame({"Line of Business": rows, col_name: 0.0})
    saved_data = saved.get("data", saved)
    if key in saved_data:
        saved_df = pd.DataFrame(saved_data[key])
        if not saved_df.empty and "Line of Business" in saved_df.columns and col_name in saved_df.columns:
            base = base.merge(saved_df, on="Line of Business", how="left", suffixes=("", "_saved"))
            base[col_name] = base[f"{col_name}_saved"].fillna(base[col_name])
            base = base.drop(columns=[f"{col_name}_saved"])
    return base

budget_df = _init_param_df(param_rows, "budget", "Budget 2025 (M FCFA)")
prec_df = _init_param_df(param_rows, "prec", "PREC 2025 (M FCFA)")
open_df = _init_param_df(param_rows, "open", "Opening PREC (M FCFA)")

budget_df = st.sidebar.data_editor(budget_df, num_rows="dynamic", use_container_width=True, key="tbr_budget")
prec_df = st.sidebar.data_editor(prec_df, num_rows="dynamic", use_container_width=True, key="tbr_prec")
open_df = st.sidebar.data_editor(open_df, num_rows="dynamic", use_container_width=True, key="tbr_open")

# Variation PREC (affichage)
var_prec_df = pd.DataFrame({
    "Line of Business": param_rows,
})
var_prec_df = var_prec_df.merge(prec_df, on="Line of Business", how="left").merge(open_df, on="Line of Business", how="left")
var_prec_df["Variation PREC (M FCFA)"] = (var_prec_df["PREC 2025 (M FCFA)"].fillna(0) - var_prec_df["Opening PREC (M FCFA)"].fillna(0))
st.sidebar.markdown("#### Variation PREC (calculée)")
st.sidebar.dataframe(
    var_prec_df[["Line of Business", "Variation PREC (M FCFA)"]],
    use_container_width=True,
    hide_index=True,
)

budget_map = dict(zip(budget_df["Line of Business"], budget_df["Budget 2025 (M FCFA)"]))
prec_map = dict(zip(prec_df["Line of Business"], prec_df["PREC 2025 (M FCFA)"]))
open_map = dict(zip(open_df["Line of Business"], open_df["Opening PREC (M FCFA)"]))

q3_manual = None
if df_q3.empty:
    st.sidebar.markdown("#### Saisie manuelle 2025Q3")
    q3_manual = pd.DataFrame({
        "Line of Business": param_rows,
        "Policy Count 2025Q3": 0,
        "GWP 2025Q3 (M FCFA)": 0.0,
        "GEP 2025Q3 (M FCFA)": 0.0,
        "CA 2025Q3 (M FCFA)": 0.0,
    })
    saved_data = saved.get("data", saved)
    if "q3_manual" in saved_data:
        saved_q3 = pd.DataFrame(_as_records(saved_data["q3_manual"]))
        if not saved_q3.empty and "Line of Business" in saved_q3.columns:
            q3_manual = q3_manual.merge(saved_q3, on="Line of Business", how="left", suffixes=("", "_saved"))
            for c in ["Policy Count 2025Q3", "GWP 2025Q3 (M FCFA)", "GEP 2025Q3 (M FCFA)", "CA 2025Q3 (M FCFA)"]:
                if f"{c}_saved" in q3_manual.columns:
                    q3_manual[c] = q3_manual[f"{c}_saved"].fillna(q3_manual[c])
                    q3_manual = q3_manual.drop(columns=[f"{c}_saved"])
    q3_manual = st.sidebar.data_editor(q3_manual, num_rows="dynamic", use_container_width=True, key="tbr_q3_manual")
    metrics_q3 = q3_manual.rename(columns={
        "Policy Count 2025Q3": "Policy Count",
        "GWP 2025Q3 (M FCFA)": "Gross Written Premium",
        "GEP 2025Q3 (M FCFA)": "Gross Earned Premium",
        "CA 2025Q3 (M FCFA)": "CA Prime",
    }).copy()
    if "CA Prime" not in metrics_q3.columns:
        metrics_q3["CA Prime"] = metrics_q3["Gross Written Premium"]
else:
    metrics_q3 = _build_metrics_cached(df_q3, prec_map, open_map)
metrics_q4 = _build_metrics_cached(df_q4, prec_map, open_map) if not df_q4.empty else pd.DataFrame()

def _prep(metrics: pd.DataFrame) -> pd.DataFrame:
    if metrics.empty:
        return pd.DataFrame(columns=["Line of Business", "Policy Count", "Gross Written Premium", "Gross Earned Premium", "Accessoires Sum", "Variation PREC"])
    return metrics.copy()

q3 = _prep(metrics_q3)
q4 = _prep(metrics_q4)

all_lobs = sorted(set(q3["Line of Business"]).union(set(q4["Line of Business"])))
base = pd.DataFrame({"Line of Business": all_lobs})
base = base.merge(q3, on="Line of Business", how="left", suffixes=("", "_Q3")).rename(columns={
    "Policy Count": "Policy Count 2025Q3",
    "Gross Written Premium": "GWP 2025Q3",
    "Gross Earned Premium": "GEP 2025Q3",
    "CA Prime": "CA 2025Q3",
    "Accessoires Sum": "ACC 2025Q3",
    "Variation PREC": "VAR 2025Q3",
})
base = base.merge(q4, on="Line of Business", how="left", suffixes=("", "_Q4")).rename(columns={
    "Policy Count": "Policy Count 2025Q4",
    "Gross Written Premium": "GWP 2025Q4",
    "Gross Earned Premium": "GEP 2025Q4",
    "CA Prime": "CA 2025Q4",
    "Accessoires Sum": "ACC 2025Q4",
    "Variation PREC": "VAR 2025Q4",
})
base = base.fillna(0)
if "ACC 2025Q4" not in base.columns and "Accessoires Sum_Q4" in base.columns:
    base["ACC 2025Q4"] = base["Accessoires Sum_Q4"]
if "VAR 2025Q4" not in base.columns and "Variation PREC_Q4" in base.columns:
    base["VAR 2025Q4"] = base["Variation PREC_Q4"]

base["% Inc Policy Count"] = np.where(
    base["Policy Count 2025Q3"] > 0,
    (base["Policy Count 2025Q4"] - base["Policy Count 2025Q3"]) / base["Policy Count 2025Q3"],
    0.0,
)
base["% Inc GWP"] = np.where(
    base["GWP 2025Q3"] > 0,
    (base["GWP 2025Q4"] - base["GWP 2025Q3"]) / base["GWP 2025Q3"],
    0.0,
)
base["% Inc GEP"] = np.where(
    base["GEP 2025Q3"] > 0,
    (base["GEP 2025Q4"] - base["GEP 2025Q3"]) / base["GEP 2025Q3"],
    0.0,
)

base["Budget 2025 (M FCFA)"] = base["Line of Business"].map(budget_map).fillna(0.0)
base["% Achievement"] = np.where(
    base["Budget 2025 (M FCFA)"] > 0,
    base["CA 2025Q4"] / base["Budget 2025 (M FCFA)"],
    0.0,
)

def _m(v: float) -> float:
    return v

table = pd.DataFrame({
    "Line of Business": base["Line of Business"],
    ("Policy Count", "2025Q3"): base["Policy Count 2025Q3"],
    ("Policy Count", "2025Q4"): base["Policy Count 2025Q4"],
    ("Policy Count", "% Inc"): base["% Inc Policy Count"],
    ("Gross Written Premium (XOF millions)", "2025Q3"): base["GWP 2025Q3"].map(_m),
    ("Gross Written Premium (XOF millions)", "2025Q4"): base["GWP 2025Q4"].map(_m),
    ("Gross Written Premium (XOF millions)", "% Inc"): base["% Inc GWP"],
    ("Forecasted Budget 2025 (XOF millions)", "Budget 2025"): base["Budget 2025 (M FCFA)"],
    ("Forecasted Budget 2025 (XOF millions)", "2025Q4"): base["CA 2025Q4"].map(_m),
    ("Forecasted Budget 2025 (XOF millions)", "% Achievement"): base["% Achievement"],
    ("Gross Earned Premium (XOF millions)", "2025Q3"): base["GEP 2025Q3"].map(_m),
    ("Gross Earned Premium (XOF millions)", "2025Q4"): base["GEP 2025Q4"].map(_m),
    ("Gross Earned Premium (XOF millions)", "% Inc"): base["% Inc GEP"],
})

table.columns = pd.MultiIndex.from_tuples(
    [("Line of Business", "")] + [c for c in table.columns if isinstance(c, tuple)]
)

num_cols = [c for c in table.columns if c != ("Line of Business", "")]
lob_col_key = ("Line of Business", "")
if lob_col_key in table.columns:
    base_mask = ~table[lob_col_key].astype("string").str.startswith("└", na=False)
else:
    base_mask = pd.Series([True] * len(table))

total_row = {c: table.loc[base_mask, c].sum() for c in num_cols}
total_row[lob_col_key] = "TOTAL"

# Recalculer les % Inc pour le TOTAL (Q4−Q3)/Q3
def _recalc_inc(q3_key, q4_key, inc_key):
    q3 = total_row.get(q3_key, 0) or 0
    q4 = total_row.get(q4_key, 0) or 0
    total_row[inc_key] = ((q4 - q3) / q3) if q3 else 0.0

_recalc_inc(("Policy Count", "2025Q3"), ("Policy Count", "2025Q4"), ("Policy Count", "% Inc"))
_recalc_inc(("Gross Written Premium (XOF millions)", "2025Q3"), ("Gross Written Premium (XOF millions)", "2025Q4"), ("Gross Written Premium (XOF millions)", "% Inc"))
_recalc_inc(("Gross Earned Premium (XOF millions)", "2025Q3"), ("Gross Earned Premium (XOF millions)", "2025Q4"), ("Gross Earned Premium (XOF millions)", "% Inc"))
budget_tot = total_row.get(("Forecasted Budget 2025 (XOF millions)", "Budget 2025"), 0) or 0
ca_tot = total_row.get(("Forecasted Budget 2025 (XOF millions)", "2025Q4"), 0) or 0
total_row[("Forecasted Budget 2025 (XOF millions)", "% Achievement")] = (ca_tot / budget_tot) if budget_tot else 0.0

table = pd.concat([table, pd.DataFrame([total_row])], ignore_index=True)

# Ordonner pour afficher les catégories juste sous 1-MOTOR et 4-HEALTH
main_order = {lob: i for i, lob in enumerate(all_lobs)}
line_col = ("Line of Business", "")
if line_col in table.columns:
    lob_series = table[line_col].astype("string")
    def _find_key(keys: list[str]) -> str | None:
        for k in keys:
            if k in main_order:
                return k
        for k in keys:
            if (lob_series.str.upper() == k).any():
                return k
        return None

    motor_key = _find_key(["1-MOTOR", "MOTOR", "1 MOTOR"])
    health_key = _find_key(["4-HEALTH", "HEALTH", "4 HEALTH"])
    motor_order = main_order.get(motor_key, 0)
    health_order = main_order.get(health_key, 0)

    order_vals = lob_series.map(lambda v: main_order.get(str(v), 9999)).astype(float)
    sub_mask = lob_series.str.startswith("└", na=False)
    if sub_mask.any():
        sub_vals = lob_series[sub_mask]
        sub_rank = pd.Series(np.arange(sub_mask.sum()), index=sub_vals.index, dtype=float)
        def _sub_order(label: str, idx: int) -> float:
            lab = str(label).upper()
            if any(k in lab for k in ["HEALTH", "GROUPE", "INDIVIDUEL"]):
                return health_order + 0.1 + (sub_rank.loc[idx] / 1000.0)
            return motor_order + 0.1 + (sub_rank.loc[idx] / 1000.0)
        for idx, val in sub_vals.items():
            order_vals.loc[idx] = _sub_order(val, idx)

    table[("___order", "")] = order_vals
    table = table.sort_values(by=("___order", ""))
    table = table.drop(columns=[("___order", "")])

def _style(df_in: pd.DataFrame) -> pd.io.formats.style.Styler:
    sty = df_in.style
    sty = sty.format({
        ("Policy Count", "2025Q3"): lambda v: f"{int(v):,}".replace(",", " "),
        ("Policy Count", "2025Q4"): lambda v: f"{int(v):,}".replace(",", " "),
        ("Policy Count", "% Inc"): lambda v: f"{float(v)*100:.0f}%",
        ("Gross Written Premium (XOF millions)", "2025Q3"): lambda v: f"{float(v):,.0f}".replace(",", " "),
        ("Gross Written Premium (XOF millions)", "2025Q4"): lambda v: f"{float(v):,.0f}".replace(",", " "),
        ("Gross Written Premium (XOF millions)", "% Inc"): lambda v: f"{float(v)*100:.0f}%",
        ("Forecasted Budget 2025 (XOF millions)", "Budget 2025"): lambda v: f"{float(v):,.0f}".replace(",", " "),
        ("Forecasted Budget 2025 (XOF millions)", "2025Q4"): lambda v: f"{float(v):,.0f}".replace(",", " "),
        ("Forecasted Budget 2025 (XOF millions)", "% Achievement"): lambda v: f"{float(v)*100:.0f}%",
        ("Gross Earned Premium (XOF millions)", "2025Q3"): lambda v: f"{float(v):,.0f}".replace(",", " "),
        ("Gross Earned Premium (XOF millions)", "2025Q4"): lambda v: f"{float(v):,.0f}".replace(",", " "),
        ("Gross Earned Premium (XOF millions)", "% Inc"): lambda v: f"{float(v)*100:.0f}%",
    })
    sty = sty.set_table_styles([
        {"selector": "th", "props": [("background-color", "#6f6f6f"), ("color", "white"), ("font-weight", "bold"), ("text-align", "center")]},
        {"selector": "th.col_heading.level0", "props": [("border", "1px solid #333")]},
        {"selector": "td", "props": [("text-align", "right"), ("border", "1px solid #c9c9c9")]},
        {"selector": "td.col0", "props": [("text-align", "left"), ("font-weight", "bold")]},
        {"selector": "td.col3, td.col6, td.col9, td.col12", "props": [("color", "black"), ("font-weight", "bold")]},
    ])
    sub_mask = df_in["Line of Business"].astype("string").str.startswith("└", na=False)
    sty = sty.set_properties(
        subset=pd.IndexSlice[sub_mask, :],
        **{"font-style": "italic", "font-size": "12px"}
    )
    inc_cols = [c for c in df_in.columns if isinstance(c, tuple) and (c[1] == "% Inc" or c[1] == "% Achievement")]
    if inc_cols:
        sty = sty.set_properties(
            subset=pd.IndexSlice[:, inc_cols],
            **{"font-weight": "bold", "color": "black"}
        )
    total_idx = df_in.index[df_in[("Line of Business", "")].astype("string") == "TOTAL"]
    if len(total_idx) > 0:
        sty = sty.set_properties(
            subset=pd.IndexSlice[total_idx, :],
            **{"font-weight": "bold", "color": "black"}
        )
    return sty

st.markdown("### Tableau TBR")
st.dataframe(_style(table), use_container_width=True, hide_index=True)

# Deuxième tableau de cotation (saisie manuelle Q3/Q4)
st.sidebar.markdown("### Cotations — Saisie manuelle")
quote_rows = [r for r in param_rows if not str(r).startswith("└")]
quote_df = pd.DataFrame({
    "Line of Business": quote_rows,
    "2025Q3": 0.0,
    "2025Q4": 0.0,
})
saved_data = saved.get("data", saved)
if "quote_table" in saved_data:
    saved_q = pd.DataFrame(_as_records(saved_data["quote_table"]))
    if not saved_q.empty and "Line of Business" in saved_q.columns:
        quote_df = quote_df.merge(saved_q, on="Line of Business", how="left", suffixes=("", "_saved"))
        for c in ["2025Q3", "2025Q4"]:
            if f"{c}_saved" in quote_df.columns:
                quote_df[c] = quote_df[f"{c}_saved"].fillna(quote_df[c])
                quote_df = quote_df.drop(columns=[f"{c}_saved"])

quote_df = st.sidebar.data_editor(quote_df, num_rows="dynamic", use_container_width=True, key="tbr_quote_table")
quote_df["% inc"] = np.where(
    quote_df["2025Q3"] > 0,
    (quote_df["2025Q4"] - quote_df["2025Q3"]) / quote_df["2025Q3"],
    0.0,
)
quote_total = {
    "Line of Business": "TOTAL",
    "2025Q3": quote_df["2025Q3"].sum(),
    "2025Q4": quote_df["2025Q4"].sum(),
}
quote_total["% inc"] = ( (quote_total["2025Q4"] - quote_total["2025Q3"]) / quote_total["2025Q3"] ) if quote_total["2025Q3"] else 0.0
quote_out = pd.concat([quote_df, pd.DataFrame([quote_total])], ignore_index=True)

st.markdown("### Tableau Cotations (rapport)")
quote_sty = quote_out.style.format({
    "2025Q3": lambda v: f"{float(v):,.0f}".replace(",", " "),
    "2025Q4": lambda v: f"{float(v):,.0f}".replace(",", " "),
    "% inc": lambda v: f"{float(v)*100:.0f}%",
})
quote_sty = quote_sty.set_properties(subset=["% inc"], **{"font-weight": "bold", "color": "black"})
q_total_idx = quote_out.index[quote_out["Line of Business"].astype("string") == "TOTAL"]
if len(q_total_idx) > 0:
    quote_sty = quote_sty.set_properties(
        subset=pd.IndexSlice[q_total_idx, :],
        **{"font-weight": "bold", "color": "black"}
    )
st.dataframe(
    quote_sty,
    use_container_width=True,
    hide_index=True,
)

# Troisième tableau : Claims
st.sidebar.markdown("### Claims — Saisie manuelle")
claims_rows = [r for r in param_rows if not str(r).startswith("└")]
claims_df = pd.DataFrame({
    "Line of Business": claims_rows,
    "Claims count 2025Q3": 0.0,
    "Claims count 2025Q4": 0.0,
    "Claims costs 2025Q3 (M FCFA)": 0.0,
    "Claims costs 2025Q4 (M FCFA)": 0.0,
})
saved_data = saved.get("data", saved)
if "claims_table" in saved_data:
    saved_c = pd.DataFrame(_as_records(saved_data["claims_table"]))
    if not saved_c.empty and "Line of Business" in saved_c.columns:
        claims_df = claims_df.merge(saved_c, on="Line of Business", how="left", suffixes=("", "_saved"))
        for c in [
            "Claims count 2025Q3", "Claims count 2025Q4",
            "Claims costs 2025Q3 (M FCFA)", "Claims costs 2025Q4 (M FCFA)",
        ]:
            if f"{c}_saved" in claims_df.columns:
                claims_df[c] = claims_df[f"{c}_saved"].fillna(claims_df[c])
                claims_df = claims_df.drop(columns=[f"{c}_saved"])

claims_df = st.sidebar.data_editor(claims_df, num_rows="dynamic", use_container_width=True, key="tbr_claims_table")
claims_df["% inc count"] = np.where(
    claims_df["Claims count 2025Q3"] > 0,
    (claims_df["Claims count 2025Q4"] - claims_df["Claims count 2025Q3"]) / claims_df["Claims count 2025Q3"],
    0.0,
)
claims_df["% inc costs"] = np.where(
    claims_df["Claims costs 2025Q3 (M FCFA)"] > 0,
    (claims_df["Claims costs 2025Q4 (M FCFA)"] - claims_df["Claims costs 2025Q3 (M FCFA)"]) / claims_df["Claims costs 2025Q3 (M FCFA)"],
    0.0,
)

claims_total = {
    "Line of Business": "TOTAL",
    "Claims count 2025Q3": claims_df["Claims count 2025Q3"].sum(),
    "Claims count 2025Q4": claims_df["Claims count 2025Q4"].sum(),
    "Claims costs 2025Q3 (M FCFA)": claims_df["Claims costs 2025Q3 (M FCFA)"].sum(),
    "Claims costs 2025Q4 (M FCFA)": claims_df["Claims costs 2025Q4 (M FCFA)"].sum(),
}
claims_total["% inc count"] = ((claims_total["Claims count 2025Q4"] - claims_total["Claims count 2025Q3"]) / claims_total["Claims count 2025Q3"]) if claims_total["Claims count 2025Q3"] else 0.0
claims_total["% inc costs"] = ((claims_total["Claims costs 2025Q4 (M FCFA)"] - claims_total["Claims costs 2025Q3 (M FCFA)"]) / claims_total["Claims costs 2025Q3 (M FCFA)"]) if claims_total["Claims costs 2025Q3 (M FCFA)"] else 0.0

claims_out = pd.concat([claims_df, pd.DataFrame([claims_total])], ignore_index=True)

claims_sty = claims_out.style.format({
    "Claims count 2025Q3": lambda v: f"{float(v):,.0f}".replace(",", " "),
    "Claims count 2025Q4": lambda v: f"{float(v):,.0f}".replace(",", " "),
    "Claims costs 2025Q3 (M FCFA)": lambda v: f"{float(v):,.0f}".replace(",", " "),
    "Claims costs 2025Q4 (M FCFA)": lambda v: f"{float(v):,.0f}".replace(",", " "),
    "% inc count": lambda v: f"{float(v)*100:.0f}%",
    "% inc costs": lambda v: f"{float(v)*100:.0f}%",
})
claims_sty = claims_sty.set_table_styles([
    {"selector": "th", "props": [("background-color", "#6f6f6f"), ("color", "white"), ("font-weight", "bold"), ("text-align", "center")]},
    {"selector": "td", "props": [("text-align", "right"), ("border", "1px solid #c9c9c9")]},
    {"selector": "td.col0", "props": [("text-align", "left"), ("font-weight", "bold")]},
])
claims_sty = claims_sty.set_properties(subset=["% inc count", "% inc costs"], **{"color": "black"})

st.markdown("### Tableau Claims (rapport)")
st.dataframe(
    claims_sty,
    use_container_width=True,
    hide_index=True,
)

# Sauvegarde globale (en bas de la sidebar) pour inclure Q3 + Cotations + Claims
st.sidebar.markdown("### 💾 Sauvegarde")
auto_save = st.sidebar.checkbox("Sauvegarde auto", value=True, key="tbr_auto_save")
if auto_save:
    payload = {
        "budget": budget_df.to_dict(orient="records"),
        "prec": prec_df.to_dict(orient="records"),
        "open": open_df.to_dict(orient="records"),
        "quote_table": _as_records(quote_df),
        "claims_table": _as_records(claims_df),
    }
    if q3_manual is not None:
        payload["q3_manual"] = _as_records(q3_manual)
    _save_params(params_path, payload)

save_cols = st.sidebar.columns(2)
with save_cols[0]:
    if st.sidebar.button("💾 Sauvegarder", use_container_width=True, key="tbr_save_main"):
        payload = {
            "budget": budget_df.to_dict(orient="records"),
            "prec": prec_df.to_dict(orient="records"),
            "open": open_df.to_dict(orient="records"),
            "quote_table": _as_records(quote_df),
            "claims_table": _as_records(claims_df),
        }
        if q3_manual is not None:
            payload["q3_manual"] = _as_records(q3_manual)
        _save_params(params_path, payload)
        st.sidebar.success("Paramètres sauvegardés.")
with save_cols[1]:
    if st.sidebar.button("🔄 Recharger", use_container_width=True, key="tbr_reload_main"):
        for k in ["tbr_budget", "tbr_prec", "tbr_open", "tbr_q3_manual", "tbr_quote_table", "tbr_claims_table"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

saved_now = _load_params(params_path)
if saved_now.get("_saved_at"):
    st.sidebar.caption(f"Paramètres sauvegardés : {saved_now.get('_saved_at')}")

history = saved_now.get("history", [])
if history:
    st.sidebar.markdown("#### Historique des sauvegardes")
    options = [h.get("_saved_at", "") for h in history if h.get("_saved_at")]
    sel = st.sidebar.selectbox("Choisir une sauvegarde", options[::-1], index=0, key="tbr_history_sel_main")
    if sel:
        snap = next((h for h in history if h.get("_saved_at") == sel), None)
        if snap:
            if st.sidebar.button("↩️ Restaurer cette sauvegarde", use_container_width=True, key="tbr_restore_main"):
                _save_params(params_path, snap.get("data", {}))
                for k in ["tbr_budget", "tbr_prec", "tbr_open", "tbr_q3_manual", "tbr_quote_table", "tbr_claims_table"]:
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()
