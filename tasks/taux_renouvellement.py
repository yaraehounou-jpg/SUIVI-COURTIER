from __future__ import annotations

from datetime import datetime, date
from io import BytesIO
import calendar

import numpy as np
import pandas as pd
import streamlit as st

from utils import fmt_int, add_segment

# --- Polars (optionnel) ---
try:
    import polars as pl
    HAS_POLARS = True
except Exception:
    pl = None
    HAS_POLARS = False


# =========================
# Utils: colonnes / formats
# =========================
def _fmt_pct(x) -> str:
    try:
        return f"{float(x) * 100:.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"


def _fmt_space_int(x) -> str:
    return fmt_int(x)


def _pl_bool_col(col_name: str):
    """
    Expression Polars pour convertir en booléen aligné avec _to_bool_series.
    True si valeur ∈ {"TRUE","1","OUI","YES","Y","VRAI"}, False si valeur ∈ {"FALSE","0","NON","NO","N","FAUX"}.
    """
    true_set = ["TRUE", "1", "OUI", "YES", "Y", "VRAI"]
    false_set = ["FALSE", "0", "NON", "NO", "N", "FAUX"]
    return (
        pl.col(col_name)
        .cast(pl.Utf8, strict=False)
        .str.strip_chars()
        .str.to_uppercase()
        .map_elements(lambda v: True if v in true_set else False if v in false_set else False, return_dtype=pl.Boolean)
        .fill_null(False)
    )


def _month_name_fr(m: int) -> str:
    mois_fr = {
        1: "JANVIER", 2: "FEVRIER", 3: "MARS", 4: "AVRIL",
        5: "MAI", 6: "JUIN", 7: "JUILLET", 8: "AOUT",
        9: "SEPTEMBRE", 10: "OCTOBRE", 11: "NOVEMBRE", 12: "DECEMBRE"
    }
    return mois_fr.get(m, str(m))


def _segment_suffix(segment: str) -> str:
    if segment == "Agent":
        return "AGENT"
    if segment == "Courtier":
        return "COURTIER"
    if segment == "Direct":
        return "DIRECT"
    return "TOUS"


def _month_bounds(y: int, m: int) -> tuple[date, date]:
    last_day = calendar.monthrange(y, m)[1]
    return date(y, m, 1), date(y, m, last_day)


def _find_col(df: pd.DataFrame, targets: list[str]) -> str | None:
    """
    Trouve une colonne dans df en comparant en UPPER + strip + sans accents.
    """
    if df is None or df.empty:
        return None
    def _norm(s: str) -> str:
        import unicodedata
        s = str(s).strip().replace("\u00A0", " ").upper()
        s = s.replace(" ", "_")
        s = s.replace("__", "_")
        s = s.strip("_")
        return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    map_up = {_norm(c): c for c in df.columns}
    for t in targets:
        key = _norm(t)
        if key in map_up:
            return map_up[key]
    return None


def _to_bool_series(s: pd.Series) -> pd.Series:
    """
    Convertit une série en bool robuste.
    """
    if s is None:
        return pd.Series([False])
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0).astype(float).isin([1.0])
    if s.dtype == bool:
        return s.fillna(False)
    ss = (
        s.astype("string")
        .fillna("")
        .str.replace("\u00A0", " ", regex=False)
        .str.strip()
        .str.upper()
    )
    true_set = {"TRUE", "1", "1.0", "OUI", "YES", "Y", "VRAI", "RENOUVELLEMENT", "RENOUVELER", "RENOUVELLER", "RENOUVELLE"}
    false_set = {"FALSE", "0", "0.0", "NON", "NO", "N", "FAUX"}
    return ss.map(lambda v: True if v in true_set else False if v in false_set else False)


def _mask_by_year_month(series: pd.Series, year_sel: object, month_sel: object) -> pd.Series:
    s = pd.to_datetime(series, errors="coerce")
    if year_sel == "Tous" and month_sel == "Tous":
        return s.notna()
    if year_sel != "Tous" and month_sel == "Tous":
        return s.dt.year == int(year_sel)
    if year_sel == "Tous" and month_sel != "Tous":
        return s.dt.month == int(month_sel)
    return (s.dt.year == int(year_sel)) & (s.dt.month == int(month_sel))


def _calc_age_mean(df: pd.DataFrame, ref_date: date) -> float | None:
    mise_col = _find_col(df, ["DATE_1ERE_MISE_EN_CIRCULATION", "DATE 1ERE MISE EN CIRCULATION", "DATE_MISE_EN_CIRCULATION"])
    if not mise_col:
        return None
    dfx = df.copy()
    dfx[mise_col] = pd.to_datetime(dfx[mise_col], errors="coerce")
    num = pd.to_numeric(dfx[mise_col], errors="coerce")
    mask_num = dfx[mise_col].isna() & num.notna()
    if mask_num.any():
        dfx.loc[mask_num, mise_col] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask_num], unit="D")
    dfx = dfx.dropna(subset=[mise_col]).copy()
    if dfx.empty:
        return None
    years = (pd.Timestamp(ref_date) - dfx[mise_col]).dt.days / 365.25
    if years.empty:
        return None
    return float(np.floor(years.mean()))


def _renewals_by_create(
    df: pd.DataFrame,
    policy_col: str,
    est_col: str,
    create_col: str,
    year_sel: object,
    month_sel: object,
) -> int:
    dfx = df.copy()
    dfx[policy_col] = dfx[policy_col].astype("string").str.strip().replace("", pd.NA)
    dfx[est_col] = _to_bool_series(dfx[est_col]).astype(bool)
    dfx[create_col] = pd.to_datetime(dfx[create_col], errors="coerce")
    dfx = dfx.dropna(subset=[policy_col, create_col]).copy()
    if dfx.empty:
        return 0
    mask_create = _mask_by_year_month(dfx[create_col], year_sel, month_sel)
    dfx = dfx.loc[mask_create & (dfx[est_col] == True), [policy_col]]
    return int(dfx[policy_col].nunique(dropna=True))


def _style_table(df: pd.DataFrame):
    if df is None or df.empty:
        return df

    def _row_style(row):
        is_total = str(row.get("PRODUIT", "")).strip().lower() in {"grand total", "total"}
        if is_total:
            return ["background-color:#1F4E79;color:white;font-weight:bold;"] * len(row)
        return [""] * len(row)

    return (
        df.style
        .format({
            "Total": lambda v: _fmt_space_int(v),
            "# New Business": lambda v: _fmt_space_int(v),
            "# Renouvellement": lambda v: _fmt_space_int(v),
            "% Renouvellement": lambda v: _fmt_pct(v),
        })
        .apply(_row_style, axis=1)
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
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "left"),
                ("font-weight", "bold"),
            ]},
        ])
    )


def _style_recap_prod(df: pd.DataFrame):
    if df is None or df.empty:
        return df
    header_bg = "#F9C74F"

    def _row_style(row):
        if str(row.get("PRODUIT", "")).strip().lower() == "grand total":
            return ["background-color:#3B7C2E;color:white;font-weight:bold;"] * len(row)
        return ["" for _ in row]

    return (
        df.style
        .format({
            "Nombre de contrats arrivant à échéance": _fmt_space_int,
            "Ceux qui ont effectivement renouvelé": _fmt_space_int,
            "Contrats non attendus (renouvellements imprévus)": _fmt_space_int,
            "% Renouvellement Pur": _fmt_pct,
        })
        .apply(_row_style, axis=1)
        .set_table_styles([
            {"selector": "th", "props": [
                ("background-color", header_bg),
                ("color", "black"),
                ("font-weight", "bold"),
                ("text-align", "center"),
                ("border", "1px solid #d0d0d0"),
            ]},
            {"selector": "tbody td", "props": [
                ("text-align", "center"),
                ("border", "1px solid #d0d0d0"),
                ("padding", "8px"),
                ("font-size", "13px"),
            ]},
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "left"),
                ("font-weight", "bold"),
            ]},
        ])
    )


def _style_age_table(df: pd.DataFrame):
    if df is None or df.empty:
        return df
    header_bg = "#7E3AAE"

    def _row_style(row):
        label = row.get("PRODUIT", row.get("BRANCHE", ""))
        if str(label).strip().lower() == "grand total":
            return ["background-color:#1F1F1F;color:white;font-weight:bold;"] * len(row)
        return ["" for _ in row]

    return (
        df.style
        .format({
            "Age Moyen": _fmt_space_int,
            "Age Minimum": _fmt_space_int,
            "Age Maximum": _fmt_space_int,
            "Nombre cumulés de plaques uniques": _fmt_space_int,
        })
        .apply(_row_style, axis=1)
        .set_table_styles([
            {"selector": "th", "props": [
                ("background-color", header_bg),
                ("color", "white"),
                ("font-weight", "bold"),
                ("text-align", "center"),
                ("border", "1px solid #5b2b7e"),
            ]},
            {"selector": "tbody td", "props": [
                ("text-align", "center"),
                ("border", "1px solid #d0d0d0"),
                ("padding", "8px"),
                ("font-size", "13px"),
            ]},
            {"selector": "tbody td:first-child", "props": [
                ("text-align", "left"),
                ("font-weight", "bold"),
            ]},
        ])
    )


def _vehicule_age_table(
    df_period: pd.DataFrame,
    prod_col: str,
    immat_col: str,
    mise_col: str,
    ref_date: date,
) -> pd.DataFrame:
    """
    Table âge véhicule par produit (basé sur DATE 1ere MISE EN CIRCULATION).
    """
    dfx = df_period[[prod_col, immat_col, mise_col]].copy()
    dfx[immat_col] = dfx[immat_col].astype("string").str.strip().replace("", pd.NA)
    dfx[mise_col] = pd.to_datetime(dfx[mise_col], errors="coerce")
    num = pd.to_numeric(dfx[mise_col], errors="coerce")
    mask_num = dfx[mise_col].isna() & num.notna()
    if mask_num.any():
        dfx.loc[mask_num, mise_col] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask_num], unit="D")
    dfx = dfx.dropna(subset=[immat_col, mise_col]).copy()
    if dfx.empty:
        return pd.DataFrame(columns=[
            "PRODUIT",
            "Age Moyen",
            "Age Minimum",
            "Age Maximum",
            "Nombre cumulés de plaques uniques",
        ])

    # 1 ligne par immatriculation
    dfx = dfx.sort_values(mise_col)
    dfx = dfx.drop_duplicates(subset=[immat_col], keep="first")

    ref_ts = pd.Timestamp(ref_date)
    ages = ((ref_ts - dfx[mise_col]).dt.days / 365.25).clip(lower=0)
    dfx["_AGE"] = ages.fillna(0)

    grp = dfx.groupby(prod_col, dropna=False)
    out = grp["_AGE"].agg(["mean", "min", "max", "count"]).reset_index()
    out.columns = [prod_col, "Age Moyen", "Age Minimum", "Age Maximum", "Nombre cumulés de plaques uniques"]

    out["Age Moyen"] = out["Age Moyen"].round(0).astype(int)
    out["Age Minimum"] = out["Age Minimum"].round(0).astype(int)
    out["Age Maximum"] = out["Age Maximum"].round(0).astype(int)
    out["Nombre cumulés de plaques uniques"] = out["Nombre cumulés de plaques uniques"].astype(int)

    out = out.rename(columns={prod_col: "PRODUIT"})
    out = out.sort_values("Nombre cumulés de plaques uniques", ascending=False).reset_index(drop=True)

    total_row = pd.DataFrame([{
        "PRODUIT": "Grand Total",
        "Age Moyen": int(round(dfx["_AGE"].mean() or 0)),
        "Age Minimum": int(round(dfx["_AGE"].min() or 0)),
        "Age Maximum": int(round(dfx["_AGE"].max() or 0)),
        "Nombre cumulés de plaques uniques": int(dfx[immat_col].nunique(dropna=True)),
    }])
    return pd.concat([out, total_row], ignore_index=True)


def _inject_kpi_css() -> None:
    """Petite couche de style pour les KPI locaux (sans dépendre du CSS global)."""
    st.markdown(
        """
        <style>
        .kpi_card {
            background: linear-gradient(135deg, #16324F, #1E4E79);
            color: #ffffff;
            border-radius: 14px;
            padding: 14px 16px;
            text-align: center;
            box-shadow: 0 8px 24px rgba(0,0,0,0.12);
            border: 1px solid rgba(255,255,255,0.08);
        }
        .kpi_card .kpi_value {
            font-size: 28px;
            font-weight: 800;
            letter-spacing: 0.3px;
            color: #ffffff;
        }
        .kpi_card .kpi_label {
            margin-top: 4px;
            font-size: 13px;
            font-weight: 500;
            opacity: 0.92;
            color: #ffffff;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================
# Calculs (Polars / Pandas)
# =========================
def _compute_pandas(df: pd.DataFrame, policy_col: str, start_dt: pd.Timestamp, end_dt: pd.Timestamp) -> pd.DataFrame:
    """
    Table par PRODUIT:
    Total / New Business (EST_RENOUVELLER=False) / Renouvellement (EST_RENOUVELLER=True) en nb unique policyNo.
    """
    df2 = df.copy()

    # produit
    prod_col = _find_col(df2, ["PRODUIT", "PRODUCT"])
    if prod_col is None:
        df2["PRODUIT"] = "INCONNU"
        prod_col = "PRODUIT"
    df2[prod_col] = df2[prod_col].astype("string").fillna("INCONNU").str.strip()

    # filtrage période
    df2 = df2[(df2["DATE_CREATE"] >= start_dt) & (df2["DATE_CREATE"] <= end_dt)].copy()
    if df2.empty:
        return pd.DataFrame(columns=["PRODUIT", "Total", "# New Business", "# Renouvellement", "% Renouvellement"])

    # policy propre
    df2[policy_col] = df2[policy_col].astype("string").fillna("").str.strip()
    df2.loc[df2[policy_col] == "", policy_col] = pd.NA

    # bool renouv
    est_col = _find_col(df2, ["EST_RENOUVELLER", "EST_RENOUVELER"])
    if est_col is None:
        df2["EST_RENOUVELLER"] = False
        est_col = "EST_RENOUVELLER"
    df2[est_col] = _to_bool_series(df2[est_col]).astype(bool)

    g_total = df2.groupby(prod_col)[policy_col].nunique(dropna=True).rename("Total").reset_index()
    mask_new = (~df2[est_col])
    mask_ren = df2[est_col]
    g_new = df2[mask_new].groupby(prod_col)[policy_col].nunique(dropna=True).rename("# New Business").reset_index()
    g_ren = df2[mask_ren].groupby(prod_col)[policy_col].nunique(dropna=True).rename("# Renouvellement").reset_index()

    out = g_total.merge(g_new, on=prod_col, how="left").merge(g_ren, on=prod_col, how="left")
    out = out.rename(columns={prod_col: "PRODUIT"})
    out["# New Business"] = out["# New Business"].fillna(0).astype(int)
    out["# Renouvellement"] = out["# Renouvellement"].fillna(0).astype(int)
    out["Total"] = out["Total"].fillna(0).astype(int)
    out["% Renouvellement"] = np.where(out["Total"] > 0, out["# Renouvellement"] / out["Total"], 0.0)

    out = out.sort_values("Total", ascending=False).reset_index(drop=True)

    # Grand total
    total_all = int(df2[policy_col].nunique(dropna=True))
    new_all = int(df2.loc[mask_new, policy_col].nunique(dropna=True))
    ren_all = int(df2.loc[mask_ren, policy_col].nunique(dropna=True))
    pct_all = (ren_all / total_all) if total_all else 0.0

    grand = pd.DataFrame([{
        "PRODUIT": "Grand Total",
        "Total": total_all,
        "# New Business": new_all,
        "# Renouvellement": ren_all,
        "% Renouvellement": pct_all,
    }])

    return pd.concat([out, grand], ignore_index=True)


def _compute_polars(df: pd.DataFrame, policy_col: str, start_dt: datetime, end_dt: datetime) -> pd.DataFrame:
    """
    Même calcul en Polars.
    """
    cols_opt = [c for c in ["PRODUIT", "SEGMENT", "EST_RENOUVELLER"] if c in df.columns]
    dfp = df[["DATE_CREATE", policy_col] + cols_opt].copy()
    pl_df = pl.from_pandas(dfp)

    # colonnes utiles
    prod_col = "PRODUIT" if "PRODUIT" in pl_df.columns else None
    if prod_col is None:
        pl_df = pl_df.with_columns(pl.lit("INCONNU").alias("PRODUIT"))
        prod_col = "PRODUIT"

    est_col = "EST_RENOUVELLER" if "EST_RENOUVELLER" in pl_df.columns else None
    if est_col is None:
        pl_df = pl_df.with_columns(pl.lit(False).alias("EST_RENOUVELLER"))
        est_col = "EST_RENOUVELLER"
    pl_df = pl_df.with_columns([
        pl.col("DATE_CREATE").cast(pl.Datetime, strict=False),
        pl.col(prod_col).cast(pl.Utf8, strict=False).fill_null("INCONNU").str.strip_chars(),
        pl.col(policy_col).cast(pl.Utf8, strict=False).fill_null("").str.strip_chars(),
    ]).with_columns([
        pl.when(pl.col(policy_col) == "").then(None).otherwise(pl.col(policy_col)).alias(policy_col),
        _pl_bool_col(est_col).alias(est_col),
    ])

    pl_m = pl_df.filter((pl.col("DATE_CREATE") >= pl.lit(start_dt)) & (pl.col("DATE_CREATE") <= pl.lit(end_dt)))
    if pl_m.height == 0:
        return pd.DataFrame(columns=["PRODUIT", "Total", "# New Business", "# Renouvellement", "% Renouvellement"])

    out = (
        pl_m.group_by(prod_col)
        .agg([
            pl.col(policy_col).n_unique().alias("Total"),
            pl.col(policy_col).filter(pl.col(est_col) == False).n_unique().alias("# New Business"),
            pl.col(policy_col).filter(pl.col(est_col) == True).n_unique().alias("# Renouvellement"),
        ])
        .with_columns([
            pl.when(pl.col("Total") > 0).then(pl.col("# Renouvellement") / pl.col("Total")).otherwise(0.0).alias("% Renouvellement")
        ])
        .sort("Total", descending=True)
    )

    out_pd = out.to_pandas().rename(columns={prod_col: "PRODUIT"})

    # grand total
    total_all = pl_m.select(pl.col(policy_col).n_unique()).item()
    new_all = pl_m.filter(pl.col(est_col) == False).select(pl.col(policy_col).n_unique()).item()
    ren_all = pl_m.filter(pl.col(est_col) == True).select(pl.col(policy_col).n_unique()).item()
    pct_all = (ren_all / total_all) if total_all else 0.0

    grand = pd.DataFrame([{
        "PRODUIT": "Grand Total",
        "Total": int(total_all or 0),
        "# New Business": int(new_all or 0),
        "# Renouvellement": int(ren_all or 0),
        "% Renouvellement": float(pct_all or 0.0),
    }])

    return pd.concat([out_pd, grand], ignore_index=True)


def _renouvellements_purs_pl(
    df: pd.DataFrame,
    policy_col: str,
    echeance_col: str,
    ren_col: str,
    date_create_col: str,
    start_d: date,
    end_d: date,
    segment: str = "Tous",
) -> int | None:
    """
    Nombre de policyNo distincts arrivant à échéance dans [start_d, end_d]
    ET renouvelés (EST_RENOUVELLER=True) sur le même mois (DATE_CREATE dans la période).
    Utilise Polars pour accélérer, retourne None si indispo.
    """
    if not HAS_POLARS or df is None or df.empty:
        return None

    cols = [policy_col, echeance_col, ren_col, date_create_col]
    seg_col = _find_col(df, ["SEGMENT"])
    if segment != "Tous" and seg_col:
        cols.append(seg_col)

    try:
        dfx = pl.from_pandas(df[cols].copy())
    except Exception:
        return None

    dfx = dfx.with_columns([
        pl.col(echeance_col).cast(pl.Date, strict=False),
        _pl_bool_col(ren_col).alias(ren_col),
        pl.col(date_create_col).cast(pl.Date, strict=False),
        pl.col(policy_col).cast(pl.Utf8, strict=False).fill_null("").str.strip_chars(),
    ])

    if segment != "Tous" and seg_col:
        dfx = dfx.filter(pl.col(seg_col) == segment)

    dfx = dfx.filter(pl.col(policy_col) != "")

    res = (
        dfx.filter(
            (pl.col(echeance_col) >= pl.lit(start_d)) &
            (pl.col(echeance_col) <= pl.lit(end_d)) &
            (pl.col(date_create_col) >= pl.lit(start_d)) &
            (pl.col(date_create_col) <= pl.lit(end_d)) &
            (pl.col(ren_col) == True)
        )
        .select(pl.col(policy_col).n_unique())
    )

    try:
        return int(res.item())
    except Exception:
        return None


def _recap_produit_pandas(
    df: pd.DataFrame,
    df_period: pd.DataFrame,
    id_col: str,
    policy_col: str,
    prod_col: str,
    echeance_col: str,
    effet_col: str,
    est_col: str | None,
    start_d: date,
    end_d: date,
    year_sel: object,
    month_sel: object,
    all_products: list[str] | None,
    total_poss: int,
    total_fav: int,
    total_ren: int,
) -> pd.DataFrame:
    """
    Calcule par produit :
    - # Cas Possibles : identifiants uniques avec DATE_ECHEANCE dans la période
    - # Cas Favorables : identifiants uniques avec DATE_ECHEANCE et DATE_EFFET dans la période
    - Difference : Renouvellements (DATE_CREATE mois) - Favorables
    - % Renouvellement Pur : Favorables / Possibles
    """
    dfp = df.copy()
    dfp[echeance_col] = pd.to_datetime(dfp[echeance_col], errors="coerce")
    dfp[effet_col] = pd.to_datetime(dfp[effet_col], errors="coerce")
    dfp[id_col] = dfp[id_col].astype("string").str.strip().replace("", pd.NA)
    dfp[policy_col] = dfp[policy_col].astype("string").str.strip().replace("", pd.NA)
    dfp[prod_col] = dfp[prod_col].astype("string").str.strip().replace("", "INCONNU").fillna("INCONNU")

    mask_ech = _mask_by_year_month(dfp[echeance_col], year_sel, month_sel)
    mask_eff = _mask_by_year_month(dfp[effet_col], year_sel, month_sel)

    ech_ids = pd.Index(dfp.loc[mask_ech, id_col].dropna().unique())
    eff_ids = pd.Index(dfp.loc[mask_eff, id_col].dropna().unique())
    fav_ids = ech_ids.intersection(eff_ids)

    # Affecte un produit unique par identifiant pour éviter les doubles comptes
    df_ech = dfp.loc[mask_ech, [policy_col, prod_col, echeance_col]].copy()
    df_ech = df_ech.dropna(subset=[policy_col]).sort_values(echeance_col).drop_duplicates(subset=[policy_col], keep="first")
    possibles = (
        df_ech.groupby(prod_col)[policy_col]
        .nunique(dropna=True)
        .rename("# Cas Possibles")
        .reset_index()
    )

    df_eff = dfp.loc[mask_eff & dfp[id_col].isin(fav_ids), [id_col, prod_col, effet_col]].copy()
    df_eff = df_eff.dropna(subset=[id_col]).sort_values(effet_col).drop_duplicates(subset=[id_col], keep="first")
    favorables = (
        df_eff.groupby(prod_col)[id_col]
        .nunique(dropna=True)
        .rename("# Cas Favorables")
        .reset_index()
    )

    renouv = pd.DataFrame(columns=[prod_col, "Renouvellements"])
    if est_col:
        dfm = dfp.copy()
        dfm[policy_col] = dfm[policy_col].astype("string").str.strip().replace("", pd.NA)
        dfm[prod_col] = dfm[prod_col].astype("string").str.strip().replace("", "INCONNU").fillna("INCONNU")
        dfm[est_col] = _to_bool_series(dfm[est_col]).astype(bool)
        dfm = dfm.loc[mask_ech & (dfm[est_col] == True), [policy_col, prod_col]].copy()
        dfm = dfm.dropna(subset=[policy_col]).copy()
        renouv = (
            dfm.groupby(prod_col)[policy_col]
            .nunique(dropna=True)
            .rename("Renouvellements")
            .reset_index()
        )

    recap = possibles.merge(favorables, on=prod_col, how="outer") \
        .merge(renouv, on=prod_col, how="left") \
        .fillna(0)
    if all_products:
        base = pd.DataFrame({prod_col: all_products})
        recap = base.merge(recap, on=prod_col, how="left").fillna(0)
    recap["# Cas Possibles"] = recap["# Cas Possibles"].astype(int)
    recap["# Cas Favorables"] = recap["# Cas Favorables"].astype(int)
    recap["Renouvellements"] = recap["Renouvellements"].astype(int)
    recap["# Cas Favorables"] = recap[["# Cas Favorables", "Renouvellements"]].min(axis=1)
    recap["Difference"] = (recap["Renouvellements"] - recap["# Cas Favorables"]).clip(lower=0)
    recap["% Renouvellement Pur"] = np.where(recap["# Cas Possibles"] > 0, recap["# Cas Favorables"] / recap["# Cas Possibles"], 0.0)

    # tri par volume et ajout Grand Total (somme des lignes)
    recap = recap.sort_values("# Cas Possibles", ascending=False).reset_index(drop=True)
    total_poss_calc = int(recap["# Cas Possibles"].sum())
    total_fav_calc = int(recap["# Cas Favorables"].sum())
    total_diff_calc = int(recap["Difference"].sum())
    total_row = pd.DataFrame([{
        prod_col: "Grand Total",
        "# Cas Possibles": total_poss_calc,
        "# Cas Favorables": total_fav_calc,
        "Difference": total_diff_calc,
        "% Renouvellement Pur": (total_fav_calc / total_poss_calc) if total_poss_calc else 0.0,
    }])
    recap = pd.concat([recap, total_row], ignore_index=True)
    recap = recap.rename(columns={prod_col: "PRODUIT"})
    recap = recap.rename(columns={
        "# Cas Possibles": "Nombre de contrats arrivant à échéance",
        "# Cas Favorables": "Ceux qui ont effectivement renouvelé",
        "Difference": "Contrats non attendus (renouvellements imprévus)",
    })
    recap = recap.drop(columns=["Renouvellements"], errors="ignore")
    return recap


def _renouvellement_effectif_immat(
    df: pd.DataFrame,
    immat_col: str,
    echeance_col: str,
    effet_col: str,
    month_ref: int,
    year_ref: int,
) -> int | None:
    """
    Compte les immatriculations dont l'échéance est dans le mois
    ET ayant une date d'effet dans le même mois.
    """
    if df is None or df.empty:
        return None

    dfx = df.copy()
    dfx[immat_col] = dfx[immat_col].astype("string").str.strip().replace("", pd.NA)
    dfx[echeance_col] = pd.to_datetime(dfx[echeance_col], errors="coerce")
    dfx[effet_col] = pd.to_datetime(dfx[effet_col], errors="coerce")
    dfx = dfx.dropna(subset=[immat_col, echeance_col, effet_col]).copy()
    if dfx.empty:
        return 0

    ech_immat = dfx.loc[
        (dfx[echeance_col].dt.year == year_ref) & (dfx[echeance_col].dt.month == month_ref),
        immat_col,
    ].dropna()
    eff_immat = dfx.loc[
        (dfx[effet_col].dt.year == year_ref) & (dfx[effet_col].dt.month == month_ref),
        immat_col,
    ].dropna()

    if ech_immat.empty or eff_immat.empty:
        return 0

    return int(pd.Index(ech_immat.unique()).intersection(eff_immat.unique()).size)


def _contrats_non_attendus(
    df: pd.DataFrame,
    id_col: str,
    echeance_col: str,
    create_col: str,
    start_d: date,
    end_d: date,
    est_col: str | None = None,
) -> int | None:
    """
    Contrats renouvelés après le mois d'échéance:
    - échéance dans [start_d, end_d]
    - renouvellement (EST_RENOUVELLER=True) avec DATE_CREATE > end_d
    - croisement par identifiant (immat/policy)
    """
    if df is None or df.empty:
        return None

    dfx = df.copy()
    dfx[id_col] = dfx[id_col].astype("string").str.strip().replace("", pd.NA)
    dfx[echeance_col] = pd.to_datetime(dfx[echeance_col], errors="coerce")
    dfx[create_col] = pd.to_datetime(dfx[create_col], errors="coerce")
    dfx = dfx.dropna(subset=[id_col, echeance_col, create_col]).copy()
    if dfx.empty:
        return 0

    ech_mask = (
        (dfx[echeance_col] >= pd.Timestamp(start_d)) &
        (dfx[echeance_col] <= pd.Timestamp(end_d) + pd.Timedelta(hours=23, minutes=59, seconds=59))
    )
    ech_ids = dfx.loc[ech_mask, id_col].dropna().unique()

    ren_mask = dfx[create_col] > pd.Timestamp(end_d) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    if est_col:
        ren_mask = ren_mask & (_to_bool_series(dfx[est_col]) == True)
    ren_ids = dfx.loc[ren_mask, id_col].dropna().unique()

    if len(ech_ids) == 0 or len(ren_ids) == 0:
        return 0

    return int(pd.Index(ech_ids).intersection(pd.Index(ren_ids)).size)


def _add_renouvellement_contrat(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ajoute la colonne Renouvellement_Contrat:
    - '#NA' par défaut
    - immatriculation renseignée pour les occurrences après la première (cumcount > 0).
    """
    immat_col = _find_col(df, ["IMMATRICULATION", "IMMAT", "IMMATRICULE", "NUMERO_IMMATRICULATION"])
    if immat_col is None:
        return df

    df2 = df.copy()
    df2["Renouvellement_Contrat"] = "#NA"

    # Tri par date si disponible pour un cumul cohérent
    date_col = _find_col(df2, ["DATE_CREATE", "DATECREATED", "DATE", "DATE CREATE", "DATE_CREATION", "DATE CREATION"])
    if date_col:
        df2 = df2.sort_values(date_col)

    mask_ren = df2.groupby(immat_col).cumcount() > 0
    df2.loc[mask_ren, "Renouvellement_Contrat"] = df2.loc[mask_ren, immat_col]
    return df2


# =========================
# Streamlit entry
# =========================
def run() -> None:
    st.subheader("🔁 TAUX DE RENOUVELLEMENT — Tableau mensuel")
    _inject_kpi_css()

    df_base = st.session_state.get("df_base")
    if df_base is None or df_base.empty:
        st.info("👉 Charge d’abord tes fichiers Excel dans « Suivi hebdomadaire Auto » (menu à gauche).")
        st.stop()

    df = add_segment(df_base).copy()
    df = _add_renouvellement_contrat(df)

    # ---- Colonnes requises ----
    date_col = _find_col(df, ["DATE_CREATE", "DATECREATED", "DATE", "DATE CREATE", "DATE_CREATION", "DATE CREATION"])
    if date_col is None:
        st.error("Colonne DATE_CREATE introuvable dans df_base.")
        st.stop()
    if date_col != "DATE_CREATE":
        df = df.rename(columns={date_col: "DATE_CREATE"})

    # Normalisation des colonnes clés (nommage stable)
    ech_col = _find_col(df, ["DATE_ECHEANCE", "DATE ECHEANCE", "DATE_ECH", "ECHEANCE", "DATE_EXPIRATION", "DATE_EXPIR"])
    if ech_col and ech_col != "DATE_ECHEANCE":
        df = df.rename(columns={ech_col: "DATE_ECHEANCE"})
    effet_col = _find_col(df, ["DATE_EFFET", "DATE EFFET", "DATE_EFF", "DATE_EFFET_CONTRAT", "DATE EFFET CONTRAT"])
    if effet_col and effet_col != "DATE_EFFET":
        df = df.rename(columns={effet_col: "DATE_EFFET"})
    est_col_raw = _find_col(df, ["EST_RENOUVELLER", "EST RENOUVELLER", "EST_RENOUVELER", "EST RENOUVELER"])
    if est_col_raw and est_col_raw != "EST_RENOUVELLER":
        df = df.rename(columns={est_col_raw: "EST_RENOUVELLER"})

    # policyNo
    policy_col = _find_col(df, ["POLICYNO", "POLICY_NO", "POLICYNUMBER", "POLICY_NUMBER", "POLICY", "POLICYNO ", "POLICYNO", "policyNo"])
    if policy_col is None:
        st.error("Colonne policyNo introuvable. Attendu: policyNo / POLICYNO / POLICY_NO ...")
        st.stop()

    # produit
    prod_col = _find_col(df, ["PRODUIT", "PRODUCT"])
    if prod_col is None:
        df["PRODUIT"] = "INCONNU"
        prod_col = "PRODUIT"

    # colonnes renouvellement
    est_col = _find_col(df, ["EST_RENOUVELLER", "EST_RENOUVELER"])
    plan_col = _find_col(df, ["RENOUVELLEMENT", "RENOUVELLEMENT_ATTENDU", "RENOUVELLEMENT_ATTENTES"])
    if est_col is None and plan_col is not None:
        est_col = plan_col  # fallback

    # dates valides pour bornes
    df["DATE_CREATE"] = pd.to_datetime(df["DATE_CREATE"], errors="coerce")
    df = df.dropna(subset=["DATE_CREATE"]).copy()
    if df.empty:
        st.warning("Aucune date valide dans DATE_CREATE.")
        st.stop()

    min_d = df["DATE_CREATE"].min().date()
    max_d = df["DATE_CREATE"].max().date()

    # UI (keys uniques pour éviter conflits avec Suivi hebdo)
    years = sorted(df["DATE_CREATE"].dt.year.dropna().unique().tolist())
    year_options = ["Tous"] + years
    default_year = int(max_d.year)
    year_index = year_options.index(default_year) if default_year in year_options else 0

    c1, c2, c3, c4 = st.columns([0.25, 0.35, 0.25, 0.15])

    with c1:
        year = st.selectbox(
            "Année",
            year_options,
            index=year_index,
            key="tr_year",
        )

    with c2:
        if year == "Tous":
            months = sorted(df["DATE_CREATE"].dt.month.dropna().unique().tolist())
        else:
            months = sorted(df[df["DATE_CREATE"].dt.year == int(year)]["DATE_CREATE"].dt.month.dropna().unique().tolist())
        if not months:
            months = list(range(1, 13))
        month_options = ["Tous"] + months
        default_month = int(max_d.month)
        month_index = month_options.index(default_month) if default_month in month_options else 0
        month = st.selectbox(
            "Mois",
            month_options,
            index=month_index,
            format_func=lambda m: "Tous" if m == "Tous" else _month_name_fr(int(m)).title(),
            key="tr_month",
        )

    # Options de segment dynamiques selon les données
    seg_unique = (
        df["SEGMENT"]
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    base_segments = ["Agent", "Courtier", "Direct"]
    seg_set = {s for s in seg_unique if isinstance(s, str)}
    seg_options = ["Tous"] + [s for s in base_segments] + sorted(seg_set - set(base_segments))

    with c3:
        segment = st.selectbox("Segment", seg_options, index=0, key="tr_segment")

    with c4:
        mois_complet = st.checkbox("Mois complet", value=True, key="tr_full_month")

    month_filter = None
    if year == "Tous" and month == "Tous":
        start_d, end_d = min_d, max_d
    elif year == "Tous" and month != "Tous":
        start_d, end_d = min_d, max_d
        month_filter = int(month)
    elif year != "Tous" and month == "Tous":
        start_d = date(int(year), 1, 1)
        end_d = date(int(year), 12, 31)
        if start_d < min_d:
            start_d = min_d
        if end_d > max_d:
            end_d = max_d
    else:
        start_m, end_m = _month_bounds(int(year), int(month))
        start_d, end_d = start_m, end_m

    if not (year == "Tous" or month == "Tous"):
        if not mois_complet:
            cc1, cc2 = st.columns([0.5, 0.5])
            last_day = calendar.monthrange(int(year), int(month))[1]
            with cc1:
                day_start = st.number_input("Jour début", 1, last_day, 1, 1, key="tr_day_start")
            with cc2:
                day_end = st.number_input("Jour fin", 1, last_day, last_day, 1, key="tr_day_end")

            if int(day_start) > int(day_end):
                day_start, day_end = day_end, day_start

            start_d = date(int(year), int(month), int(day_start))
            end_d = date(int(year), int(month), int(day_end))

    # clamp aux bornes data
    if start_d < min_d:
        start_d = min_d
    if end_d > max_d:
        end_d = max_d
    if start_d > end_d:
        st.warning("Période hors des données disponibles.")
        st.stop()

    # filtre segment (sur df) AVANT calcul
    if segment != "Tous":
        seg_col = _find_col(df, ["SEGMENT"])
        if seg_col is not None:
            df = df[
                df[seg_col]
                .astype("string")
                .str.strip()
                .str.lower()
                .eq(str(segment).strip().lower())
            ].copy()

    if df.empty:
        st.warning("Aucune donnée après filtre segment ou période. Vérifie les valeurs de SEGMENT/OPERATEUR (BROKER) et les dates.")
        st.stop()

    df_seg = df.copy()

    # timestamps inclusifs
    start_ts = pd.Timestamp(start_d)
    end_ts = pd.Timestamp(end_d) + pd.Timedelta(hours=23, minutes=59, seconds=59)

    # titre dynamique
    if year == "Tous" and month == "Tous":
        title = f"TOUS_{_segment_suffix(segment)}"
    elif year == "Tous":
        title = f"{_month_name_fr(int(month))}_TOUTES_ANNÉES_{_segment_suffix(segment)}"
    elif month == "Tous":
        title = f"TOUS_MOIS_{int(year)}_{_segment_suffix(segment)}"
    else:
        title = f"{_month_name_fr(int(month))}_{int(year)}_{_segment_suffix(segment)}"

    st.markdown(
        f"""
        <div style="
            padding:16px 14px;
            border-radius:22px;
            background:#1F4E79;
            color:white;
            font-weight:900;
            font-size:28px;
            text-align:center;
            letter-spacing:0.6px;
            margin-top:8px;
            margin-bottom:18px;
        ">
            {title}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Filtre période sur le DF segmenté
    df_period = df[(df["DATE_CREATE"] >= start_ts) & (df["DATE_CREATE"] <= end_ts)].copy()
    if month_filter is not None:
        df_period = df_period[df_period["DATE_CREATE"].dt.month == month_filter].copy()
    if df_period.empty:
        st.warning("Aucune donnée sur cette période (après filtre segment/date).")
        st.stop()

    # calcul table
    if HAS_POLARS:
        df_out = _compute_polars(df=df_period, policy_col=policy_col, start_dt=start_ts.to_pydatetime(), end_dt=end_ts.to_pydatetime())
    else:
        df_out = _compute_pandas(df=df_period, policy_col=policy_col, start_dt=start_ts, end_dt=end_ts)

    if df_out is None or df_out.empty:
        st.warning("Aucune donnée sur cette période.")
        st.stop()

    # CA (Prime TTC / Prime nette + Accessoires) sur la période et segment filtrés
    prime_ttc_col = _find_col(df_period, ["PRIME_TTC", "PRIME TTC", "PRIME_TTC_(FCFA)", "PRIME_TTC_FCFA", "PRIME_TTC_(XOF)"])
    prime_nette_col = _find_col(df_period, ["PRIME_NETTE", "PRIME NETTE"])
    accessoires_col = _find_col(df_period, ["ACCESSOIRES", "ACCESSORY", "ACCESSORIES"])
    ca_ttc = float(df_period[prime_ttc_col].sum()) if prime_ttc_col else 0.0
    ca_net = float((df_period[prime_nette_col].sum() if prime_nette_col else 0.0) + (df_period[accessoires_col].sum() if accessoires_col else 0.0))
    # CA même période dans le mois (jours min/max de la sélection, même mois/année que end_d)
    day_start = start_d.day
    day_end = end_d.day
    month_ref = end_d.month
    year_ref = end_d.year
    df_same_month = df_period[
        (df["DATE_CREATE"].dt.month == month_ref)
        & (df["DATE_CREATE"].dt.year == year_ref)
        & (df["DATE_CREATE"].dt.day >= day_start)
        & (df["DATE_CREATE"].dt.day <= day_end)
    ].copy()
    ca_meme_periode = float(df_same_month[prime_ttc_col].sum()) if prime_ttc_col else 0.0

    # KPI depuis Grand Total
    grand = df_out[df_out["PRODUIT"].astype(str).str.strip().str.lower().eq("grand total")]
    if grand.empty:
        st.warning("Ligne Grand Total introuvable.")
        st.dataframe(_style_table(df_out), use_container_width=True, hide_index=True)
        st.stop()

    g = grand.iloc[0].to_dict()
    total = int(g.get("Total", 0) or 0)
    newb = int(g.get("# New Business", 0) or 0)
    ren = int(g.get("# Renouvellement", 0) or 0)
    pct = float(g.get("% Renouvellement", 0) or 0.0)

    ca_row = st.columns(3)
    with ca_row[0]:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(ca_ttc)}</div>
              <div class="kpi_label">CA (Prime TTC)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with ca_row[1]:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(ca_net)}</div>
              <div class="kpi_label">CA net (Prime nette + Accessoires)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with ca_row[2]:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(ca_meme_periode)}</div>
              <div class="kpi_label">CA même période (mois courant)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # Contrats arrivant à échéance (distinct policyNo sur DATE_ECHEANCE dans le mois sélectionné)
    echeance_col = _find_col(df_seg, ["DATE_ECHEANCE", "DATE_ECH", "ECHEANCE", "DATE_EXPIRATION", "DATE_EXPIR"])
    contrats_echeance = None
    if echeance_col:
        df_ech = df_seg.copy()
        df_ech[echeance_col] = pd.to_datetime(df_ech[echeance_col], errors="coerce")
        df_ech = df_ech.dropna(subset=[echeance_col]).copy()
        mask_ech = _mask_by_year_month(df_ech[echeance_col], year, month)
        ech = df_ech.loc[mask_ech, policy_col].astype("string").str.strip()
        ech = ech.replace("", pd.NA).dropna()
        contrats_echeance = int(ech.nunique())
    else:
        st.warning("Colonne DATE_ECHEANCE manquante : 'Nombre de contrats arrivant à échéance' utilise le total des souscriptions.")
        contrats_echeance = total

    contrats_renouveles = ren
    if echeance_col and est_col:
        df_ren = df_seg.copy()
        df_ren[echeance_col] = pd.to_datetime(df_ren[echeance_col], errors="coerce")
        df_ren = df_ren.dropna(subset=[echeance_col, policy_col]).copy()
        df_ren[est_col] = _to_bool_series(df_ren[est_col])
        mask_ech = _mask_by_year_month(df_ren[echeance_col], year, month)
        ren_ids = df_ren.loc[mask_ech & (df_ren[est_col] == True), policy_col].astype("string").replace("", pd.NA).dropna()
        contrats_renouveles = int(ren_ids.nunique())
    elif echeance_col:
        st.warning("Colonne EST RENOUVELLER manquante : calcul 'effectivement renouvelé' basé sur cette colonne impossible.")

    ren_create = None
    if est_col and "DATE_CREATE" in df_period.columns:
        ren_create = _renewals_by_create(
            df=df_period,
            policy_col=policy_col,
            est_col=est_col,
            create_col="DATE_CREATE",
            year_sel=year,
            month_sel=month,
        )

    # KPI affichage (ceux demandés)
    k1, k2, k3 = st.columns(3)

    with k1:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(total)}</div>
              <div class="kpi_label">Toutes les souscriptions</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k2:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(newb)}</div>
              <div class="kpi_label">New Business</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k3:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(ren)}</div>
              <div class="kpi_label">Renouvellements</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    k4a, k4b, k4c = st.columns([0.25, 0.5, 0.25])
    with k4b:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_pct(pct)}</div>
              <div class="kpi_label">Part des renouvellements</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # KPI renouv spécifiques mois m
    ren_base = ren_create if ren_create is not None else ren
    contrats_imprevus = max(0, ren_base - contrats_renouveles)
    taux_ren_pur = contrats_renouveles / contrats_echeance if contrats_echeance else 0.0

    r1, r2, r3, r4 = st.columns(4)

    with r1:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(contrats_echeance)}</div>
              <div class="kpi_label">Nombre de contrats arrivant à échéance au mois m</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with r2:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(contrats_renouveles)}</div>
              <div class="kpi_label">Ceux qui ont effectivement renouvelé au mois m</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with r3:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_space_int(contrats_imprevus)}</div>
              <div class="kpi_label">Contrats non attendus (renouvellements imprévus)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with r4:
        st.markdown(
            f"""
            <div class="kpi_card">
              <div class="kpi_value">{_fmt_pct(taux_ren_pur)}</div>
              <div class="kpi_label">Taux de renouvellement pur du mois m</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # Tableau récapitulatif (valeurs réelles)
    recap_df = pd.DataFrame([{
        "Nombre de contrats arrivant à échéance": contrats_echeance,
        "Ceux qui ont effectivement renouvelé": contrats_renouveles,
        "Contrats non attendus (renouvellements imprévus)": max(0, contrats_imprevus),
        "% Renouvellement": taux_ren_pur,
    }])

    def _style_recap(df_in: pd.DataFrame):
        sty = df_in.style.format({
            "Nombre de contrats arrivant à échéance": _fmt_space_int,
            "Ceux qui ont effectivement renouvelé": _fmt_space_int,
            "Contrats non attendus (renouvellements imprévus)": _fmt_space_int,
            "% Renouvellement": _fmt_pct,
        })
        sty = sty.set_table_styles([
            {"selector": "th", "props": [("background-color", "#A1D884"), ("color", "black"), ("font-weight", "bold"), ("text-align", "center")]}
        ])
        sty = sty.set_properties(**{
            "text-align": "center",
            "font-weight": "bold",
        })
        sty = sty.apply(
            lambda row: ["background-color:#F28E1C;color:white;font-weight:bold;" if row.name == "Contrats non attendus (renouvellements imprévus)" else "" for _ in row],
            axis=1
        )
        return sty

    st.markdown("### 📊 Récap renouvellements (période sélectionnée)")
    st.dataframe(_style_recap(recap_df), use_container_width=True, hide_index=True)

    # Top 20 agents: contrats à renouveler / renouvelés / taux
    agent_col = _find_col(df_seg, ["NOM_AGENT", "NOM AGENT", "AGENT", "SOUSCRIPTEUR", "SUBSCRIBER", "UTILISATEUR", "Utilisateur"])
    if agent_col and echeance_col and est_col:
        top_df = df_seg[[agent_col, policy_col, echeance_col, est_col]].copy()
        top_df[agent_col] = top_df[agent_col].astype("string").str.strip().replace("", pd.NA)
        top_df[echeance_col] = pd.to_datetime(top_df[echeance_col], errors="coerce")
        top_df[est_col] = _to_bool_series(top_df[est_col])
        top_df = top_df.dropna(subset=[agent_col, echeance_col, policy_col]).copy()

        mask_ech = _mask_by_year_month(top_df[echeance_col], year, month)

        ech = top_df.loc[mask_ech].groupby(agent_col)[policy_col].nunique().rename("Contrats à renouveler")
        eff = top_df.loc[mask_ech & (top_df[est_col] == True)].groupby(agent_col)[policy_col].nunique().rename("Contrats renouvelés")

        top_tbl = ech.to_frame().merge(eff, left_index=True, right_index=True, how="left").fillna(0)
        top_tbl["Taux de renouvellement"] = np.where(
            top_tbl["Contrats à renouveler"] > 0,
            top_tbl["Contrats renouvelés"] / top_tbl["Contrats à renouveler"],
            0.0,
        )
        top_tbl = top_tbl.sort_values("Contrats à renouveler", ascending=False).head(20).reset_index()
        top_tbl = top_tbl.rename(columns={agent_col: "Agent"})

        st.markdown("### 🧾 Top 20 agents — Renouvellements")
        st.dataframe(
            top_tbl.style.format({
                "Contrats à renouveler": _fmt_space_int,
                "Contrats renouvelés": _fmt_space_int,
                "Taux de renouvellement": _fmt_pct,
            }),
            use_container_width=True,
            hide_index=True,
        )
    else:
        if not agent_col:
            st.warning("Colonne agent manquante : tableau Top 20 agents indisponible.")
        if not echeance_col or not est_col:
            st.warning("Colonnes DATE_ECHEANCE/EST_RENOUVELLER manquantes : tableau Top 20 agents indisponible.")

    # Prépare un export Excel (sidebar)
    export_tables: dict[str, pd.DataFrame] = {
        "Recap_renouvellement": recap_df.copy(),
        "Detail_produits": df_out.copy(),
    }

    # Tableau synthèse Agents/Courtiers/Leadway/Général
    if year != "Tous" and month != "Tous":
        month_label = f"{_month_name_fr(int(month)).title()} {int(year)}"
    else:
        month_label = "Période sélectionnée"

    def _segment_metrics(df_all: pd.DataFrame, seg_name: str) -> dict[str, object]:
        if seg_name == "Général":
            df_seg = df_all.copy()
        elif seg_name == "Leadway":
            df_seg = df_all[df_all["SEGMENT"].astype("string").str.upper() == "DIRECT"].copy()
        else:
            df_seg = df_all[df_all["SEGMENT"].astype("string").str.upper() == seg_name.upper()].copy()

        if df_seg.empty:
            return {
                "total": 0,
                "newb": 0,
                "ren": 0,
                "part_ren": 0.0,
                "ech": 0,
                "fav": 0,
                "diff": 0,
                "taux_pur": 0.0,
            }

        df_create = df_seg.copy()
        df_create["DATE_CREATE"] = pd.to_datetime(df_create["DATE_CREATE"], errors="coerce")
        mask_create = _mask_by_year_month(df_create["DATE_CREATE"], year, month)
        df_create = df_create.loc[mask_create].copy()

        total = int(df_create[policy_col].astype("string").replace("", pd.NA).dropna().nunique()) if not df_create.empty else 0
        if est_col:
            est_vals = _to_bool_series(df_create[est_col]).astype(bool) if not df_create.empty else pd.Series([], dtype=bool)
            ren = int(df_create.loc[est_vals, policy_col].astype("string").replace("", pd.NA).dropna().nunique()) if not df_create.empty else 0
            newb = int(df_create.loc[~est_vals, policy_col].astype("string").replace("", pd.NA).dropna().nunique()) if not df_create.empty else 0
        else:
            ren = 0
            newb = 0

        part_ren = (ren / total) if total else 0.0

        if "DATE_ECHEANCE" in df_seg.columns:
            df_ech = df_seg.copy()
            df_ech["DATE_ECHEANCE"] = pd.to_datetime(df_ech["DATE_ECHEANCE"], errors="coerce")
            mask_ech = _mask_by_year_month(df_ech["DATE_ECHEANCE"], year, month)
            ech = df_ech.loc[mask_ech, policy_col].astype("string").replace("", pd.NA).dropna().nunique()
        else:
            ech = 0

        # Favorables = échéance et effet dans le mois
        if "DATE_ECHEANCE" in df_seg.columns and "DATE_EFFET" in df_seg.columns:
            dfx = df_seg.copy()
            dfx["DATE_ECHEANCE"] = pd.to_datetime(dfx["DATE_ECHEANCE"], errors="coerce")
            dfx["DATE_EFFET"] = pd.to_datetime(dfx["DATE_EFFET"], errors="coerce")
            dfx = dfx.dropna(subset=["DATE_ECHEANCE", "DATE_EFFET", policy_col]).copy()
            mask_ech = _mask_by_year_month(dfx["DATE_ECHEANCE"], year, month)
            mask_eff = _mask_by_year_month(dfx["DATE_EFFET"], year, month)
            id_col = _find_col(dfx, ["IMMATRICULATION", "IMMAT", "IMMATRICULE", "NUMERO_IMMATRICULATION"]) or policy_col
            ech_ids = pd.Index(dfx.loc[mask_ech, id_col].astype("string").replace("", pd.NA).dropna().unique())
            eff_ids = pd.Index(dfx.loc[mask_eff, id_col].astype("string").replace("", pd.NA).dropna().unique())
            fav = int(ech_ids.intersection(eff_ids).size)
        else:
            fav = 0

        diff = max(0, ren - fav)
        taux_pur = (fav / ech) if ech else 0.0

        return {
            "total": total,
            "newb": newb,
            "ren": ren,
            "part_ren": part_ren,
            "ech": int(ech),
            "fav": int(fav),
            "diff": int(diff),
            "taux_pur": taux_pur,
        }

    segments = ["Agent", "Courtier", "Leadway", "Général"]
    metrics_map = {seg: _segment_metrics(df, seg) for seg in segments}

    age_mean = _calc_age_mean(df_period, end_d)
    ind_rows = [
        (f"Toutes les souscriptions ({month_label})", "total", "int"),
        ("Nouvelles affaires", "newb", "int"),
        (f"Tous les cas de renouvellement ({month_label})", "ren", "int"),
        ("Part des renouvellements dans les souscriptions", "part_ren", "pct"),
        (f"Nombre de contrats arrivant à échéance en {month_label.split()[0]}", "ech", "int"),
        ("Contrats non attendus (renouvellements imprévus)", "diff", "int"),
        (f"Contrats arrivant à échéance en {month_label.split()[0]} et effectivement renouvelés", "fav", "int"),
        (f"Taux de renouvellement pur ({month_label})", "taux_pur", "pct"),
        ("Age moyen", "age", "int"),
    ]

    data_rows = []
    for label, key, kind in ind_rows:
        row = {"Indicateur": label}
        for seg in segments:
            col_name = seg + "s" if seg in ["Agent", "Courtier"] else seg
            if key == "age":
                if seg == "Général" and age_mean is not None:
                    row[col_name] = _fmt_space_int(age_mean)
                else:
                    row[col_name] = ""
                continue
            val = metrics_map[seg].get(key, 0)
            if kind == "pct":
                row[col_name] = _fmt_pct(val)
            else:
                row[col_name] = _fmt_space_int(val)
        data_rows.append(row)

    df_summary = pd.DataFrame(data_rows)
    col_map = {"Agents": "Agents", "Courtiers": "Courtiers", "Leadway": "Leadway", "Général": "Général"}
    df_summary = df_summary.rename(columns=col_map)

    def _style_summary(df_in: pd.DataFrame):
        sty = df_in.style
        sty = sty.set_table_styles([
            {"selector": "thead th", "props": [("background-color", "#8FD14F"), ("color", "black"), ("font-weight", "bold"), ("text-align", "center")]},
            {"selector": "tbody td", "props": [("text-align", "center"), ("border", "1px solid #d0d0d0"), ("padding", "8px")]},
            {"selector": "tbody td:first-child", "props": [("text-align", "left"), ("font-weight", "bold")]},
        ])
        return sty

    st.markdown("### 📋 Synthèse par segment")
    st.dataframe(_style_summary(df_summary), use_container_width=True, hide_index=True)
    export_tables["Synthese_par_segment"] = df_summary.copy()

    # Tableau par produit (renouvellement pur)
    if echeance_col and est_col:
        try:
            id_col = _find_col(df, ["IMMATRICULATION", "IMMAT", "IMMATRICULE", "NUMERO_IMMATRICULATION"]) or policy_col
            effet_col = _find_col(df, ["DATE_EFFET", "DATE_EFF", "DATE EFFET", "DATE_EFFET_CONTRAT", "DATE EFFET CONTRAT"])
            if not effet_col:
                st.warning("Colonne DATE_EFFET manquante : tableau par produit des renouvellements effectifs indisponible.")
                raise ValueError("DATE_EFFET missing")
            recap_prod = _recap_produit_pandas(
                df=df,
                df_period=df_period,
                id_col=id_col,
                policy_col=policy_col,
                prod_col=prod_col,
                echeance_col=echeance_col,
                effet_col=effet_col,
                est_col=est_col,
                start_d=start_d,
                end_d=end_d,
                year_sel=year,
                month_sel=month,
                all_products=[p for p in df_out["PRODUIT"].astype("string").dropna().unique().tolist() if str(p).strip().lower() != "grand total"],
                total_poss=contrats_echeance,
                total_fav=contrats_renouveles,
                total_ren=ren_base,
            )
            if recap_prod is not None and not recap_prod.empty:
                st.markdown("### 🗂️ Renouvellement pur par produit")
                st.dataframe(_style_recap_prod(recap_prod), use_container_width=True, hide_index=True)
                export_tables["Renouvellement_pur_produit"] = recap_prod.copy()
        except Exception as e:
            st.warning(f"Impossible de calculer le tableau par produit : {e}")

    # Tableau âge véhicule par produit
    mise_col = _find_col(df_period, ["DATE_1ERE_MISE_EN_CIRCULATION", "DATE 1ERE MISE EN CIRCULATION", "DATE_MISE_EN_CIRCULATION"])
    immat_col = _find_col(df_period, ["IMMATRICULATION", "IMMAT", "IMMATRICULE", "NUMERO_IMMATRICULATION", "IMMATRICULATION_VEHICULE", "PLAQUE"])
    if mise_col and immat_col:
        age_tbl = _vehicule_age_table(
            df_period=df_period,
            prod_col=prod_col,
            immat_col=immat_col,
            mise_col=mise_col,
            ref_date=end_d,
        )
        if age_tbl is not None and not age_tbl.empty:
            st.markdown("### 🧾 Âge véhicule par produit")
            st.dataframe(_style_age_table(age_tbl), use_container_width=True, hide_index=True)
            export_tables["Age_vehicule_par_produit"] = age_tbl.copy()
    else:
        if not mise_col:
            st.warning("Colonne DATE 1ere MISE EN CIRCULATION manquante : tableau âge véhicule indisponible.")
        if not immat_col:
            st.warning("Colonne immatriculation manquante : tableau âge véhicule indisponible.")
        with st.expander("Voir les colonnes disponibles"):
            st.write(list(df_period.columns))

    st.divider()
    st.markdown("## 📌 Détail par produit")
    st.dataframe(_style_table(df_out), use_container_width=True, hide_index=True)

    if export_tables:
        with st.sidebar.expander("📥 Exporter les tableaux (Excel)", expanded=False):
            excel_buf = BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                for sheet_name, table in export_tables.items():
                    safe_name = sheet_name[:31]
                    table.to_excel(writer, index=False, sheet_name=safe_name)
            excel_buf.seek(0)
            st.sidebar.download_button(
                "Télécharger les tableaux",
                data=excel_buf,
                file_name=f"taux_renouvellement_{title}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
