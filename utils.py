# SUIVI_AUTO/utils.py
from __future__ import annotations

import numpy as np
import pandas as pd


# ----------------- BASE NORMALIZATION -----------------
def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise les colonnes (UPPER + underscores) et crée les champs essentiels.

    Champs visés :
      - OPERATEUR (str)
      - DATE_CREATE (datetime)
      - PRIME_TTC (num)
      - TAXES (num)
      - POOLS_TPV (0/1)  ✅ 1=POOL, 0=Hors POOL
      - EST_RENOUVELLER (bool)
      - NOM_AGENT (str)
      - PRODUIT (str)
    """
    df = df.copy()

    # Colonnes -> UPPER + underscores (robuste)
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\u00A0", " ", regex=False)
        .str.upper()
        .str.replace(" ", "_", regex=False)
        .str.replace(r"_+", "_", regex=True)
        .str.strip("_")
    )

    # -------- OPERATEUR --------
    if "OPERATEUR" not in df.columns:
        if "OPERATOR" in df.columns:
            df["OPERATEUR"] = df["OPERATOR"]
        elif "OPÉRATEUR" in df.columns:
            df["OPERATEUR"] = df["OPÉRATEUR"]
        else:
            df["OPERATEUR"] = ""
    df["OPERATEUR"] = df["OPERATEUR"].astype(str).str.strip()

    # -------- DATE_CREATE --------
    if "DATE_CREATE" in df.columns:
        s = df["DATE_CREATE"]
        dt = pd.to_datetime(s, errors="coerce")
        num = pd.to_numeric(s, errors="coerce")
        mask = dt.isna() & num.notna()
        if mask.any():
            dt.loc[mask] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask], unit="D")
        df["DATE_CREATE"] = dt
    elif "DATE" in df.columns:
        s = df["DATE"]
        dt = pd.to_datetime(s, errors="coerce")
        num = pd.to_numeric(s, errors="coerce")
        mask = dt.isna() & num.notna()
        if mask.any():
            dt.loc[mask] = pd.to_datetime("1899-12-30") + pd.to_timedelta(num.loc[mask], unit="D")
        df["DATE_CREATE"] = dt
    else:
        df["DATE_CREATE"] = pd.NaT

    # -------- PRIME_TTC --------
    if "PRIME_TTC" in df.columns:
        df["PRIME_TTC"] = pd.to_numeric(df["PRIME_TTC"], errors="coerce").fillna(0)
    elif "PRIME_TTC_(FCFA)" in df.columns:
        df["PRIME_TTC"] = pd.to_numeric(df["PRIME_TTC_(FCFA)"], errors="coerce").fillna(0)
    elif "PRIME_TTC_FCFA" in df.columns:
        df["PRIME_TTC"] = pd.to_numeric(df["PRIME_TTC_FCFA"], errors="coerce").fillna(0)
    elif "PRIME_TTC_(XOF)" in df.columns:
        df["PRIME_TTC"] = pd.to_numeric(df["PRIME_TTC_(XOF)"], errors="coerce").fillna(0)
    else:
        df["PRIME_TTC"] = 0

    # -------- TAXES --------
    if "TAXES" in df.columns:
        df["TAXES"] = pd.to_numeric(df["TAXES"], errors="coerce").fillna(0)
    else:
        df["TAXES"] = 0

    # -------- POOLS_TPV --------
    if "POOLS_TPV" in df.columns:
        df["POOLS_TPV"] = pd.to_numeric(df["POOLS_TPV"], errors="coerce").fillna(0).astype(int)
    elif "POOL_TPV" in df.columns:
        df["POOLS_TPV"] = pd.to_numeric(df["POOL_TPV"], errors="coerce").fillna(0).astype(int)
    elif "POOLS_TPV_" in df.columns:
        df["POOLS_TPV"] = pd.to_numeric(df["POOLS_TPV_"], errors="coerce").fillna(0).astype(int)
    else:
        df["POOLS_TPV"] = 0

    df["POOLS_TPV"] = np.where(df["POOLS_TPV"] == 1, 1, 0).astype(int)

    # -------- EST_RENOUVELLER --------
    def _to_bool_ren(x) -> bool:
        s = str(x).strip().upper()
        return s in {
            "TRUE", "1", "OUI", "YES", "Y", "VRAI",
            "RENOUVELLEMENT", "RENOUVELER", "RENOUVELLER", "RENOUVELLE",
        }

    if "EST_RENOUVELLER" in df.columns:
        df["EST_RENOUVELLER"] = df["EST_RENOUVELLER"].map(_to_bool_ren)
    elif "RENOUVELLEMENT" in df.columns:
        df["EST_RENOUVELLER"] = df["RENOUVELLEMENT"].map(_to_bool_ren)
    else:
        df["EST_RENOUVELLER"] = False

    # -------- NOM_AGENT --------
    if "NOM_AGENT" in df.columns:
        df["NOM_AGENT"] = df["NOM_AGENT"].astype(str).str.strip().replace("", pd.NA)
    elif "AGENT" in df.columns:
        df["NOM_AGENT"] = df["AGENT"].astype(str).str.strip().replace("", pd.NA)
    else:
        df["NOM_AGENT"] = pd.NA

    # -------- PRODUIT --------
    if "PRODUIT" not in df.columns and "PRODUCT" in df.columns:
        df["PRODUIT"] = df["PRODUCT"]
    if "PRODUIT" in df.columns:
        df["PRODUIT"] = df["PRODUIT"].astype(str).str.strip().replace("", pd.NA)
    else:
        df["PRODUIT"] = pd.NA

    return df


# ----------------- SEGMENT / INTERMEDIAIRE -----------------
def add_segment(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crée la colonne SEGMENT ∈ {"Agent","Courtier","Direct"} (si absente).
    - Courtier si OPERATEUR contient 'BROKER' ou 'COURTIER'
    - sinon Agent
    (Direct est conservé uniquement si déjà présent dans la colonne SEGMENT)
    """
    df = df.copy()

    op = df.get("OPERATEUR", pd.Series([""] * len(df), index=df.index)).astype("string").fillna("").str.strip()
    nom_agent = df.get("NOM_AGENT", pd.Series([""] * len(df), index=df.index)).astype("string").fillna("").str.strip()

    is_broker = op.str.contains(r"BROKER|COURTIER", case=False, na=False)
    is_direct = nom_agent.str.contains(r"LEADWAY", case=False, na=False) & (~is_broker)
    seg_calc = np.where(is_broker, "Courtier", np.where(is_direct, "Direct", "Agent"))
    df["SEGMENT"] = seg_calc
    return df


def add_intermediaire(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crée la colonne INTERMEDIAIRE (si absente).
    - Agent => NOM_AGENT (si dispo), sinon OPERATEUR
    - Courtier => OPERATEUR (si dispo), sinon NOM_AGENT
    """
    df = df.copy()

    if "INTERMEDIAIRE" in df.columns:
        df["INTERMEDIAIRE"] = df["INTERMEDIAIRE"].astype(str).str.strip().replace("", pd.NA)
        return df

    if "SEGMENT" not in df.columns:
        df = add_segment(df)

    nom_agent = df.get("NOM_AGENT", pd.Series([pd.NA] * len(df), index=df.index))
    operateur = df.get("OPERATEUR", pd.Series([""] * len(df), index=df.index))

    nom_agent = nom_agent.astype("string").str.strip().replace("", pd.NA)
    operateur = operateur.astype("string").str.strip().replace("", pd.NA)

    inter = np.where(
        df["SEGMENT"].astype(str).str.upper().eq("AGENT"),
        nom_agent.fillna(operateur),
        operateur.fillna(nom_agent),
    )

    df["INTERMEDIAIRE"] = pd.Series(inter, index=df.index).astype("string").str.strip().replace("", pd.NA)
    return df


# ----------------- FORMATS -----------------
def fmt_int(x) -> str:
    """Format entier avec séparateur espace."""
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except Exception:
        return "0"


def _to_num_robust(series: pd.Series) -> pd.Series:
    """Conversion numerique robuste (espaces, separateurs, textes)."""
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


# ----------------- KPI HELPERS -----------------
def agents_fideles(df_f: pd.DataFrame) -> int:
    """Agents fidèles = agents présents sur tous les mois (YYYYMM) dans la période filtrée."""
    if df_f is None or df_f.empty:
        return 0
    if "NOM_AGENT" not in df_f.columns or "DATE_CREATE" not in df_f.columns:
        return 0

    tmp = df_f.dropna(subset=["NOM_AGENT", "DATE_CREATE"]).copy()
    if tmp.empty:
        return 0

    tmp["YM"] = tmp["DATE_CREATE"].dt.strftime("%Y%m")
    total_months = tmp["YM"].nunique()
    if total_months == 0:
        return 0

    months_by_agent = tmp.groupby("NOM_AGENT")["YM"].nunique()
    return int((months_by_agent == total_months).sum())


def compute_kpis(df_f: pd.DataFrame, objectif_annuel: float) -> dict:
    """
    KPI utilisés dans app.py / tasks (sur df filtré).
    IMPORTANT: POOLS_TPV = 1 (POOL), 0 (Hors POOL)
    """
    base = {
        "taxes": 0, "ca": 0,
        "montant_pool": 0, "montant_hors_pool": 0,
        "souscriptions": 0, "renouvellements": 0, "nouvelles_affaires": 0,
        "nb_agents_actifs": 0, "nb_agents_fideles": 0,
        "ratio_obj": 0, "pct_pool": 0, "pct_hors_pool": 0, "pct_ren": 0, "pct_new": 0,
    }
    if df_f is None or df_f.empty:
        return base

    dfx = df_f.copy()

    dfx["PRIME_TTC"] = pd.to_numeric(dfx.get("PRIME_TTC", 0), errors="coerce").fillna(0)
    dfx["TAXES"] = pd.to_numeric(dfx.get("TAXES", 0), errors="coerce").fillna(0)

    # Regle metier: CA = PRIME_NETTE + ACCESSOIRES
    prime_nette = _to_num_robust(dfx["PRIME_NETTE"]).fillna(0) if "PRIME_NETTE" in dfx.columns else pd.Series(0, index=dfx.index)
    accessoires = _to_num_robust(dfx["ACCESSOIRES"]).fillna(0) if "ACCESSOIRES" in dfx.columns else pd.Series(0, index=dfx.index)
    use_ca_net = ("PRIME_NETTE" in dfx.columns) or ("ACCESSOIRES" in dfx.columns)
    ca_series = (prime_nette + accessoires) if use_ca_net else dfx["PRIME_TTC"]

    if "POOLS_TPV" not in dfx.columns:
        dfx["POOLS_TPV"] = 0
    dfx["POOLS_TPV"] = pd.to_numeric(dfx["POOLS_TPV"], errors="coerce").fillna(0).astype(int)
    dfx["POOLS_TPV"] = np.where(dfx["POOLS_TPV"] == 1, 1, 0).astype(int)

    if "EST_RENOUVELLER" not in dfx.columns:
        dfx["EST_RENOUVELLER"] = False
    dfx["EST_RENOUVELLER"] = dfx["EST_RENOUVELLER"].astype(bool)

    taxes = float(dfx["TAXES"].sum())
    ca = float(ca_series.sum())

    montant_pool = float(ca_series[dfx["POOLS_TPV"] == 1].sum())
    montant_hors_pool = float(ca_series[dfx["POOLS_TPV"] == 0].sum())

    souscriptions = int(len(dfx))
    renouvellements = int(dfx["EST_RENOUVELLER"].sum())
    nouvelles_affaires = int(souscriptions - renouvellements)

    nb_agents_actifs = int(dfx["NOM_AGENT"].dropna().nunique()) if "NOM_AGENT" in dfx.columns else 0
    nb_agents_fideles = int(agents_fideles(dfx))

    ratio_obj = (ca / float(objectif_annuel)) if objectif_annuel else 0.0
    pct_pool = (montant_pool / ca) if ca else 0.0
    pct_hors_pool = (montant_hors_pool / ca) if ca else 0.0
    pct_ren = (renouvellements / souscriptions) if souscriptions else 0.0
    pct_new = (nouvelles_affaires / souscriptions) if souscriptions else 0.0

    return {
        "taxes": taxes,
        "ca": ca,
        "montant_pool": montant_pool,
        "montant_hors_pool": montant_hors_pool,
        "souscriptions": souscriptions,
        "renouvellements": renouvellements,
        "nouvelles_affaires": nouvelles_affaires,
        "nb_agents_actifs": nb_agents_actifs,
        "nb_agents_fideles": nb_agents_fideles,
        "ratio_obj": ratio_obj,
        "pct_pool": pct_pool,
        "pct_hors_pool": pct_hors_pool,
        "pct_ren": pct_ren,
        "pct_new": pct_new,
    }


# ----------------- TABLE: REPARTITION PAR PRODUIT -----------------
def _policy_col(df: pd.DataFrame) -> str | None:
    for col in ["POLICYNO", "POLICY_NO", "POLICYNUMBER", "POLICY_NUMBER", "POLICY"]:
        if col in df.columns:
            return col
    return None


def tableau_produit(df_f: pd.DataFrame) -> pd.DataFrame:
    """
    Tableau: Répartition CA par produit
    Colonnes: PRODUIT | Nombre de contrats | % Contrats | Prime TTC (FCFA) | % Prime TTC
    """
    cols = ["PRODUIT", "Nombre de contrats", "% Contrats", "Prime TTC (FCFA)", "% Prime TTC"]
    if df_f is None or df_f.empty:
        return pd.DataFrame(columns=cols)
    if "PRODUIT" not in df_f.columns:
        raise KeyError("Colonne PRODUIT introuvable dans le fichier.")

    dfx = df_f.copy()
    dfx["PRIME_TTC"] = pd.to_numeric(dfx.get("PRIME_TTC", 0), errors="coerce").fillna(0)

    pol = _policy_col(dfx)
    if pol:
        nb = dfx.groupby("PRODUIT")[pol].nunique()
    else:
        nb = dfx.groupby("PRODUIT").size()

    prime = dfx.groupby("PRODUIT")["PRIME_TTC"].sum()

    t = pd.DataFrame({
        "PRODUIT": nb.index.astype(str),
        "Nombre de contrats": nb.values,
        "Prime TTC (FCFA)": prime.reindex(nb.index).values
    })

    total_nb = float(t["Nombre de contrats"].sum())
    total_prime = float(t["Prime TTC (FCFA)"].sum())

    t["% Contrats"] = (t["Nombre de contrats"] / total_nb).fillna(0) if total_nb else 0
    t["% Prime TTC"] = (t["Prime TTC (FCFA)"] / total_prime).fillna(0) if total_prime else 0

    t = t.sort_values("Prime TTC (FCFA)", ascending=False).reset_index(drop=True)

    total_row = pd.DataFrame([{
        "PRODUIT": "Total",
        "Nombre de contrats": int(total_nb),
        "% Contrats": 1.0 if total_nb else 0.0,
        "Prime TTC (FCFA)": int(round(total_prime)),
        "% Prime TTC": 1.0 if total_prime else 0.0,
    }])

    return pd.concat([t, total_row], ignore_index=True)


# ----------------- TABLE: CA MENSUEL -----------------
def monthly_ca_table(df_f: pd.DataFrame, objectif_annuel: float) -> pd.DataFrame:
    cols = ["Mois", "CA par Mois", "Objectif Mensuel", "Ratio Objectif Mensuel"]
    if df_f is None or df_f.empty or "DATE_CREATE" not in df_f.columns:
        return pd.DataFrame(columns=cols)

    dfx = df_f.dropna(subset=["DATE_CREATE"]).copy()
    if dfx.empty:
        return pd.DataFrame(columns=cols)

    dfx["PRIME_TTC"] = pd.to_numeric(dfx.get("PRIME_TTC", 0), errors="coerce").fillna(0)

    mois_fr = {
        1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
        5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
        9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
    }
    order_months = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]

    dfx["MONTH_NUM"] = pd.to_datetime(dfx["DATE_CREATE"]).dt.month
    dfx["Mois"] = dfx["MONTH_NUM"].map(mois_fr)

    g = dfx.groupby(["MONTH_NUM", "Mois"], as_index=False)["PRIME_TTC"].sum()
    g = g.sort_values("MONTH_NUM")

    obj_m = float(objectif_annuel) / 12.0 if objectif_annuel else 0.0
    g["CA par Mois"] = g["PRIME_TTC"].round(0).astype(int)
    g["Objectif Mensuel"] = int(round(obj_m))
    g["Ratio Objectif Mensuel"] = np.where(g["Objectif Mensuel"] > 0, g["CA par Mois"] / g["Objectif Mensuel"], 0.0)

    out = g[["Mois", "CA par Mois", "Objectif Mensuel", "Ratio Objectif Mensuel"]].copy()
    months_present = dfx["MONTH_NUM"].dropna().unique().tolist()
    months_present = [mois_fr[m] for m in sorted(months_present) if m in mois_fr]
    if months_present:
        order_months = months_present
    out["Mois"] = pd.Categorical(out["Mois"], categories=order_months, ordered=True)
    out = out.sort_values("Mois").reset_index(drop=True)

    total_ca = int(out["CA par Mois"].sum()) if not out.empty else 0
    total_obj = int(round(objectif_annuel)) if objectif_annuel else int(out["Objectif Mensuel"].sum())
    total_ratio = (total_ca / total_obj) if total_obj else 0.0

    total_label = f"Total ({order_months[0][:3]} → {order_months[-1][:3]})" if order_months else "Total"
    total_row = pd.DataFrame([{
        "Mois": total_label,
        "CA par Mois": total_ca,
        "Objectif Mensuel": total_obj,
        "Ratio Objectif Mensuel": total_ratio,
    }])

    return pd.concat([out, total_row], ignore_index=True)


# ----------------- TABLE: CA MOYEN JOURNALIER PAR JOUR -----------------
def weekly_avg_table(df_f: pd.DataFrame) -> pd.DataFrame:
    cols = ["Mois", "dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "Total"]
    if df_f is None or df_f.empty:
        return pd.DataFrame(columns=cols)
    if "DATE_CREATE" not in df_f.columns:
        raise KeyError("Colonne DATE_CREATE introuvable (requise pour weekly_avg_table).")

    df = df_f.copy().dropna(subset=["DATE_CREATE"]).copy()
    df["PRIME_TTC"] = pd.to_numeric(df.get("PRIME_TTC", 0), errors="coerce").fillna(0)

    df["DATE"] = pd.to_datetime(df["DATE_CREATE"], errors="coerce").dt.date
    df = df.dropna(subset=["DATE"]).copy()
    if df.empty:
        return pd.DataFrame(columns=cols)

    mois_fr = {
        1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
        5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
        9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
    }
    df["MONTH_NUM"] = pd.to_datetime(df["DATE_CREATE"]).dt.month
    df["Mois"] = df["MONTH_NUM"].map(mois_fr).fillna("")

    jours_fr = {0: "lundi", 1: "mardi", 2: "mercredi", 3: "jeudi", 4: "vendredi", 5: "samedi", 6: "dimanche"}
    df["WD"] = pd.to_datetime(df["DATE_CREATE"]).dt.weekday
    df["Jour"] = df["WD"].map(jours_fr)

    order_days = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"]
    order_months = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]

    daily = (
        df.groupby(["MONTH_NUM", "Mois", "Jour", "DATE"], as_index=False)["PRIME_TTC"]
        .sum()
        .rename(columns={"PRIME_TTC": "CA_jour"})
    )

    avg_mois_jour = (
        daily.groupby(["MONTH_NUM", "Mois", "Jour"], as_index=False)["CA_jour"]
        .mean()
        .rename(columns={"CA_jour": "CA_moyen"})
    )

    piv = avg_mois_jour.pivot_table(
        index=["MONTH_NUM", "Mois"],
        columns="Jour",
        values="CA_moyen",
        aggfunc="mean",
        fill_value=0
    ).reset_index()

    for d in order_days:
        if d not in piv.columns:
            piv[d] = 0

    daily_all = (
        df.groupby(["MONTH_NUM", "Mois", "DATE"], as_index=False)["PRIME_TTC"]
        .sum()
        .rename(columns={"PRIME_TTC": "CA_jour"})
    )
    total_mois = (
        daily_all.groupby(["MONTH_NUM", "Mois"], as_index=False)["CA_jour"]
        .mean()
        .rename(columns={"CA_jour": "Total"})
    )
    piv = piv.merge(total_mois, on=["MONTH_NUM", "Mois"], how="left")

    piv["Mois"] = pd.Categorical(piv["Mois"], categories=order_months, ordered=True)
    piv = piv.sort_values("Mois")

    for c in order_days + ["Total"]:
        piv[c] = pd.to_numeric(piv[c], errors="coerce").fillna(0).round(0).astype(int)

    daily2 = (
        df.groupby(["Jour", "DATE"], as_index=False)["PRIME_TTC"]
        .sum()
        .rename(columns={"PRIME_TTC": "CA_jour"})
    )
    avg_jour = daily2.groupby("Jour", as_index=False)["CA_jour"].mean().rename(columns={"CA_jour": "CA_moyen"})
    total_all = (
        df.groupby("DATE", as_index=False)["PRIME_TTC"]
        .sum()
        .rename(columns={"PRIME_TTC": "CA_jour"})
    )["CA_jour"].mean()

    total_row = {"Mois": "Total", "Total": int(round(float(total_all))) if pd.notna(total_all) else 0}
    for d in order_days:
        v = avg_jour.loc[avg_jour["Jour"] == d, "CA_moyen"]
        total_row[d] = int(round(float(v.iloc[0]))) if len(v) else 0

    piv = piv[["Mois"] + order_days + ["Total"]].copy()
    piv = pd.concat([piv, pd.DataFrame([total_row])[["Mois"] + order_days + ["Total"]]], ignore_index=True)
    return piv


# =========================
# BONUS: TOP 10 (Total / Pool / Hors Pool)
# =========================
def top_intermediaires(df_f: pd.DataFrame, n: int = 10, pool_value: int | None = None) -> pd.DataFrame:
    """
    Retourne les Top N INTERMEDIAIRE par CA (PRIME_TTC).
    pool_value:
      - None => total
      - 1 => POOL uniquement
      - 0 => Hors POOL uniquement
    """
    if df_f is None or df_f.empty:
        return pd.DataFrame(columns=["Rang", "INTERMEDIAIRE", "CA (FCFA)", "Souscriptions"])

    dfx = df_f.copy()
    if "INTERMEDIAIRE" not in dfx.columns:
        dfx["INTERMEDIAIRE"] = dfx.get("NOM_AGENT", pd.NA)

    dfx["PRIME_TTC"] = pd.to_numeric(dfx.get("PRIME_TTC", 0), errors="coerce").fillna(0)

    if "POOLS_TPV" not in dfx.columns:
        dfx["POOLS_TPV"] = 0
    dfx["POOLS_TPV"] = pd.to_numeric(dfx["POOLS_TPV"], errors="coerce").fillna(0).astype(int)
    dfx["POOLS_TPV"] = np.where(dfx["POOLS_TPV"] == 1, 1, 0).astype(int)

    if pool_value in (0, 1):
        dfx = dfx[dfx["POOLS_TPV"] == int(pool_value)].copy()

    g = (
        dfx.groupby("INTERMEDIAIRE", dropna=True)
        .agg(CA_FCFA=("PRIME_TTC", "sum"), Souscriptions=("PRIME_TTC", "size"))
        .reset_index()
    )
    g = g.sort_values("CA_FCFA", ascending=False).head(n).reset_index(drop=True)
    g.insert(0, "Rang", np.arange(1, len(g) + 1))
    g = g.rename(columns={"CA_FCFA": "CA (FCFA)"})
    return g


# =========================
# TAUX DE RENOUVELLEMENT (par mois)
# =========================
def _detect_client_col(df: pd.DataFrame) -> str | None:
    candidates = [
        "ID_CLIENT", "CLIENT_ID", "CUSTOMER_ID", "CUSTOMERID", "CODE_CLIENT",
        "NUM_CLIENT", "CLIENT_CODE", "CODEASSURE", "ID_ASSURE", "ASSURE_ID",
        "PHONE", "TEL", "TELEPHONE", "MOBILE",
        "EMAIL",
        "NOM_CLIENT", "CLIENT_NAME",
    ]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _clean_client_id(s: pd.Series) -> pd.Series:
    out = s.astype("string").str.strip()
    out = out.replace("", pd.NA)
    return out


def renewal_table_month(
    df_base: pd.DataFrame,
    year: int,
    month: int,
    segment: str = "Tous",
) -> tuple[str, pd.DataFrame]:
    cols_out = ["PRODUIT", "Total", "# New Business", "# Renouvellement", "% Renouvellement"]

    if df_base is None or df_base.empty:
        return "", pd.DataFrame(columns=cols_out)

    df = df_base.copy()

    if "DATE_CREATE" not in df.columns:
        return "", pd.DataFrame(columns=cols_out)

    df["DATE_CREATE"] = pd.to_datetime(df["DATE_CREATE"], errors="coerce")
    df = df.dropna(subset=["DATE_CREATE"]).copy()
    if df.empty:
        return "", pd.DataFrame(columns=cols_out)

    if "EST_RENOUVELLER" not in df.columns:
        df["EST_RENOUVELLER"] = False
    df["EST_RENOUVELLER"] = df["EST_RENOUVELLER"].astype(bool)

    if "PRODUIT" not in df.columns:
        df["PRODUIT"] = pd.NA

    if segment != "Tous" and "SEGMENT" in df.columns:
        df = df[df["SEGMENT"] == segment].copy()

    start = pd.Timestamp(year=year, month=month, day=1)
    end = (start + pd.DateOffset(months=1)) - pd.DateOffset(days=1)

    df_m = df[(df["DATE_CREATE"] >= start) & (df["DATE_CREATE"] <= end)].copy()
    if df_m.empty:
        title = f"{start.strftime('%B').upper()}_{year}_{segment.upper()}"
        return title, pd.DataFrame(columns=cols_out)

    df_m["PRODUIT"] = df_m["PRODUIT"].astype("string").str.strip().fillna("INCONNU")

    g_total = df_m.groupby("PRODUIT", as_index=False).size().rename(columns={"size": "Total"})
    g_ren = (
        df_m.groupby("PRODUIT", as_index=False)["EST_RENOUVELLER"]
        .sum()
        .rename(columns={"EST_RENOUVELLER": "# Renouvellement"})
    )

    out = g_total.merge(g_ren, on="PRODUIT", how="left")
    out["# Renouvellement"] = out["# Renouvellement"].fillna(0).astype(int)

    client_col = _detect_client_col(df)
    if client_col is not None:
        df_all = df.dropna(subset=[client_col]).copy()
        df_all[client_col] = _clean_client_id(df_all[client_col])
        df_all = df_all.dropna(subset=[client_col]).copy()

        if df_all.empty:
            out["# New Business"] = (out["Total"] - out["# Renouvellement"]).clip(lower=0).astype(int)
        else:
            first_seen = df_all.groupby(client_col)["DATE_CREATE"].min()
            new_clients_ids = first_seen[(first_seen >= start) & (first_seen <= end)].index

            df_m2 = df_m.copy()
            df_m2[client_col] = _clean_client_id(
                df_m2.get(client_col, pd.Series([pd.NA] * len(df_m2), index=df_m2.index))
            )
            df_m2 = df_m2.dropna(subset=[client_col]).copy()

            if df_m2.empty:
                out["# New Business"] = (out["Total"] - out["# Renouvellement"]).clip(lower=0).astype(int)
            else:
                df_new = df_m2[df_m2[client_col].isin(new_clients_ids)].copy()
                g_new = df_new.groupby("PRODUIT", as_index=False).size().rename(columns={"size": "# New Business"})
                out = out.merge(g_new, on="PRODUIT", how="left")
                out["# New Business"] = out["# New Business"].fillna(0).astype(int)
    else:
        out["# New Business"] = (out["Total"] - out["# Renouvellement"]).clip(lower=0).astype(int)

    out["% Renouvellement"] = np.where(out["Total"] > 0, out["# Renouvellement"] / out["Total"], 0.0)
    out = out.sort_values("Total", ascending=False).reset_index(drop=True)

    tot_total = int(out["Total"].sum())
    tot_new = int(out["# New Business"].sum())
    tot_ren = int(out["# Renouvellement"].sum())
    tot_pct = (tot_ren / tot_total) if tot_total else 0.0

    grand = pd.DataFrame([{
        "PRODUIT": "Grand Total",
        "Total": tot_total,
        "# New Business": tot_new,
        "# Renouvellement": tot_ren,
        "% Renouvellement": tot_pct,
    }])

    out = pd.concat([out, grand], ignore_index=True)

    mois_fr = {
        1: "JANVIER", 2: "FEVRIER", 3: "MARS", 4: "AVRIL",
        5: "MAI", 6: "JUIN", 7: "JUILLET", 8: "AOUT",
        9: "SEPTEMBRE", 10: "OCTOBRE", 11: "NOVEMBRE", 12: "DECEMBRE"
    }
    seg_label = "TOUS" if segment == "Tous" else (segment.upper() + "S")
    title = f"{mois_fr.get(month, str(month))}_{year}_{seg_label}"

    return title, out[cols_out]
