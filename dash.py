from __future__ import annotations

import pandas as pd
import streamlit as st
import plotly.express as px


def _find_col(df: pd.DataFrame, keys: list[str]) -> str | None:
    cols = set(df.columns)
    for k in keys:
        if k in cols:
            return k
    return None


def _norm_key(s: str) -> str:
    return (
        str(s)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("'", "")
    )


def _find_col_fuzzy(df: pd.DataFrame, keys: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    norm_map = {_norm_key(c): c for c in df.columns}
    for k in keys:
        kk = _norm_key(k)
        if kk in norm_map:
            return norm_map[kk]
    return None


def _to_num(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("")
    s = s.str.replace("\u00A0", " ", regex=False)
    s = s.str.replace("\u202F", " ", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(r"[^0-9,.\-]", "", regex=True)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def run() -> None:
    st.header("📊 Dashboard Production")

    df = st.session_state.get("df_base")
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        st.info("Aucune donnée détectée. Charge un fichier CSV/Excel pour afficher le dashboard.")
        up = st.file_uploader("Données dashboard (CSV/Excel)", type=["csv", "xlsx", "xls"], key="dash_upload")
        if up is None:
            return
        try:
            if up.name.lower().endswith(".csv"):
                try:
                    df = pd.read_csv(up, sep=None, engine="python")
                except Exception:
                    up.seek(0)
                    df = pd.read_csv(up, sep=";", encoding="utf-8")
            else:
                df = pd.read_excel(up)
            st.session_state["df_base"] = df
            st.success("Fichier chargé. Le dashboard est maintenant alimenté.")
        except Exception as exc:
            st.error(f"Erreur de lecture fichier: {exc}")
            return

    date_col = _find_col_fuzzy(
        df,
        [
            "DATE_CREATE",
            "DATE CREATION",
            "DATE_CREATION",
            "DATE_EMISSION",
            "DATE EMISSION",
            "ISSUE_DATE",
            "ISSUE DATE",
            "Effective Date",
            "DATE",
        ],
    )
    if not date_col:
        st.error("Colonne de date introuvable (ex: DATE_CREATE, DATE_EMISSION, ISSUE_DATE).")
        return

    dfx = df.copy()
    dfx[date_col] = pd.to_datetime(dfx[date_col], errors="coerce")
    dfx = dfx.dropna(subset=[date_col]).copy()
    if dfx.empty:
        st.warning(f"Aucune date valide dans {date_col}.")
        return

    min_d = dfx[date_col].min().date()
    max_d = dfx[date_col].max().date()
    period = st.sidebar.date_input(
        f"Période Dashboard ({date_col})",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
        key="dash_period",
    )
    if isinstance(period, tuple) and len(period) == 2:
        start_d, end_d = period
    else:
        start_d, end_d = min_d, max_d

    dfx = dfx[(dfx[date_col].dt.date >= start_d) & (dfx[date_col].dt.date <= end_d)].copy()
    if dfx.empty:
        st.warning("Aucune donnée sur la période sélectionnée.")
        return

    col_prime = _find_col(dfx, ["PRIME_NETTE", "PRIME_NETTE_"])
    col_acc = _find_col(dfx, ["ACCESSOIRES", "ACCESSOIRE"])
    col_lob = _find_col(dfx, ["LINE_OF_BUSINESS", "BRANCHE"])
    col_policy = _find_col(dfx, ["N_POLICE", "N_POLICE_", "POLICY_NO", "N_POLICE_INTERMEDIAIRE"])
    col_ren = _find_col(dfx, ["EST_RENOUVELLER", "EST_RENOUVELER"])

    ca = (
        (_to_num(dfx[col_prime]) if col_prime else 0)
        + (_to_num(dfx[col_acc]) if col_acc else 0)
    )
    dfx["_CA"] = ca if isinstance(ca, pd.Series) else 0.0

    nb_contrats = int(len(dfx)) if not col_policy else int(dfx[col_policy].astype("string").replace("", pd.NA).dropna().nunique())
    ca_total = float(dfx["_CA"].sum())

    if col_ren:
        ren_s = dfx[col_ren].astype("string").str.upper().str.strip()
        nb_ren = int(ren_s.isin(["TRUE", "VRAI", "1", "OUI", "YES"]).sum())
    else:
        nb_ren = 0
    tx_ren = (nb_ren / nb_contrats) if nb_contrats else 0.0

    c1, c2, c3 = st.columns(3)
    c1.metric("Nombre de contrats", f"{nb_contrats:,}".replace(",", " "))
    c2.metric("Chiffre d'affaires (Prime nette + Accessoires)", f"{ca_total:,.0f}".replace(",", " "))
    c3.metric("Renouvellements", f"{nb_ren:,}".replace(",", " "), f"{tx_ren*100:.1f}%")

    if col_lob:
        lob = (
            dfx.groupby(col_lob, dropna=False)
            .agg(**{"Contrats": ("_CA", "size"), "CA": ("_CA", "sum")})
            .reset_index()
            .rename(columns={col_lob: "Line of Business"})
            .sort_values("CA", ascending=False)
        )
        st.subheader("Line of Business")
        st.dataframe(
            lob.style.format({"Contrats": lambda x: f"{int(x):,}".replace(",", " "), "CA": lambda x: f"{x:,.0f}".replace(",", " ")}),
            use_container_width=True,
            hide_index=True,
        )
        f1, f2 = st.columns(2)
        with f1:
            fig = px.bar(lob, y="Line of Business", x="CA", orientation="h", title="CA par Line of Business")
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig, use_container_width=True)
        with f2:
            fig2 = px.pie(lob, values="CA", names="Line of Business", title="Répartition CA par Line of Business")
            fig2.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig2, use_container_width=True)

    month = (
        dfx.assign(MOIS=dfx[date_col].dt.to_period("M").astype(str))
        .groupby("MOIS", dropna=False)
        .agg(**{"Contrats": ("_CA", "size"), "CA": ("_CA", "sum")})
        .reset_index()
    )
    st.subheader("Évolution mensuelle")
    fig_m = px.line(month, x="MOIS", y="CA", markers=True, title="Évolution CA mensuel")
    fig_m.update_layout(height=320, margin=dict(l=10, r=10, t=40, b=10))
    st.plotly_chart(fig_m, use_container_width=True)


if __name__ == "__main__":
    run()
