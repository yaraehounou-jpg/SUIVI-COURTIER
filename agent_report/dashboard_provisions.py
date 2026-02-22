import streamlit as st
import pandas as pd

# -----------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------
st.set_page_config(page_title="Calcul Automatique PREC / PPNA", layout="wide")
st.markdown("<h1 style='text-align:center;color:#003366;'>📊 Dashboard de Calcul des Provisions</h1>", unsafe_allow_html=True)

file = st.file_uploader("📂 Charger un fichier Excel", type=["xlsx"])
date_eval = st.date_input("📌 Date d'évaluation", value=pd.to_datetime("2025-12-31"))

if file:
    df = pd.read_excel(file)

    # -----------------------------------------------------------
    # NORMALISATION DES COLONNES
    # -----------------------------------------------------------
    df.columns = (
        df.columns.str.strip()
                  .str.replace("\u00A0", " ")
                  .str.lower()
                  .str.replace(" ", "_")
                  .str.replace("__", "_")
    )

    st.write("🔎 Colonnes détectées :", df.columns.tolist())

    # Colonnes requises
    required_cols = ["date_d'expiration", "date_d'effet", "prime_nette", "branche"]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"❌ Colonne manquante : {col}")
            st.stop()

    # -----------------------------------------------------------
    # CONVERSION DES DATES
    # -----------------------------------------------------------
    df["date_d'effet"] = pd.to_datetime(df["date_d'effet"])
    df["date_d'expiration"] = pd.to_datetime(df["date_d'expiration"])
    date_eval = pd.to_datetime(date_eval)

    # -----------------------------------------------------------
    # CALCUL PREC
    # -----------------------------------------------------------
    df["nombre_de_jours"] = (df["date_d'expiration"] - df["date_d'effet"]).dt.days
    df["jours_restants"] = (df["date_d'expiration"] - date_eval).dt.days
    df["jours_restants"] = df["jours_restants"].clip(lower=0)
    df["PREC"] = df["prime_nette"] * df["jours_restants"] / df["nombre_de_jours"]

    st.success("🎉 Calcul PREC effectué avec succès !")
    st.metric("PREC Totale du portefeuille", f"{df['PREC'].sum():,.0f} FCFA")

    # -----------------------------------------------------------
    # TABLEAU PREC PAR BRANCHE — FORMAT CIMA
    # -----------------------------------------------------------
    st.subheader("📊 PROVISION POUR RISQUES EN COURS – Tableau Consolidé")

    # Année de PREC = année d’expiration
    df["annee_prec"] = df["date_d'expiration"].dt.year

    # Colonnes FIXES dans l’ordre CIMA
    colonnes_fixes = [2022, 2023, 2024]

    # Pivot PREC par branche et année
    table_prec = df.pivot_table(
        values="PREC",
        index="branche",
        columns="annee_prec",
        aggfunc="sum",
        fill_value=0
    )

    # Forcer colonnes manquantes
    for c in colonnes_fixes:
        if c not in table_prec.columns:
            table_prec[c] = 0

    # Réordonner les colonnes
    table_prec = table_prec[[2022, 2023, 2024]]

    # -----------------------------------------------------------
    # Colonne : Octobre 2025
    # -----------------------------------------------------------
    df_oct = df[
        (df["date_d'expiration"].dt.year == 2025) &
        (df["date_d'expiration"].dt.month == 10)
    ]

    prec_oct = df_oct.groupby("branche")["PREC"].sum()
    table_prec["Octobre 2025"] = prec_oct
    table_prec["Octobre 2025"] = table_prec["Octobre 2025"].fillna(0)

    # -----------------------------------------------------------
    # TOTAL
    # -----------------------------------------------------------
    table_prec["TOTAL"] = table_prec.sum(axis=1)

    # -----------------------------------------------------------
    # GRAND TOTAL
    # -----------------------------------------------------------
    table_prec.loc["Grand Total"] = table_prec.sum()

    # -----------------------------------------------------------
    # STYLE DU TABLEAU
    # -----------------------------------------------------------
    def style_total(row):
        if row.name == "Grand Total":
            return ["background-color:#AA0000;color:white;font-weight:bold;"] * len(row)
        return [""] * len(row)

    styled_prec = (
        table_prec.style
            .format("{:,.0f}")
            .apply(style_total, axis=1)
            .set_table_styles([
                {"selector": "th.col_heading", 
                 "props": "background-color:#C54E00; color:white; font-weight:bold; border:1px solid black;"},
                {"selector": "th.row_heading", 
                 "props": "background-color:#333333; color:white; font-weight:bold; border:1px solid black;"},
            ])
    )

    # -----------------------------------------------------------
    # AFFICHAGE STREAMLIT
    # -----------------------------------------------------------
    st.write(styled_prec.to_html(), unsafe_allow_html=True)

    # -----------------------------------------------------------
    # EXPORT EXCEL
    # -----------------------------------------------------------
    @st.cache_data
    def export_excel(df, table_prec):
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="DETAIL_PREC")
            table_prec.to_excel(writer, sheet_name="PREC_PAR_BRANCHE")
        return output.getvalue()

    st.download_button(
        "📥 Télécharger le fichier Excel (PREC + Tableau Consolidé)",
        data=export_excel(df, table_prec),
        file_name="provisions_PREC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("🔔 Veuillez charger un fichier Excel pour commencer.")
