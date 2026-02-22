# SUIVI_AUTO/main.py
from __future__ import annotations

import streamlit as st
from tasks.suivi_hebdomadaire_auto import run as run_suivi_hebdo

st.set_page_config(page_title="SUIVI AUTO", layout="wide")

st.sidebar.title("📌 TÂCHES")
task = st.sidebar.radio(
    "Choisir une tâche",
    ["SUIVI HEBDOMMADIARE AUTO", "TAUX DE RENOUVELLEMENT (Bientôt)"],
    index=0
)

if task == "SUIVI HEBDOMMADIARE AUTO":
    run_suivi_hebdo()
else:
    st.info("Bientôt 🙂")
