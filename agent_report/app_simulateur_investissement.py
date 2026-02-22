import streamlit as st

# --------------------------------------------------------
# CONFIGURATION DE LA PAGE
# --------------------------------------------------------
st.set_page_config(
    page_title="Simulateur d'investissement FCP vs DAT",
    layout="centered"
)

# Petite feuille de style pour une interface propre
st.markdown(
    """
    <style>
    body {
        font-family: "Raleway", sans-serif;
    }
    .main-title {
        text-align: center;
        color: #003366;
        font-weight: 800;
    }
    .sub-title {
        text-align: center;
        color: #555555;
        font-size: 0.95rem;
    }
    .bloc {
        padding: 1.3rem 1rem;
        border-radius: 0.9rem;
        border: 1px solid #e0e0e0;
        background-color: #ffffff;
        margin-bottom: 1rem;
    }
    .bloc-header {
        font-weight: 700;
        color: #003366;
        margin-bottom: 0.8rem;
        font-size: 1.05rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------
# TITRE
# --------------------------------------------------------
st.markdown('<h1 class="main-title">Simulateur d\'investissement</h1>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">Comparez la rentabilité d\'un placement en FCP et d\'un DAT à partir de vos paramètres.</p>',
    unsafe_allow_html=True,
)

st.write("")

# --------------------------------------------------------
# ZONE 1 - INFORMATIONS GÉNÉRALES
# --------------------------------------------------------
with st.container():
    st.markdown('<div class="bloc"><div class="bloc-header">Informations générales</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        devise = st.selectbox("Devise", ["F CFA", "EUR", "USD"], index=0)
        nom = st.text_input("Nom et prénoms")
        tel = st.text_input("Numéro de téléphone")
    with col2:
        email = st.text_input("Email")
        type_client = st.selectbox("Type de client", ["PARTICULIER", "ENTREPRISE", "AUTRE"], index=0)
        fcp = st.selectbox("Fonds Commun de Placement (FCP)", ["AURORE OPPORTUNITES", "Autre FCP"])
    st.markdown("</div>", unsafe_allow_html=True)

# --------------------------------------------------------
# ZONE 2 - PARAMÈTRES D'INVESTISSEMENT
# --------------------------------------------------------
with st.container():
    st.markdown('<div class="bloc"><div class="bloc-header">Paramètres d\'investissement</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        duree_annees = st.number_input("Durée de l'investissement (années)", min_value=1, max_value=50, value=5)
        capital_initial = st.number_input("Capital initial", min_value=0.0, value=500000.0, step=50000.0, format="%.0f")
        cotisation_mensuelle = st.number_input("Cotisation mensuelle", min_value=0.0, value=200000.0, step=50000.0, format="%.0f")
    with col2:
        contribution_employeur = st.number_input("Contribution mensuelle de l'employeur", min_value=0.0, value=0.0, step=50000.0, format="%.0f")
        taux_fcp_annuel = st.number_input("Taux annuel moyen du FCP (%)", min_value=0.0, max_value=100.0, value=13.0, step=0.25)
        taux_dat_annuel = st.number_input("Taux annuel du DAT (%)", min_value=0.0, max_value=100.0, value=3.75, step=0.25)
        taux_irc = st.number_input("Taux de l'IRC sur DAT (%)", min_value=0.0, max_value=100.0, value=5.0, step=0.5)

    st.markdown("</div>", unsafe_allow_html=True)

# --------------------------------------------------------
# ZONE 3 - CALCUL & RÉSULTATS
# --------------------------------------------------------
lancer = st.button("Lancer la simulation")

if lancer:
    # ---- Préparation des données ----
    n_mois = int(duree_annees * 12)
    versement_mensuel_total = cotisation_mensuelle + contribution_employeur
    total_verse = capital_initial + versement_mensuel_total * n_mois

    # Conversion des taux
    r_fcp_annuel = taux_fcp_annuel / 100.0
    r_dat_annuel = taux_dat_annuel / 100.0
    r_irc = taux_irc / 100.0

    # Taux mensuels équivalents
    r_fcp_mensuel = (1 + r_fcp_annuel) ** (1 / 12) - 1 if r_fcp_annuel > 0 else 0
    r_dat_mensuel = (1 + r_dat_annuel) ** (1 / 12) - 1 if r_dat_annuel > 0 else 0

    # ---- FCP ----
    if r_fcp_mensuel > 0:
        vf_fcp_capital = capital_initial * (1 + r_fcp_mensuel) ** n_mois
        vf_fcp_versements = versement_mensuel_total * (((1 + r_fcp_mensuel) ** n_mois - 1) / r_fcp_mensuel)
    else:
        vf_fcp_capital = capital_initial
        vf_fcp_versements = versement_mensuel_total * n_mois

    vf_fcp_total = vf_fcp_capital + vf_fcp_versements
    interets_fcp = vf_fcp_total - total_verse

    # ---- DAT ----
    if r_dat_mensuel > 0:
        vf_dat_capital = capital_initial * (1 + r_dat_mensuel) ** n_mois
        vf_dat_versements = versement_mensuel_total * (((1 + r_dat_mensuel) ** n_mois - 1) / r_dat_mensuel)
    else:
        vf_dat_capital = capital_initial
        vf_dat_versements = versement_mensuel_total * n_mois

    vf_dat_total_brut = vf_dat_capital + vf_dat_versements
    interets_dat_brut = vf_dat_total_brut - total_verse
    montant_irc = interets_dat_brut * r_irc
    vf_dat_total_net = vf_dat_total_brut - montant_irc
    interets_dat_net = interets_dat_brut - montant_irc

    # ---- Comparaison ----
    gain_vs_dat = vf_fcp_total - vf_dat_total_net

    # ----------------------------------------------------
    # AFFICHAGE DES RÉSULTATS
    # ----------------------------------------------------
    st.write("")
    st.markdown('<div class="bloc"><div class="bloc-header">Résultat de votre simulation</div>', unsafe_allow_html=True)

    st.markdown(
        f"En fin de période, vous aurez **cotisé un montant total de : {total_verse:,.0f} {devise}**",
        unsafe_allow_html=True,
    )

    colA, colB = st.columns(2)
    with colA:
        st.subheader("Investissement dans le FCP")
        st.metric("Montant final FCP", f"{vf_fcp_total:,.0f} {devise}")
        st.metric("Intérêts générés (FCP)", f"{interets_fcp:,.0f} {devise}")
    with colB:
        st.subheader("Investissement dans un DAT")
        st.metric("Montant final DAT (net IRC)", f"{vf_dat_total_net:,.0f} {devise}")
        st.metric("Intérêts nets (DAT)", f"{interets_dat_net:,.0f} {devise}")

    st.divider()
    st.subheader("Comparaison")
    st.metric("Surplus du FCP par rapport au DAT", f"{gain_vs_dat:,.0f} {devise}")

    st.markdown("</div>", unsafe_allow_html=True)

    # ----------------------------------------------------
    # TEXTE SYNTHÉTIQUE (style écran final)
    # ----------------------------------------------------
    st.write("")
    st.markdown('<div class="bloc">', unsafe_allow_html=True)
    st.markdown("### Merci d'avoir utilisé notre simulateur.", unsafe_allow_html=True)

    txt = (
        f"Si vous réalisez un apport initial de **{capital_initial:,.0f} {devise}** "
        f"et des cotisations mensuelles de **{cotisation_mensuelle:,.0f} {devise}**"
    )
    if contribution_employeur > 0:
        txt += f", avec une contribution mensuelle de votre employeur de **{contribution_employeur:,.0f} {devise}**"
    txt += (
        f" dans le FCP **{fcp}** sur une période de **{duree_annees:.0f} an(s)**, "
        f"la valeur de votre investissement s'élèvera à **{vf_fcp_total:,.0f} {devise}**.\n\n"
        f"Tandis que ce même investissement dans un **DAT** vaudra **{vf_dat_total_net:,.0f} {devise}** (net d'IRC).\n\n"
        f"Le FCP génère donc un surplus de **{gain_vs_dat:,.0f} {devise}** par rapport au DAT."
    )

    st.markdown(txt)
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info("Renseignez les paramètres puis cliquez sur **Lancer la simulation**.")
