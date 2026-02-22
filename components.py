import streamlit as st

def load_css(path="assets/styles.css"):
    with open(path, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

def banner(title: str):
    st.markdown(f'<div class="banner">{title}</div>', unsafe_allow_html=True)

def kpi_card(value: str, label: str, orange=False, value_color="#d6793f", pct: str | None = None):
    cls = "kpi-card kpi-orange" if orange else "kpi-card"
    pct_html = f"<div class='kpi-pct'>{pct}</div>" if pct is not None else ""
    html = f"""
    <div class="{cls}">
      <p class="kpi-value" style="color:{value_color};">{value}</p>
      <div class="kpi-label">{label}</div>
      {pct_html}
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def ratio_card(ratio_pct: str):
    html = f"""
    <div class="kpi-ratio">
      <p class="ratio-value">{ratio_pct}</p>
      <div class="ratio-label">Ratio Objectif Annuel</div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)
