from __future__ import annotations

from datetime import date
import hmac
from io import BytesIO
from pathlib import Path

from PIL import Image
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
import matplotlib.pyplot as plt


def _to_datetime(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors="coerce")
    if dt.isna().any():
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
    return pd.to_numeric(s, errors="coerce").fillna(0)


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


def _card(label: str, value: str, bg: str, icon: str) -> str:
    return f"""
    <div class="pc-card" style="background:{bg};">
      <div>
        <div class="pc-card-title">{label}</div>
        <div class="pc-card-value">{value}</div>
      </div>
      <div class="pc-card-icon">{icon}</div>
    </div>
    """


def _kpi_block(title: str, value: str, accent: str) -> str:
    return f"""
    <div style="background:#ffffff;border:1px solid #eceff3;border-top:4px solid {accent};border-radius:16px;padding:18px 20px;box-shadow:0 8px 18px rgba(0,0,0,0.06);min-height:170px;">
      <div style="font-weight:700;margin-bottom:12px;font-size:16px;color:#111827;">{title}</div>
      <div style="font-size:28px;font-weight:800;color:{accent};">{value}</div>
    </div>
    """


def _norm_id(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("")
    s = s.str.replace("\u00A0", " ", regex=False).str.strip()
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(r"\.0+$", "", regex=True)
    return s.replace("", pd.NA)


def _policy_key(series: pd.Series) -> pd.Series:
    s = _norm_id(series).astype("string")
    s = s.str.split("_").str[0]
    s = s.str.upper()
    return s.replace("", pd.NA)


def _fmt_money(v: float) -> str:
    return f"{float(v):,.0f} FCFA".replace(",", " ")


def _fmt_int(v: float | int) -> str:
    return f"{int(v):,}".replace(",", " ")


def _build_dashboard_pdf(
    title: str,
    meta: list[tuple[str, str]],
    sections: list[tuple[str, list[tuple[str, str]]]],
    payment_paid_donut_png: bytes | None = None,
    payment_remaining_donut_png: bytes | None = None,
    recap_lines: list[str] | None = None,
    logo_png: bytes | None = None,
    logo_align: str = "right",
) -> bytes:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=landscape(A4))
    w, h = landscape(A4)

    def _new_page():
        c.setFillColorRGB(0.95, 0.95, 0.95)
        c.rect(0, 0, w, h, stroke=0, fill=1)
        if logo_png:
            try:
                logo_w = 46 * mm
                logo_h = 14 * mm
                if logo_align == "left":
                    logo_x = 14 * mm
                else:
                    logo_x = w - 14 * mm - logo_w
                logo_y = h - 4 * mm - logo_h
                c.drawImage(
                    ImageReader(BytesIO(logo_png)),
                    logo_x,
                    logo_y,
                    width=logo_w,
                    height=logo_h,
                    mask=None,
                    preserveAspectRatio=True,
                    anchor="ne",
                )
            except Exception:
                pass
        # Footer decoration (orange arcs, bottom-left)
        c.setStrokeColor(colors.HexColor("#ff6f00"))
        c.setLineCap(1)
        for r_mm in [30, 42, 54, 66]:
            r = r_mm * mm
            c.setLineWidth(8)
            c.arc(-r, -r, r, r, 0, 90)
        # Footer legal text
        c.setFillColor(colors.HexColor("#111111"))
        c.setFont("Helvetica", 10)
        c.drawCentredString(
            w / 2,
            16 * mm,
            "Leadway Assurance IARD - S.A au capital de 5.000.000.000 entierement libere - RCCM CI-ABJ-2020-11167, CC : 2033877P",
        )
        c.drawCentredString(
            w / 2,
            11 * mm,
            "Siege : Abidjan Cocody Angre 7eme tranche pres du centre commercial Tera et CIRAD - Tel: (225) 20013100/ 20013101 - E-mail: rc@leadway.com - Website: www.ci.leadway.com",
        )

    def _draw_card(
        x: float,
        y: float,
        cw: float,
        ch: float,
        label: str,
        value: str,
        accent=colors.HexColor("#2e6bd9"),
        donut_png: bytes | None = None,
        payment_style: bool = False,
    ):
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#d8dee7"))
        c.roundRect(x, y - ch, cw, ch, 4 * mm, stroke=1, fill=1)
        c.setFillColor(accent)
        c.rect(x, y - 2, cw, 2, stroke=0, fill=1)
        c.setFillColor(colors.HexColor("#1f2937"))
        lab_txt = str(label)
        lab_len = len(lab_txt)
        lab_size = 10
        if lab_len > 34:
            lab_size = 9
        if lab_len > 48:
            lab_size = 8
        if lab_len > 62:
            lab_txt = lab_txt[:59] + "..."
        c.setFont("Helvetica-Bold", lab_size)
        c.drawString(x + 4 * mm, y - 8 * mm, lab_txt)
        c.setFillColor(accent)
        txt = str(value)
        val_len = len(txt)
        val_size = 16
        if val_len > 24:
            val_size = 14
        if val_len > 32:
            val_size = 13
        if val_len > 46:
            val_size = 11
        if val_len > 56:
            val_size = 9
        if val_len > 70:
            val_size = 8
        c.setFont("Helvetica-Bold", val_size)
        value_y = y - (15 * mm if payment_style else 15 * mm)
        c.drawString(x + 4 * mm, value_y, txt)
        if payment_style:
            c.setFillColor(colors.HexColor("#6b7280"))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(x + 4 * mm, y - 22 * mm, "from 0 FCFA")
            badge_w = 22 * mm
            badge_h = 8 * mm
            bx = x + 4 * mm
            by = y - ch + 5 * mm
            c.setFillColor(colors.HexColor("#d1fae5"))
            c.roundRect(bx, by, badge_w, badge_h, 4 * mm, stroke=0, fill=1)
            c.setFillColor(colors.HexColor("#047857"))
            c.setFont("Helvetica-Bold", 12)
            c.drawString(bx + 3 * mm, by + 2.1 * mm, "↑ 0.0%")
        if donut_png:
            d = 24 * mm
            dx = x + cw - d - 4 * mm
            dy = y - ch + 8 * mm
            c.drawImage(ImageReader(BytesIO(donut_png)), dx, dy, width=d, height=d, mask="auto")

    _new_page()
    banner_h = 10 * mm
    banner_top_gap = 20 * mm
    c.setFillColor(colors.HexColor("#d79d86"))
    c.roundRect(12 * mm, h - banner_h - banner_top_gap, w - 24 * mm, banner_h, 4 * mm, stroke=0, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(w / 2, h - banner_h - banner_top_gap + 3.1 * mm, title)

    y = h - banner_h - banner_top_gap - 8 * mm
    c.setFillColor(colors.HexColor("#000000"))
    c.setFont("Helvetica-Bold", 10)
    meta_map = {str(k): str(v) for k, v in meta}
    courtier_val = meta_map.get("Courtier", "-")
    periode_val = meta_map.get("Période Analysée", meta_map.get("Période", "-"))
    c.drawString(14 * mm, y, f"Courtier: {courtier_val}")
    c.drawRightString(w - 14 * mm, y, f"Période Analysée : {periode_val}")
    y -= 7 * mm

    accents = [colors.HexColor("#6d5efc"), colors.HexColor("#ff7a59"), colors.HexColor("#22c55e"), colors.HexColor("#0ea5e9")]
    gap = 6 * mm
    cols = 3
    cw = (w - 2 * 14 * mm - (cols - 1) * gap) / cols
    ch_default = 24 * mm
    ch_payment = 34 * mm

    for sec_title, rows in sections:
        if y < 35 * mm:
            c.showPage()
            _new_page()
            y = h - 18 * mm
        c.setFillColor(colors.HexColor("#111827"))
        c.setFont("Helvetica-Bold", 13)
        c.drawString(14 * mm, y, sec_title)
        y -= 4 * mm
        i = 0
        while i < len(rows):
            section_is_payment = sec_title.strip().lower() == "paiements"
            ch = ch_payment if section_is_payment else ch_default
            if y - ch < 10 * mm:
                c.showPage()
                _new_page()
                y = h - 18 * mm
            row = rows[i:i + cols]
            for j, (k, v) in enumerate(row):
                x = 14 * mm + j * (cw + gap)
                donut = None
                if section_is_payment:
                    if i + j == 1:
                        donut = payment_paid_donut_png
                    elif i + j == 2:
                        donut = payment_remaining_donut_png
                _draw_card(
                    x,
                    y,
                    cw,
                    ch,
                    k,
                    v,
                    accent=accents[(i + j) % len(accents)],
                    donut_png=donut,
                    payment_style=section_is_payment,
                )
            y -= ch + 4 * mm
            i += cols
        y -= 2 * mm
    if recap_lines:
        c.showPage()
        _new_page()
        y2 = h - 22 * mm
        c.setFillColor(colors.HexColor("#111827"))
        c.setFont("Helvetica-Bold", 18)
        c.drawString(14 * mm, y2, "REMARQUE")
        y2 -= 10 * mm
        c.setFont("Helvetica", 11)
        for line in recap_lines:
            if y2 < 14 * mm:
                c.showPage()
                _new_page()
                y2 = h - 18 * mm
                c.setFillColor(colors.HexColor("#111827"))
                c.setFont("Helvetica", 11)
            c.drawString(16 * mm, y2, f"- {line}")
            y2 -= 7 * mm
    c.save()
    buf.seek(0)
    return buf.getvalue()


def _build_ratio_donut_png(ratio: float) -> bytes:
    r = max(0.0, min(1.0, float(ratio)))
    fig, ax = plt.subplots(figsize=(2.4, 2.4), dpi=180)
    vals = [r, 1.0 - r]
    colors_list = ["#22c55e", "#d1d5db"]
    ax.pie(vals, colors=colors_list, startangle=90, counterclock=False, wedgeprops={"width": 0.35, "edgecolor": "white"})
    pct = r * 100.0
    ax.text(0, 0, f"{pct:.1f}%", ha="center", va="center", fontsize=12, fontweight="bold", color="#111827")
    ax.set(aspect="equal")
    ax.axis("off")
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _read_first_logo_bytes() -> bytes | None:
    for p in [
        Path("agent_report/assets/logo_leadway.jpg"),
        Path("assets/logo_leadway.jpg"),
        Path("logo_leadway.jpg"),
        Path("assets/logo_assur_defender.png"),
        Path("assets/logo_assur_defender.jpg"),
    ]:
        if p.exists():
            try:
                return p.read_bytes()
            except Exception:
                continue
    return None


def _remove_dark_background(logo_bytes: bytes) -> bytes:
    try:
        img = Image.open(BytesIO(logo_bytes)).convert("RGBA")
        data = img.getdata()
        cleaned = []
        for r, g, b, a in data:
            if r < 35 and g < 35 and b < 35:
                cleaned.append((r, g, b, 0))
            else:
                cleaned.append((r, g, b, a))
        img.putdata(cleaned)
        out = BytesIO()
        img.save(out, format="PNG")
        out.seek(0)
        return out.getvalue()
    except Exception:
        return logo_bytes


def _render_login_gate() -> bool:
    if st.session_state.get("auth_ok", False):
        return True

    st.markdown(
        """
        <style>
        .stApp { background: #efefef; }
        .auth-topbar {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            height: 58px;
            background: #ef6c00;
            color: #ffffff;
            display: flex;
            align-items: center;
            padding: 0 14px;
            font-size: 18px;
            font-weight: 500;
            z-index: 999;
        }
        .auth-wrap {
            max-width: 520px;
            margin: 0 auto;
            padding-top: 92px;
        }
        .auth-sub {
            font-size: 48px;
            font-weight: 500;
            margin: 6px 0 16px 0;
            color: #1f2937;
            text-align: left;
        }
        .auth-footer {
            text-align: center;
            color: #6b7280;
            margin-top: 56px;
            font-size: 16px;
        }
        div[data-testid="stTextInput"] input {
            height: 56px;
            border-radius: 10px;
            border: 1px solid #d1d5db;
            background: #ffffff;
            font-size: 18px;
        }
        div[data-testid="stCheckbox"] label {
            font-size: 18px;
            color: #1f2937;
        }
        div[data-testid="stFormSubmitButton"] button {
            height: 56px;
            border-radius: 12px;
            border: none;
            background: #f39a73;
            color: #ffffff;
            font-size: 22px;
            font-weight: 500;
        }
        div[data-testid="stFormSubmitButton"] button:hover {
            background: #ef8c63;
            color: #ffffff;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div class='auth-topbar'>Leadway Central</div>", unsafe_allow_html=True)
    left, center, right = st.columns([1, 1.4, 1])
    with center:
        st.markdown("<div class='auth-wrap'>", unsafe_allow_html=True)
        st.markdown("<div class='auth-sub'>Connexion</div>", unsafe_allow_html=True)
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Nom d'utilisateur", placeholder="Nom d'utilisateur", label_visibility="collapsed")
            password = st.text_input("Mot de passe", type="password", placeholder="Mot de passe", label_visibility="collapsed")
            st.checkbox("Se souvenir de moi", key="remember_me")
            submitted = st.form_submit_button("Connexion", use_container_width=True)
        st.markdown("<div style='text-align:center;margin-top:12px;font-size:22px;'><a href='mailto:rc@leadway.com'>Mot de passe oublié ?</a></div>", unsafe_allow_html=True)
        st.markdown("<div class='auth-footer'>© Leadway Assurance Côte d'ivoire</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    if not submitted:
        return False

    def _normalize_role(v: str) -> str:
        vv = str(v).strip().lower()
        if vv in {"admin", "viewer"}:
            return vv
        return "viewer"

    def _load_auth_users() -> dict[str, dict[str, str]]:
        users: dict[str, dict[str, str]] = {}
        try:
            raw = st.secrets.get("auth_users", {})
        except Exception:
            raw = {}
        # Optional nested form: [auth.users.<username>]
        if not raw:
            try:
                auth_root = st.secrets.get("auth", {})
                raw = auth_root.get("users", {}) if isinstance(auth_root, dict) else {}
            except Exception:
                raw = {}
        if not isinstance(raw, dict):
            return users
        for uname, cfg in raw.items():
            if isinstance(cfg, dict):
                upass = str(cfg.get("password", "")).strip()
                urole = _normalize_role(str(cfg.get("role", "viewer")))
            else:
                upass = str(cfg).strip()
                urole = "viewer"
            uname_clean = str(uname).strip()
            if uname_clean and upass:
                users[uname_clean] = {"password": upass, "role": urole}
        return users

    username_clean = username.strip()
    auth_users = _load_auth_users()
    if not auth_users:
        st.error("Aucun utilisateur configure. Ajoute auth_users dans Streamlit secrets.")
        return False
    current = auth_users.get(username_clean)
    valid = bool(current) and hmac.compare_digest(str(current.get("password", "")), password)

    if not valid:
        st.error("Nom d'utilisateur ou mot de passe invalide.")
        return False

    st.session_state["auth_ok"] = True
    st.session_state["auth_user"] = username_clean
    st.session_state["auth_role"] = str(current.get("role", "viewer")) if current else "viewer"
    st.rerun()
    return True


st.set_page_config(page_title="RAPPORT DE VOTRE ACTIVITE", layout="wide")

if not _render_login_gate():
    st.stop()

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&display=swap');
    .pc-title {
        font-family: 'Montserrat', sans-serif;
        font-weight: 800;
        font-size: 32px;
        margin-bottom: 6px;
    }
    .pc-sub {
        color: #5f6b7a;
        margin-bottom: 18px;
        font-size: 14px;
    }
    .pc-card {
        padding: 22px 24px;
        border-radius: 18px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 8px 22px rgba(0,0,0,0.12);
        min-height: 140px;
        font-family: 'Montserrat', sans-serif;
    }
    .pc-card-title {
        font-size: 20px;
        font-weight: 700;
        margin-bottom: 24px;
    }
    .pc-card-value {
        font-size: 26px;
        font-weight: 800;
    }
    .pc-card-icon {
        width: 52px;
        height: 52px;
        border-radius: 26px;
        background: #0f1115;
        color: #fff;
        display:flex;
        align-items:center;
        justify-content:center;
        font-size: 22px;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='pc-title'>RAPPORT DE VOTRE ACTIVITE</div>", unsafe_allow_html=True)
st.markdown("<div class='pc-sub'>Dashboard courtiers — cotations, polices et commissions.</div>", unsafe_allow_html=True)

is_admin = st.session_state.get("auth_role") == "admin"
quotes_file = None
policies_file = None
comm_file = None
pay_file = None
platform_file = None
pdf_logo_file = None
pdf_logo_mode = "Auto"
if is_admin:
    st.sidebar.header("📂 Fichiers")
    quotes_file = st.sidebar.file_uploader("Cotations (Excel/CSV)", type=["xlsx", "xls", "csv"], key="pc_quotes")
    policies_file = st.sidebar.file_uploader("Polices (Excel/CSV)", type=["xlsx", "xls", "csv"], key="pc_policies")
    comm_file = st.sidebar.file_uploader("Commissions (Excel/CSV)", type=["xlsx", "xls", "csv"], key="pc_comm")
    pay_file = st.sidebar.file_uploader("Paiements (Excel/CSV)", type=["xlsx", "xls", "csv"], key="pc_pay")
    platform_file = st.sidebar.file_uploader("Data Plateforme (Excel/CSV)", type=["xlsx", "xls", "csv"], key="pc_platform")
    st.sidebar.header("🖼️ Logo PDF")
    pdf_logo_mode = st.sidebar.selectbox("Choix du logo", ["Auto", "Leadway", "Assur Defender", "Upload"], index=0)
    if pdf_logo_mode == "Upload":
        pdf_logo_file = st.sidebar.file_uploader("Logo PDF (PNG/JPG)", type=["png", "jpg", "jpeg"], key="pc_pdf_logo")
else:
    st.sidebar.caption("Mode utilisateur: fichiers et paramètres admin masqués.")

def _safe_read_excel(uploaded, label: str) -> pd.DataFrame:
    if not uploaded:
        return pd.DataFrame()
    try:
        if uploaded.name.lower().endswith(".csv"):
            try:
                return pd.read_csv(uploaded, sep=None, engine="python")
            except Exception:
                return pd.read_csv(uploaded, sep=";", encoding="utf-8")
        return pd.read_excel(uploaded)
    except Exception as exc:
        st.sidebar.error(f"Erreur lecture {label}: {exc}")
        return pd.DataFrame()

data_store_dir = Path(".dashboard_store")
data_store_dir.mkdir(parents=True, exist_ok=True)

def _cache_path(key: str) -> Path:
    return data_store_dir / f"{key}.pkl"

def _load_cached_df(key: str) -> pd.DataFrame:
    p = _cache_path(key)
    if p.exists():
        try:
            return pd.read_pickle(p)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def _resolve_df(uploaded, label: str, key: str) -> pd.DataFrame:
    if is_admin and uploaded is not None:
        fresh = _safe_read_excel(uploaded, label)
        if not fresh.empty:
            try:
                fresh.to_pickle(_cache_path(key))
            except Exception:
                pass
            return fresh
    return _load_cached_df(key)

quotes_df = _resolve_df(quotes_file, "Cotations", "quotes")
policies_df = _resolve_df(policies_file, "Polices", "policies")
comm_df = _resolve_df(comm_file, "Commissions", "commissions")
pay_df = _resolve_df(pay_file, "Paiements", "paiements")
platform_df = _resolve_df(platform_file, "Data Plateforme", "plateforme")

if quotes_df.empty and policies_df.empty and comm_df.empty and pay_df.empty and platform_df.empty:
    if is_admin:
        st.info("Charge les fichiers pour afficher les KPI.")
    else:
        st.warning("Aucune donnée publiée par l'admin. Merci de contacter l'administrateur.")
    st.stop()

st.sidebar.header("🎛️ Filtres")
date_mode = "Polices — Issue Date"

def _clean_broker_series(s: pd.Series) -> pd.Series:
    ss = s.astype("string").fillna("").str.replace("\u00A0", " ", regex=False).str.strip()
    ss = ss.replace({"": pd.NA, "NAN": pd.NA, "NONE": pd.NA, "NULL": pd.NA}, regex=False)
    ss = ss.str.replace(r"\s+", " ", regex=True)
    return ss

broker_values = set()
q_broker_col = _find_col(quotes_df, ["COURTIER", "Courtier"])
p_broker_col = _find_col(policies_df, ["BROKER", "Broker"])
c_broker_col = _find_col(comm_df, ["COURTIER", "Courtier"])
p_pay_broker_col = _find_col(pay_df, ["COURTIER", "Courtier", "BROKER", "Broker"])
p_platform_broker_col = _find_col(platform_df, ["COURTIER", "Courtier", "BROKER", "Broker"])
if q_broker_col:
    broker_values.update(_clean_broker_series(quotes_df[q_broker_col]).dropna().unique().tolist())
if p_broker_col:
    broker_values.update(_clean_broker_series(policies_df[p_broker_col]).dropna().unique().tolist())
if c_broker_col:
    broker_values.update(_clean_broker_series(comm_df[c_broker_col]).dropna().unique().tolist())
if p_pay_broker_col:
    broker_values.update(_clean_broker_series(pay_df[p_pay_broker_col]).dropna().unique().tolist())
if p_platform_broker_col:
    broker_values.update(_clean_broker_series(platform_df[p_platform_broker_col]).dropna().unique().tolist())

broker_names = sorted({str(x).strip() for x in broker_values if str(x).strip()})
broker_names = ["Tous courtiers"] + broker_names if broker_names else ["Tous courtiers"]
if is_admin:
    broker_scope = st.sidebar.selectbox("Courtier", ["Tous", "Sans courtier", "Courtier"], index=0)
    broker_choice = st.sidebar.selectbox("Detail Courtier", broker_names, index=0 if broker_names else 0)
else:
    broker_scope = "Courtier"
    broker_choice = st.sidebar.selectbox("Detail Courtier", broker_names, index=0 if broker_names else 0)

if broker_scope == "Tous":
    broker_filter_label = "Tous les courtiers"
elif broker_scope == "Sans courtier":
    broker_filter_label = "Sans courtier"
elif broker_choice == "Tous courtiers":
    broker_filter_label = "Tous les courtiers renseignés"
else:
    broker_filter_label = str(broker_choice)
st.caption(f"Filtre courtier actif : {broker_filter_label}")

if is_admin and not policies_df.empty:
    st.sidebar.markdown("#### Export Polices (CSV)")
    st.sidebar.download_button(
        "Télécharger Polices (CSV)",
        data=policies_df.to_csv(index=False).encode("utf-8"),
        file_name="polices_completes.csv",
        mime="text/csv",
    )

if is_admin and not pay_df.empty:
    st.sidebar.markdown("#### Export Paiements (CSV)")
    st.sidebar.download_button(
        "Télécharger Paiements (CSV)",
        data=pay_df.to_csv(index=False).encode("utf-8"),
        file_name="paiements_complets.csv",
        mime="text/csv",
    )

def _date_col_from_mode() -> tuple[pd.DataFrame, str | None]:
    return policies_df, _find_col(policies_df, ["ISSUE_DATE", "ISSUE DATE", "Effective Date"])

df_ref, date_col = _date_col_from_mode()
if is_admin and df_ref is not None and date_col:
    df_ref = df_ref.copy()
    df_ref[date_col] = _to_datetime(df_ref[date_col])
    df_ref = df_ref.dropna(subset=[date_col]).copy()
    if not df_ref.empty:
        min_d = df_ref[date_col].min().date()
        max_d = df_ref[date_col].max().date()
        date_range = st.sidebar.date_input("Période", value=(min_d, max_d), min_value=min_d, max_value=max_d)
    else:
        date_range = None
else:
    date_range = None

def _filter_by_range(df: pd.DataFrame, col: str | None, dr: tuple[date, date] | None) -> pd.DataFrame:
    if df is None or df.empty or not col or not dr:
        return df
    s = _to_datetime(df[col])
    start_d, end_d = dr
    return df[(s.dt.date >= start_d) & (s.dt.date <= end_d)].copy()

quotes_f = _filter_by_range(quotes_df, _find_col(quotes_df, ["DATE_CREATION", "DATE CREATION"]), date_range)
policies_f = _filter_by_range(policies_df, _find_col(policies_df, ["ISSUE_DATE", "ISSUE DATE", "Effective Date"]), date_range)
comm_f = _filter_by_range(comm_df, _find_col(comm_df, ["DATE", "Date paiement", "DATE PAIEMENT"]), date_range)
pay_f = _filter_by_range(pay_df, _find_col(pay_df, ["DATE", "Date paiement", "DATE PAIEMENT", "PAYMENT_DATE", "PAYMENT DATE"]), date_range)
platform_f = platform_df.copy()
quotes_date_f = quotes_f.copy()
policies_date_f = policies_f.copy()

cost_col = _find_col(quotes_df, ["COUT_TOTAL", "COÛT_TOTAL", "COUT TOTAL", "COÛT TOTAL"])
if cost_col:
    for _df in [quotes_df, quotes_f]:
        if _df is None or _df.empty or cost_col not in _df.columns:
            continue
        mask_calc = _df[cost_col].astype("string").str.contains("calcul en cours", case=False, na=False)
        _df.drop(_df[mask_calc].index, inplace=True)

if broker_scope != "Tous":
    if broker_scope == "Sans courtier":
        if q_broker_col and not quotes_f.empty:
            quotes_f = quotes_f[_clean_broker_series(quotes_f[q_broker_col]).isna()].copy()
        if p_broker_col and not policies_f.empty:
            policies_f = policies_f[_clean_broker_series(policies_f[p_broker_col]).isna()].copy()
        if c_broker_col and not comm_f.empty:
            comm_f = comm_f[_clean_broker_series(comm_f[c_broker_col]).isna()].copy()
        if p_pay_broker_col and not pay_f.empty:
            pay_f = pay_f[_clean_broker_series(pay_f[p_pay_broker_col]).isna()].copy()
        if p_platform_broker_col and not platform_f.empty:
            platform_f = platform_f[_clean_broker_series(platform_f[p_platform_broker_col]).isna()].copy()
    elif broker_scope == "Courtier":
        if broker_choice == "Tous courtiers":
            if q_broker_col and not quotes_f.empty:
                quotes_f = quotes_f[_clean_broker_series(quotes_f[q_broker_col]).notna()].copy()
            if p_broker_col and not policies_f.empty:
                policies_f = policies_f[_clean_broker_series(policies_f[p_broker_col]).notna()].copy()
            if c_broker_col and not comm_f.empty:
                comm_f = comm_f[_clean_broker_series(comm_f[c_broker_col]).notna()].copy()
            if p_pay_broker_col and not pay_f.empty:
                pay_f = pay_f[_clean_broker_series(pay_f[p_pay_broker_col]).notna()].copy()
            if p_platform_broker_col and not platform_f.empty:
                platform_f = platform_f[_clean_broker_series(platform_f[p_platform_broker_col]).notna()].copy()
        else:
            if q_broker_col and not quotes_f.empty:
                quotes_f = quotes_f[_clean_broker_series(quotes_f[q_broker_col]) == broker_choice].copy()
            if p_broker_col and not policies_f.empty:
                policies_f = policies_f[_clean_broker_series(policies_f[p_broker_col]) == broker_choice].copy()
            if c_broker_col and not comm_f.empty:
                comm_f = comm_f[_clean_broker_series(comm_f[c_broker_col]) == broker_choice].copy()
            if p_pay_broker_col and not pay_f.empty:
                pay_f = pay_f[_clean_broker_series(pay_f[p_pay_broker_col]) == broker_choice].copy()
            if p_platform_broker_col and not platform_f.empty:
                platform_f = platform_f[_clean_broker_series(platform_f[p_platform_broker_col]) == broker_choice].copy()

# KPI: Polices créées (unique Policy No)
policy_col = _find_col(policies_f, ["POLICY_NO", "POLICY NO", "Policy No"])
pol_quote_col = _find_col(policies_f, ["BROKER_QUOTE_NO", "BROKER QUOTE NO", "Broker Quote No"])
polices_creees = int(policies_f[policy_col].astype("string").str.strip().replace("", pd.NA).dropna().nunique()) if policy_col else 0

# KPI: Chiffre d'affaires (Amount)
amount_col = _find_col(policies_f, ["AMOUNT", "Amount"])
ca_total = float(_to_num(policies_f[amount_col]).sum()) if amount_col else 0.0

# KPI: CA / jour (moy.)
jours_periode = 0
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    start_d, end_d = date_range
    jours_periode = int((end_d - start_d).days) + 1
if jours_periode <= 0:
    pol_date_for_avg = _find_col(policies_f, ["ISSUE_DATE", "ISSUE DATE", "Effective Date"])
    if pol_date_for_avg and not policies_f.empty:
        dts = _to_datetime(policies_f[pol_date_for_avg]).dropna()
        jours_periode = int(dts.dt.date.nunique()) if not dts.empty else 0
ca_jour_moy = (ca_total / jours_periode) if jours_periode else 0.0

def _pct_change(cur: float, prev: float) -> float:
    if prev == 0:
        return 0.0
    return (cur - prev) / prev

def _delta_badge(pct: float) -> str:
    if pct >= 0:
        return f"<span style='background:#dff7ea;color:#1b7f4b;padding:6px 12px;border-radius:999px;font-weight:700;display:inline-block;'>↑ {pct*100:.1f}%</span>"
    return f"<span style='background:#ffe4e4;color:#b00020;padding:6px 12px;border-radius:999px;font-weight:700;display:inline-block;'>↓ {abs(pct)*100:.1f}%</span>"

def _comm_block(title: str, cur: float, prev: float, pct: float, accent: str) -> str:
    cur_txt = f"{cur:,.0f} FCFA".replace(",", " ")
    prev_txt = f"{prev:,.0f} FCFA".replace(",", " ")
    return f"""
    <div style="background:#ffffff;border:1px solid #eceff3;border-top:4px solid {accent};border-radius:16px;padding:18px 20px;box-shadow:0 8px 18px rgba(0,0,0,0.06);">
      <div style="font-weight:700;margin-bottom:12px;font-size:16px;color:#111827;">{title}</div>
      <div style="font-size:28px;font-weight:800;color:{accent};">{cur_txt}</div>
      <div style="color:#6b7280;font-weight:600;margin-top:4px;">from {prev_txt}</div>
      <div style="margin-top:10px;">{_delta_badge(pct)}</div>
    </div>
    """


def _comm_block_with_ratio(title: str, cur: float, prev: float, pct: float, accent: str, ratio: float) -> str:
    cur_txt = f"{cur:,.0f} FCFA".replace(",", " ")
    prev_txt = f"{prev:,.0f} FCFA".replace(",", " ")
    rr = max(0.0, min(1.0, float(ratio)))
    angle = rr * 360.0
    rr_txt = f"{rr*100:.1f}%"
    return f"""
    <div style="background:#ffffff;border:1px solid #eceff3;border-top:4px solid {accent};border-radius:16px;padding:18px 20px;box-shadow:0 8px 18px rgba(0,0,0,0.06);">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:14px;">
        <div style="min-width:0;">
          <div style="font-weight:700;margin-bottom:12px;font-size:16px;color:#111827;">{title}</div>
          <div style="font-size:28px;font-weight:800;color:{accent};">{cur_txt}</div>
          <div style="color:#6b7280;font-weight:600;margin-top:4px;">from {prev_txt}</div>
          <div style="margin-top:10px;">{_delta_badge(pct)}</div>
        </div>
        <div style="width:92px;height:92px;border-radius:50%;background:conic-gradient(#22c55e {angle}deg, #e5e7eb 0deg);display:flex;align-items:center;justify-content:center;flex-shrink:0;">
          <div style="width:64px;height:64px;border-radius:50%;background:#ffffff;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:800;color:#111827;">{rr_txt}</div>
        </div>
      </div>
    </div>
    """

def _build_broker_perf(q_src: pd.DataFrame, p_src: pd.DataFrame) -> pd.DataFrame:
    tbl = pd.DataFrame(columns=["Courtier", "Cotations", "Convertis", "Montant", "Taux"])
    if not q_broker_col or q_src.empty:
        return tbl
    quote_col = _find_col(q_src, ["N°_COTATION", "NO_COTATION", "N_COTATION", "N° COTATION", "N°COTATION", "NUMERO_COTATION"])
    if not quote_col:
        return tbl
    q_base = q_src[[q_broker_col, quote_col]].copy()
    q_base[q_broker_col] = _clean_broker_series(q_base[q_broker_col])
    q_base[quote_col] = q_base[quote_col].astype("string").str.strip().replace("", pd.NA)
    q_base = q_base.dropna(subset=[q_broker_col, quote_col]).copy()
    cot_counts = q_base.groupby(q_broker_col)[quote_col].nunique().rename("Cotations")

    conv_counts = pd.Series(dtype=int, name="Convertis")
    conv_amounts = pd.Series(dtype=float, name="Montant")
    if not p_src.empty and p_broker_col:
        pol_id_col = _find_col(p_src, ["POLICY_NO", "POLICY NO", "Policy No"])
        pol_amt_col = _find_col(p_src, ["AMOUNT", "Amount"])
        cols = [p_broker_col]
        if pol_id_col:
            cols.append(pol_id_col)
        if pol_amt_col:
            cols.append(pol_amt_col)
        p_base = p_src[cols].copy()
        p_base[p_broker_col] = _clean_broker_series(p_base[p_broker_col])
        p_base = p_base.dropna(subset=[p_broker_col]).copy()
        if pol_id_col and pol_id_col in p_base.columns:
            p_base[pol_id_col] = _norm_id(p_base[pol_id_col])
            p_base = p_base.dropna(subset=[pol_id_col]).copy()
            conv_counts = p_base.groupby(p_broker_col)[pol_id_col].nunique().rename("Convertis")
        else:
            conv_counts = p_base.groupby(p_broker_col).size().rename("Convertis")
        if pol_amt_col and pol_amt_col in p_base.columns:
            conv_amounts = p_base.groupby(p_broker_col)[pol_amt_col].apply(lambda s: float(_to_num(s).sum())).rename("Montant")

    tbl = cot_counts.to_frame().merge(conv_counts, left_index=True, right_index=True, how="left")
    if not conv_amounts.empty:
        tbl = tbl.merge(conv_amounts, left_index=True, right_index=True, how="left")
    else:
        tbl["Montant"] = 0.0
    tbl = tbl.fillna(0)
    tbl["Taux"] = tbl.apply(lambda r: (r["Convertis"] / r["Cotations"]) if r["Cotations"] else 0.0, axis=1)
    tbl = tbl.reset_index().rename(columns={q_broker_col: "Courtier"})
    tbl = tbl[tbl["Courtier"].astype("string").str.strip().ne("")].copy()
    return tbl.sort_values("Montant", ascending=False).reset_index(drop=True)

broker_perf = _build_broker_perf(quotes_date_f, policies_date_f)

st.markdown("### Mes Cotations")
q_date_col = _find_col(quotes_df, ["DATE_CREATION", "DATE CREATION"])
q_id_col = _find_col(quotes_df, ["N°_COTATION", "NO_COTATION", "N_COTATION", "N° COTATION", "N°COTATION", "NUMERO_COTATION"])
q_base = quotes_df.copy()
if broker_scope != "Tous":
    if broker_scope == "Sans courtier" and q_broker_col:
        q_base = q_base[_clean_broker_series(q_base[q_broker_col]).isna()].copy()
    elif broker_scope == "Courtier" and broker_choice and q_broker_col:
        if broker_choice == "Tous courtiers":
            q_base = q_base[_clean_broker_series(q_base[q_broker_col]).notna()].copy()
        else:
            q_base = q_base[_clean_broker_series(q_base[q_broker_col]) == broker_choice].copy()
if q_date_col:
    q_base[q_date_col] = _to_datetime(q_base[q_date_col])
    q_base = q_base.dropna(subset=[q_date_col]).copy()
    if date_mode.startswith("Cotations") and date_range:
        start_d, end_d = date_range
        q_base = q_base[(q_base[q_date_col].dt.date >= start_d) & (q_base[q_date_col].dt.date <= end_d)].copy()

year_count = 0
month_count = 0
week_count = 0
if not q_base.empty and q_date_col:
    ref_date = q_base[q_date_col].max().date()
    year_mask = q_base[q_date_col].dt.year == ref_date.year
    month_mask = (q_base[q_date_col].dt.year == ref_date.year) & (q_base[q_date_col].dt.month == ref_date.month)
    week_start = ref_date - pd.Timedelta(days=ref_date.weekday())
    week_end = week_start + pd.Timedelta(days=6)
    week_mask = (q_base[q_date_col].dt.date >= week_start) & (q_base[q_date_col].dt.date <= week_end)
    if q_id_col:
        year_count = int(q_base.loc[year_mask, q_id_col].astype("string").replace("", pd.NA).dropna().nunique())
        month_count = int(q_base.loc[month_mask, q_id_col].astype("string").replace("", pd.NA).dropna().nunique())
        week_count = int(q_base.loc[week_mask, q_id_col].astype("string").replace("", pd.NA).dropna().nunique())
    else:
        year_count = int(q_base.loc[year_mask].shape[0])
        month_count = int(q_base.loc[month_mask].shape[0])
        week_count = int(q_base.loc[week_mask].shape[0])
elif not q_base.empty and not q_date_col:
    if q_id_col:
        year_count = int(q_base[q_id_col].astype("string").replace("", pd.NA).dropna().nunique())
        month_count = year_count
        week_count = year_count
    else:
        year_count = int(q_base.shape[0])
        month_count = year_count
        week_count = year_count

avg_cotations = 0.0
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    start_d, end_d = date_range
    nb_days = int((end_d - start_d).days) + 1
    if nb_days > 0:
        avg_cotations = year_count / nb_days
else:
    avg_cotations = float(month_count) if month_count else 0.0

quote_amount_col = _find_col(quotes_f, ["COUT_TOTAL", "COÛT_TOTAL", "COUT TOTAL", "COÛT TOTAL"])
quotes_for_amount = quotes_f if (quote_amount_col and not quotes_f.empty) else quotes_df
quote_amount_col = quote_amount_col or _find_col(quotes_for_amount, ["COUT_TOTAL", "COÛT_TOTAL", "COUT TOTAL", "COÛT TOTAL"])
quote_amount = float(_to_num(quotes_for_amount[quote_amount_col]).sum()) if quote_amount_col and not quotes_for_amount.empty else 0.0
cotations_total = int(quotes_f[q_id_col].astype("string").replace("", pd.NA).dropna().nunique()) if q_id_col and not quotes_f.empty else 0
conversion_rate = (polices_creees / cotations_total) if cotations_total else 0.0

mc1, mc2, mc3 = st.columns(3)
with mc1:
    st.markdown(_kpi_block("Cotations", f"{year_count:,}".replace(",", " "), "#6d5efc"), unsafe_allow_html=True)
with mc2:
    st.markdown(_kpi_block("Montant cotations", f"{quote_amount:,.0f} FCFA".replace(",", " "), "#ff7a59"), unsafe_allow_html=True)
with mc3:
    st.markdown(_kpi_block("Cotations moyenne", f"{int(round(avg_cotations)):,}".replace(",", " "), "#22c55e"), unsafe_allow_html=True)

st.markdown("### Police")
part_realisee = (ca_total / quote_amount) if quote_amount else 0.0
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(_kpi_block("Polices créées", f"{polices_creees:,}".replace(",", " "), "#6d5efc"), unsafe_allow_html=True)
with c2:
    st.markdown(_kpi_block("Chiffres d'affaires", f"{ca_total:,.0f} FCFA".replace(",", " "), "#ff7a59"), unsafe_allow_html=True)
with c3:
    st.markdown(_kpi_block("Part réalisée", f"{part_realisee*100:.1f}%".replace(".", ","), "#22c55e"), unsafe_allow_html=True)

st.markdown("### Conversion")
cc1, cc2 = st.columns(2)
with cc1:
    st.markdown(_kpi_block("Conversion", f"{conversion_rate*100:.1f}%".replace(".", ","), "#6d5efc"), unsafe_allow_html=True)
with cc2:
    st.markdown(_kpi_block("Part réalisée", f"{part_realisee*100:.1f}%".replace(".", ","), "#ff7a59"), unsafe_allow_html=True)

def _sum_amount_in_range(df: pd.DataFrame, col_date: str | None, col_val: str | None, start_d: date, end_d: date) -> float:
    if df.empty or not col_date or not col_val:
        return 0.0
    dfx = df.copy()
    dfx[col_date] = _to_datetime(dfx[col_date])
    dfx = dfx.dropna(subset=[col_date]).copy()
    mask = (dfx[col_date].dt.date >= start_d) & (dfx[col_date].dt.date <= end_d)
    return float(_to_num(dfx.loc[mask, col_val]).sum())


def _mean_amount_in_range(df: pd.DataFrame, col_date: str | None, col_val: str | None, start_d: date, end_d: date) -> float:
    if df.empty or not col_date or not col_val:
        return 0.0
    dfx = df.copy()
    dfx[col_date] = _to_datetime(dfx[col_date])
    dfx = dfx.dropna(subset=[col_date]).copy()
    mask = (dfx[col_date].dt.date >= start_d) & (dfx[col_date].dt.date <= end_d)
    vals = _to_num(dfx.loc[mask, col_val])
    return float(vals.mean()) if not vals.empty else 0.0

ref_d = date.today()
comm_date_col = _find_col(comm_f, ["DATE", "Date paiement", "DATE PAIEMENT"])
comm_val_col = _find_col(comm_f, ["COMM._A_PAYER", "COMM_A_PAYER", "COMM. A PAYER", "COMM A PAYER"])
if comm_val_col is None:
    comm_val_col = _find_col(comm_f, ["COMM.", "COMM"])
if comm_date_col and not comm_f.empty:
    ref_d = _to_datetime(comm_f[comm_date_col]).dropna().max().date()

year_start = date(ref_d.year, 1, 1)
year_end = date(ref_d.year, 12, 31)
month_start = ref_d.replace(day=1)
month_end = (month_start + pd.offsets.MonthEnd(1)).date()
day_start = ref_d
day_end = ref_d

prev_year_start = date(ref_d.year - 1, 1, 1)
prev_year_end = date(ref_d.year - 1, 12, 31)
prev_month_end = month_start - pd.Timedelta(days=1)
prev_month_start = prev_month_end.replace(day=1)
prev_day = ref_d - pd.Timedelta(days=1)

year_val = _sum_amount_in_range(comm_f, comm_date_col, comm_val_col, year_start, year_end)
year_prev = _sum_amount_in_range(comm_f, comm_date_col, comm_val_col, prev_year_start, prev_year_end)
month_val = _sum_amount_in_range(comm_f, comm_date_col, comm_val_col, month_start, month_end)
month_prev = _sum_amount_in_range(comm_f, comm_date_col, comm_val_col, prev_month_start, prev_month_end)
day_val = _sum_amount_in_range(comm_f, comm_date_col, comm_val_col, day_start, day_end)
day_prev = _sum_amount_in_range(comm_f, comm_date_col, comm_val_col, prev_day, prev_day)
mean_val = _mean_amount_in_range(comm_f, comm_date_col, comm_val_col, month_start, month_end)
mean_prev = _mean_amount_in_range(comm_f, comm_date_col, comm_val_col, prev_month_start, prev_month_end)

st.markdown("### Paiements")
pay_date_col = _find_col(pay_df, ["DATE", "Date paiement", "DATE PAIEMENT", "PAYMENT_DATE", "PAYMENT DATE"])
pay_val_col = _find_col(
    pay_df,
    [
        "MONTANT",
        "Montant",
        "MONTANT_PAIEMENT",
        "MONTANT PAIEMENT",
        "AMOUNT",
        "NET_A_PAYER",
        "NET A PAYER",
        "PAID_AMOUNT",
        "PAID AMOUNT",
    ],
)
if pay_val_col is None:
    pay_val_col = _find_col(pay_df, ["COMM._A_PAYER", "COMM_A_PAYER", "COMM. A PAYER", "COMM A PAYER", "COMM.", "COMM"])
pay_paid_col = _find_col(
    pay_df,
    [
        "MONTANT_PAYE",
        "MONTANT PAYE",
        "Montant Payé",
        "Montant Paye",
        "PAID_AMOUNT",
        "PAID AMOUNT",
    ],
)
pay_remaining_col = _find_col(
    pay_df,
    [
        "MONTANT_RESTANT",
        "MONTANT RESTANT",
        "Montant Restant",
        "REMAINING_AMOUNT",
        "REMAINING AMOUNT",
    ],
)
pay_due_col = _find_col(
    pay_df,
    [
        "DATE_BUTOIR",
        "DATE BUTOIR",
        "DATE ECHEANCE",
        "DATE_ECHEANCE",
        "ECHEANCE",
        "Echéance",
        "DUE_DATE",
        "DUE DATE",
    ],
)

pay_ref_d = date.today()
if pay_date_col and not pay_f.empty:
    pay_ref_d = _to_datetime(pay_f[pay_date_col]).dropna().max().date()

pay_year_start = date(pay_ref_d.year, 1, 1)
pay_year_end = date(pay_ref_d.year, 12, 31)
pay_month_start = pay_ref_d.replace(day=1)
pay_month_end = (pay_month_start + pd.offsets.MonthEnd(1)).date()
pay_day_start = pay_ref_d
pay_day_end = pay_ref_d

pay_prev_year_start = date(pay_ref_d.year - 1, 1, 1)
pay_prev_year_end = date(pay_ref_d.year - 1, 12, 31)
pay_prev_month_end = pay_month_start - pd.Timedelta(days=1)
pay_prev_month_start = pay_prev_month_end.replace(day=1)
pay_prev_day = pay_ref_d - pd.Timedelta(days=1)

pay_year_val = _sum_amount_in_range(pay_f, pay_date_col, pay_val_col, pay_year_start, pay_year_end)
pay_year_prev = _sum_amount_in_range(pay_f, pay_date_col, pay_val_col, pay_prev_year_start, pay_prev_year_end)
pay_month_val = _sum_amount_in_range(pay_f, pay_date_col, pay_val_col, pay_month_start, pay_month_end)
pay_month_prev = _sum_amount_in_range(pay_f, pay_date_col, pay_val_col, pay_prev_month_start, pay_prev_month_end)
pay_day_val = _sum_amount_in_range(pay_f, pay_date_col, pay_val_col, pay_day_start, pay_day_end)
pay_day_prev = _sum_amount_in_range(pay_f, pay_date_col, pay_val_col, pay_prev_day, pay_prev_day)
pay_paid_month_val = _sum_amount_in_range(pay_f, pay_date_col, pay_paid_col, pay_month_start, pay_month_end)
pay_paid_month_prev = _sum_amount_in_range(pay_f, pay_date_col, pay_paid_col, pay_prev_month_start, pay_prev_month_end)
pay_remaining_val = float(_to_num(pay_f[pay_remaining_col]).sum()) if (pay_remaining_col and not pay_f.empty) else 0.0
pay_remaining_prev = 0.0
pay_paid_total = float(_to_num(pay_f[pay_paid_col]).sum()) if (pay_paid_col and not pay_f.empty) else 0.0
pay_remaining_total = float(_to_num(pay_f[pay_remaining_col]).sum()) if (pay_remaining_col and not pay_f.empty) else 0.0
pay_due_total = float(_to_num(pay_f[pay_val_col]).sum()) if (pay_val_col and not pay_f.empty) else 0.0
pay_ratio = (pay_paid_total / (pay_paid_total + pay_remaining_total)) if (pay_paid_total + pay_remaining_total) else 0.0
pay_remaining_ratio = (pay_remaining_total / (pay_paid_total + pay_remaining_total)) if (pay_paid_total + pay_remaining_total) else 0.0

if pay_file and (not pay_date_col or not pay_val_col):
    st.warning("Fichier Paiements chargé, mais colonne date ou montant introuvable. Vérifie les en-têtes.")

pj1, pj2, pj3 = st.columns(3)
with pj1:
    st.markdown(_comm_block("Montant a payer", pay_due_total, 0.0, 0.0, "#6d5efc"), unsafe_allow_html=True)
with pj2:
    st.markdown(_comm_block_with_ratio("Montant Payé", pay_paid_total, 0.0, 0.0, "#ff7a59", pay_ratio), unsafe_allow_html=True)
with pj3:
    st.markdown(_comm_block_with_ratio("Montant Restant", pay_remaining_total, 0.0, 0.0, "#22c55e", pay_remaining_ratio), unsafe_allow_html=True)

pay_alert_df = pay_df.copy()
if broker_scope != "Tous" and p_pay_broker_col and not pay_alert_df.empty:
    if broker_scope == "Sans courtier":
        pay_alert_df = pay_alert_df[_clean_broker_series(pay_alert_df[p_pay_broker_col]).isna()].copy()
    elif broker_scope == "Courtier":
        if broker_choice == "Tous courtiers":
            pay_alert_df = pay_alert_df[_clean_broker_series(pay_alert_df[p_pay_broker_col]).notna()].copy()
        else:
            pay_alert_df = pay_alert_df[_clean_broker_series(pay_alert_df[p_pay_broker_col]) == broker_choice].copy()

if pay_due_col and pay_remaining_col and not pay_alert_df.empty:
    pay_due_df = pay_alert_df.copy()
    pay_due_df[pay_due_col] = _to_datetime(pay_due_df[pay_due_col])
    pay_due_df = pay_due_df.dropna(subset=[pay_due_col]).copy()
    if not pay_due_df.empty:
        pay_due_df["_RESTE"] = _to_num(pay_due_df[pay_remaining_col])
        pay_due_df = pay_due_df[pay_due_df["_RESTE"] > 0].copy()
        if not pay_due_df.empty:
            if pay_date_col and pay_date_col in pay_due_df.columns:
                pay_due_df[pay_date_col] = _to_datetime(pay_due_df[pay_date_col])
                pay_due_df = pay_due_df.dropna(subset=[pay_date_col]).copy()
                pay_due_df["_JOURS_RESTANTS"] = (pay_due_df[pay_due_col].dt.date - pay_due_df[pay_date_col].dt.date).apply(lambda x: int(x.days))
            else:
                pay_due_df["_JOURS_RESTANTS"] = (pay_due_df[pay_due_col].dt.date - date.today()).apply(lambda x: int(x.days))
            jours_restants_min = int(pay_due_df["_JOURS_RESTANTS"].min())
            montant_en_retard = float(pay_due_df.loc[pay_due_df["_JOURS_RESTANTS"] < 0, "_RESTE"].sum())
            nb_lignes_en_retard = int((pay_due_df["_JOURS_RESTANTS"] < 0).sum())

            k1, k2 = st.columns(2)
            with k1:
                st.markdown(_kpi_block("Jours restants (échéance la plus proche)", f"{jours_restants_min} jours", "#f59e0b"), unsafe_allow_html=True)
            with k2:
                st.markdown(_kpi_block("Montant en retard", f"{montant_en_retard:,.0f} FCFA".replace(",", " "), "#ef4444"), unsafe_allow_html=True)

            if nb_lignes_en_retard > 0:
                overdue_df = pay_due_df[pay_due_df["_JOURS_RESTANTS"] < 0].copy()
                if p_pay_broker_col and p_pay_broker_col in overdue_df.columns:
                    overdue_df["Courtier"] = _clean_broker_series(overdue_df[p_pay_broker_col]).fillna("Sans courtier")
                else:
                    overdue_df["Courtier"] = "Courtier non renseigné"
                overdue_tbl = (
                    overdue_df.groupby("Courtier", dropna=False)["_RESTE"]
                    .sum()
                    .reset_index()
                    .rename(columns={"_RESTE": "Montant restant en retard"})
                    .sort_values("Montant restant en retard", ascending=False)
                )
                st.error(f"Alerte: {nb_lignes_en_retard} dossier(s) ont dépassé la date butoir sans être soldés.")
                st.dataframe(
                    overdue_tbl.style.format({"Montant restant en retard": lambda v: f"{v:,.0f}".replace(",", " ")}),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.success("Aucun dossier en retard par rapport à la date butoir.")

pay_status_col = _find_col(pay_f, ["STATUT", "Status", "STATUS", "ETAT", "Etat"])
if pay_status_col and not pay_f.empty:
    pay_status_df = pay_f.copy()
    pay_status_df[pay_status_col] = _clean_broker_series(pay_status_df[pay_status_col]).fillna("Sans statut")
    status_norm = pay_status_df[pay_status_col].astype("string").str.strip().str.upper()
    pay_status_df = pay_status_df[~status_norm.isin(["CANCELLED", "CANCELED", "ANNULE", "ANNULEE"])].copy()
    stat_amount_col = pay_val_col or pay_paid_col or pay_remaining_col
    if stat_amount_col:
        pay_status_df["_STAT_AMOUNT"] = _to_num(pay_status_df[stat_amount_col])
        stat_tbl = (
            pay_status_df.groupby(pay_status_col, dropna=False)["_STAT_AMOUNT"]
            .sum()
            .rename("Chiffre d'affaires")
            .reset_index()
            .rename(columns={pay_status_col: "Statut"})
            .sort_values("Chiffre d'affaires", ascending=False)
        )
        st.markdown("#### Statut Chiffre d'affaires")
        accents = ["#6d5efc", "#ff7a59", "#22c55e"]
        for i in range(0, len(stat_tbl), 3):
            row_df = stat_tbl.iloc[i:i + 3].reset_index(drop=True)
            row_cols = st.columns(3)
            for j in range(3):
                if j >= len(row_df):
                    continue
                r = row_df.iloc[j]
                with row_cols[j]:
                    st.markdown(
                        _comm_block(
                            str(r["Statut"]),
                            float(r["Chiffre d'affaires"]),
                            0.0,
                            0.0,
                            accents[(i + j) % len(accents)],
                        ),
                        unsafe_allow_html=True,
                    )
    else:
        st.warning("Impossible de calculer le chiffre d'affaires par statut: colonne montant introuvable.")

st.markdown("### Commissions")
cj1, cj2, cj3 = st.columns(3)
with cj1:
    st.markdown(_comm_block("Commissions Année", year_val, year_prev, _pct_change(year_val, year_prev), "#6d5efc"), unsafe_allow_html=True)
with cj2:
    st.markdown(_comm_block("Commissions Mois", month_val, month_prev, _pct_change(month_val, month_prev), "#ff7a59"), unsafe_allow_html=True)
with cj3:
    st.markdown(_comm_block("Commission moyenne", mean_val, mean_prev, _pct_change(mean_val, mean_prev), "#22c55e"), unsafe_allow_html=True)

st.markdown("### Renouvellements")
renew_ech = 0
renew_done = 0
renew_rate = 0.0
renew_flag_count = 0
new_business_count = 0
renew_base_df = pd.DataFrame()
renew_policy_col = _find_col(policies_f, ["Policy No", "POLICY NO", "POLICY_NO", "N_POLICE", "N°_POLICE", "N° POLICE"])
renew_flag_col = _find_col(policies_f, ["RENOUVELLEMENT", "Renouvellement", "EST_RENOUVELLER", "EST RENOUVELLER", "EST_RENOUVELER", "EST RENOUVELER"])
pol_product_col = _find_col(policies_f, ["PRODUIT", "Produit", "PRODUCT", "Product", "LINE_OF_BUSINESS", "BRANCHE"])
pol_issue_col = _find_col(policies_f, ["ISSUE_DATE", "Issue Date", "DATE_EMISSION", "DATE EMISSION"])
pol_agent_col = _find_col(policies_f, ["SUBSCRIBER", "Subscriber", "NOM_AGENT", "NOM AGENT", "AGENT", "UTILISATEUR", "Utilisateur"])
if pol_agent_col is None:
    pol_agent_col = p_broker_col
if not policies_f.empty:
    dfx = policies_f.copy()
    if pol_agent_col:
        agent_vals = _clean_broker_series(dfx[pol_agent_col]).dropna().unique().tolist()
        agent_vals = sorted(agent_vals)
        agent_choice = st.selectbox("Agent (renouvellement)", ["Tous"] + agent_vals, index=0)
        if agent_choice != "Tous":
            dfx = dfx[_clean_broker_series(dfx[pol_agent_col]) == agent_choice].copy()
    renew_base_df = dfx.copy()

    if renew_flag_col and not renew_base_df.empty:
        ren_s = renew_base_df[renew_flag_col].astype("string").str.upper().str.strip()
        ren_mask = ren_s.isin(["TRUE", "VRAI", "1", "OUI", "YES"])
        new_mask = ren_s.isin(["FALSE", "FAUX", "0", "NON", "NO"])
        if renew_policy_col and renew_policy_col in renew_base_df.columns:
            all_ids = _norm_id(renew_base_df[renew_policy_col])
            ren_ids = _norm_id(renew_base_df.loc[ren_mask, renew_policy_col])
            new_ids = _norm_id(renew_base_df.loc[new_mask, renew_policy_col])
            all_set = set(all_ids.dropna().unique())
            ren_set = set(ren_ids.dropna().unique())
            new_set = set(new_ids.dropna().unique())
            renew_ech = len(all_set)
            renew_done = len(ren_set)
            renew_flag_count = renew_done
            new_business_count = len(new_set)
        else:
            renew_ech = int(len(renew_base_df))
            renew_done = int(ren_mask.sum())
            renew_flag_count = renew_done
            new_business_count = int(new_mask.sum())
    elif renew_policy_col and not renew_base_df.empty:
        renew_ech = int(_norm_id(renew_base_df[renew_policy_col]).dropna().nunique())
renew_rate = (renew_done / renew_ech) if renew_ech else 0.0

if renew_flag_col:
    st.caption(f"Source Renouvellement: colonne {renew_flag_col} du fichier Polices.")
else:
    st.warning("Colonne RENOUVELLEMENT introuvable dans Polices.")

r1, r2, r3 = st.columns(3)
with r1:
    st.markdown(_kpi_block("Polices à Renouveler", f"{renew_ech:,}".replace(",", " "), "#6d5efc"), unsafe_allow_html=True)
with r2:
    st.markdown(_kpi_block("Polices Renouvelées", f"{renew_done:,}".replace(",", " "), "#ff7a59"), unsafe_allow_html=True)
with r3:
    st.markdown(_kpi_block("Taux de Renouvellement", f"+{renew_rate*100:.1f}%".replace(".", ","), "#22c55e"), unsafe_allow_html=True)

if renew_flag_col and renew_ech:
    nr1, nr2 = st.columns(2)
    with nr1:
        st.markdown(_kpi_block("Renouvellement", f"{renew_flag_count:,}".replace(",", " "), "#ff7a59"), unsafe_allow_html=True)
    with nr2:
        st.markdown(_kpi_block("Nouvelles affaires", f"{new_business_count:,}".replace(",", " "), "#22c55e"), unsafe_allow_html=True)

st.markdown("### Produit")
prod_tbl = pd.DataFrame(columns=["Produit", "Renouvellement", "Nouvelles affaires", "Total contrats", "Taux renouvellement"])
if pol_product_col and renew_flag_col and not renew_base_df.empty:
    prod_df = renew_base_df.copy()
    prod_df[pol_product_col] = prod_df[pol_product_col].astype("string").str.strip().replace("", pd.NA)
    prod_df = prod_df.dropna(subset=[pol_product_col]).copy()
    ren_s = prod_df[renew_flag_col].astype("string").str.upper().str.strip()
    prod_df["_IS_REN"] = ren_s.isin(["TRUE", "VRAI", "1", "OUI", "YES"])
    prod_df["_IS_NEW"] = ren_s.isin(["FALSE", "FAUX", "0", "NON", "NO"])
    if renew_policy_col and renew_policy_col in prod_df.columns:
        prod_df["__PID"] = _norm_id(prod_df[renew_policy_col])
        prod_df = prod_df.dropna(subset=["__PID"]).copy()

    if "__PID" in prod_df.columns:
        ren_by_prod = (
            prod_df.loc[prod_df["_IS_REN"]]
            .groupby(pol_product_col)["__PID"]
            .nunique()
            .rename("Renouvellement")
        )
        new_by_prod = (
            prod_df.loc[prod_df["_IS_NEW"]]
            .groupby(pol_product_col)["__PID"]
            .nunique()
            .rename("Nouvelles affaires")
        )
        tot_by_prod = (
            prod_df.groupby(pol_product_col)["__PID"]
            .nunique()
            .rename("Total contrats")
        )
        prod_tbl = (
            tot_by_prod.to_frame()
            .merge(ren_by_prod, left_index=True, right_index=True, how="left")
            .merge(new_by_prod, left_index=True, right_index=True, how="left")
            .fillna(0)
            .reset_index()
            .rename(columns={pol_product_col: "Produit"})
        )
    else:
        prod_tbl = (
            prod_df.groupby(pol_product_col, dropna=False)
            .agg(
                **{
                    "Renouvellement": ("_IS_REN", "sum"),
                    "Nouvelles affaires": ("_IS_NEW", "sum"),
                    "Total contrats": ("_IS_REN", "size"),
                }
            )
            .reset_index()
            .rename(columns={pol_product_col: "Produit"})
        )
    prod_tbl["Taux renouvellement"] = prod_tbl.apply(
        lambda r: (float(r["Renouvellement"]) / float(r["Total contrats"])) if r["Total contrats"] else 0.0,
        axis=1,
    )
    prod_tbl = prod_tbl.sort_values("Total contrats", ascending=False)
    prod_tbl = prod_tbl[prod_tbl["Produit"].astype("string").str.strip().ne("")].copy()

    st.dataframe(
        prod_tbl.style.format(
            {
                "Renouvellement": lambda v: f"{int(v):,}".replace(",", " "),
                "Nouvelles affaires": lambda v: f"{int(v):,}".replace(",", " "),
                "Total contrats": lambda v: f"{int(v):,}".replace(",", " "),
                "Taux renouvellement": lambda v: f"{float(v)*100:.1f}%".replace(".", ","),
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("Section Produit indisponible: colonnes PRODUIT et RENOUVELLEMENT requises dans Polices.")

# Totaux par période
st.markdown("### Totaux par Période")
pol_date_col = _find_col(policies_df, ["ISSUE_DATE", "ISSUE DATE", "Effective Date"])
quot_date_col = _find_col(quotes_df, ["DATE_CREATION", "DATE CREATION", "DATE CRÉATION", "DATE_CREATION_"])
if pol_date_col:
    pol_dates = _to_datetime(policies_df[pol_date_col]).dropna()
else:
    pol_dates = pd.Series([], dtype="datetime64[ns]")
if quot_date_col:
    quot_dates = _to_datetime(quotes_df[quot_date_col]).dropna()
else:
    quot_dates = pd.Series([], dtype="datetime64[ns]")
ref_date = (pol_dates.max().date() if not pol_dates.empty else (quot_dates.max().date() if not quot_dates.empty else date.today()))

def _bounds_week(d: date) -> tuple[date, date]:
    start = d - pd.Timedelta(days=d.weekday())
    end = start + pd.Timedelta(days=6)
    return start, end

def _bounds_month(d: date) -> tuple[date, date]:
    start = d.replace(day=1)
    end = (start + pd.offsets.MonthEnd(1)).date()
    return start, end

def _bounds_year(d: date) -> tuple[date, date]:
    start = date(d.year, 1, 1)
    end = date(d.year, 12, 31)
    return start, end

week_ref = quot_dates.max().date() if not quot_dates.empty else ref_date
periods = {
    "Semaine": _bounds_week(week_ref),
    "Mois": _bounds_month(ref_date),
    "Année": _bounds_year(ref_date),
}

def _count_quotes(df: pd.DataFrame, start_d: date, end_d: date, only_broker: bool) -> int:
    if df.empty or not q_id_col:
        return 0
    tmp = df.copy()
    if q_broker_col:
        broker_mask = _clean_broker_series(tmp[q_broker_col]).notna()
        if only_broker:
            tmp = tmp[broker_mask].copy()
        else:
            tmp = tmp[~broker_mask].copy()
    if quot_date_col and quot_date_col in tmp.columns:
        tmp[quot_date_col] = _to_datetime(tmp[quot_date_col])
        valid = tmp[quot_date_col].notna()
        if valid.any():
            tmp = tmp[valid].copy()
            mask = (tmp[quot_date_col].dt.date >= start_d) & (tmp[quot_date_col].dt.date <= end_d)
            tmp = tmp.loc[mask]
    return int(tmp[q_id_col].astype("string").replace("", pd.NA).dropna().nunique())

def _count_policies(df: pd.DataFrame, start_d: date, end_d: date, only_broker: bool) -> tuple[int, float]:
    if df.empty or not policy_col or not pol_date_col:
        return 0, 0.0
    tmp = df.copy()
    tmp[pol_date_col] = _to_datetime(tmp[pol_date_col])
    tmp = tmp.dropna(subset=[pol_date_col])
    mask = (tmp[pol_date_col].dt.date >= start_d) & (tmp[pol_date_col].dt.date <= end_d)
    if p_broker_col:
        broker_mask = _clean_broker_series(tmp[p_broker_col]).notna()
        if only_broker:
            mask = mask & broker_mask
        else:
            mask = mask & (~broker_mask)
    tmp = tmp.loc[mask]
    count = int(tmp[policy_col].astype("string").replace("", pd.NA).dropna().nunique())
    amount = float(_to_num(tmp[amount_col]).sum()) if amount_col and not tmp.empty else 0.0
    return count, amount

q_base_period = quotes_df.copy()
p_base_period = policies_df.copy()
if broker_scope != "Tous":
    if broker_scope == "Sans courtier":
        if q_broker_col:
            q_base_period = q_base_period[_clean_broker_series(q_base_period[q_broker_col]).isna()].copy()
        if p_broker_col:
            p_base_period = p_base_period[_clean_broker_series(p_base_period[p_broker_col]).isna()].copy()
    elif broker_scope == "Courtier" and broker_choice:
        if broker_choice == "Tous courtiers":
            if q_broker_col:
                q_base_period = q_base_period[_clean_broker_series(q_base_period[q_broker_col]).notna()].copy()
            if p_broker_col:
                p_base_period = p_base_period[_clean_broker_series(p_base_period[p_broker_col]).notna()].copy()
        else:
            if q_broker_col:
                q_base_period = q_base_period[_clean_broker_series(q_base_period[q_broker_col]) == broker_choice].copy()
            if p_broker_col:
                p_base_period = p_base_period[_clean_broker_series(p_base_period[p_broker_col]) == broker_choice].copy()

rows = []
for per_label, (p_start, p_end) in periods.items():
    for typ_label, only_broker in [("Courtiers", True), ("Compagnie", False)]:
        cot = _count_quotes(q_base_period, p_start, p_end, only_broker)
        conv, amt = _count_policies(p_base_period, p_start, p_end, only_broker)
        rows.append({
            "Période": per_label,
            "Type": typ_label,
            "Cotations": cot,
            "Convertis": conv,
            "Montant Total": amt,
        })

df_period_totals = pd.DataFrame(rows)
df_period_totals["Montant Total"] = df_period_totals["Montant Total"].astype(float)
st.dataframe(
    df_period_totals.style.format({
        "Cotations": lambda v: f"{int(v):,}".replace(",", " "),
        "Convertis": lambda v: f"{int(v):,}".replace(",", " "),
        "Montant Total": lambda v: f"{float(v):,.0f} FCFA".replace(",", " "),
    }),
    use_container_width=True,
    hide_index=True,
)

# Top 10 Performeurs
st.markdown("### Top 10 Courtiers")
show_top10 = broker_scope == "Tous" or (broker_scope == "Courtier" and broker_choice == "Tous courtiers")

if show_top10:
    top_tbl = broker_perf.head(10).copy()
    st.dataframe(
        top_tbl.style.format({
            "Cotations": lambda v: f"{int(v):,}".replace(",", " "),
            "Convertis": lambda v: f"{int(v):,}".replace(",", " "),
            "Montant": lambda v: f"{float(v):,.0f}".replace(",", " "),
            "Taux": lambda v: f"{float(v)*100:.1f}%".replace(".", ","),
        }),
        use_container_width=True,
        hide_index=True,
    )
elif broker_scope == "Courtier" and broker_choice != "Tous courtiers":
    selected = broker_perf[broker_perf["Courtier"] == broker_choice].copy()
    if not selected.empty:
        selected = selected.iloc[0]
        total_montant = float(broker_perf["Montant"].sum()) if not broker_perf.empty else 0.0
        share = (float(selected["Montant"]) / total_montant) if total_montant else 0.0
        rank = int(broker_perf.index[broker_perf["Courtier"] == broker_choice][0]) + 1
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1:
            st.markdown(_kpi_block("Rang global", f"{rank}", "#6d5efc"), unsafe_allow_html=True)
        with sc2:
            st.markdown(_kpi_block("Part de réalisation", f"{share*100:.1f}%".replace(".", ","), "#ff7a59"), unsafe_allow_html=True)
        with sc3:
            st.markdown(_kpi_block("Montant", f"{float(selected['Montant']):,.0f}".replace(",", " ") + " FCFA", "#22c55e"), unsafe_allow_html=True)
        with sc4:
            st.markdown(_kpi_block("Taux conversion", f"{float(selected['Taux'])*100:.1f}%".replace(".", ","), "#0ea5e9"), unsafe_allow_html=True)
    else:
        st.info("Aucune donnée trouvée pour ce courtier sur la période.")

# Répartition par Type d'Affaires
st.markdown("### Répartition par Type d'Affaires")
q_type_col = _find_col(quotes_f, ["BRANCHE", "Branche", "BUSINESS_TYPE", "Business Type"])
p_type_col = _find_col(policies_f, ["BUSINESS_TYPE", "Business Type", "BRANCHE", "Branche"])
pol_amount_col = _find_col(policies_f, ["AMOUNT", "Amount"])

type_tbl = pd.DataFrame(columns=["Type", "Cotations", "Convertis", "Montant", "Taux"])
if q_type_col and q_id_col:
    q_type = quotes_f[[q_type_col, q_id_col]].copy()
    q_type[q_type_col] = q_type[q_type_col].astype("string").str.strip().replace("", pd.NA)
    q_type["__qid"] = _norm_id(q_type[q_id_col])
    q_type = q_type.dropna(subset=[q_type_col, "__qid"]).copy()
    cot_by_type = q_type.groupby(q_type_col)["__qid"].nunique().rename("Cotations")

    conv_by_type = pd.Series(dtype=int, name="Convertis")
    amt_by_type = pd.Series(dtype=float, name="Montant")
    if not policies_f.empty and p_type_col and pol_quote_col:
        cols = [p_type_col, pol_quote_col]
        if policy_col:
            cols.append(policy_col)
        if pol_amount_col:
            cols.append(pol_amount_col)
        p_type = policies_f[cols].copy()
        p_type[p_type_col] = p_type[p_type_col].astype("string").str.strip().replace("", pd.NA)
        p_type["__pid"] = _norm_id(p_type[pol_quote_col])
        p_type = p_type.dropna(subset=[p_type_col, "__pid"]).copy()

        merged_t = q_type.merge(p_type, left_on="__qid", right_on="__pid", how="inner")
        if policy_col and policy_col in merged_t.columns:
            conv_by_type = merged_t.groupby(q_type_col)[policy_col].nunique().rename("Convertis")
        else:
            conv_by_type = merged_t.groupby(q_type_col)["__pid"].nunique().rename("Convertis")
        if pol_amount_col and pol_amount_col in merged_t.columns:
            amt_by_type = merged_t.groupby(q_type_col)[pol_amount_col].apply(lambda s: float(_to_num(s).sum())).rename("Montant")

    type_tbl = cot_by_type.to_frame().merge(conv_by_type, left_index=True, right_index=True, how="outer")
    if not amt_by_type.empty:
        type_tbl = type_tbl.merge(amt_by_type, left_index=True, right_index=True, how="outer")
    else:
        type_tbl["Montant"] = 0.0
    type_tbl = type_tbl.fillna(0)
    type_tbl["Taux"] = type_tbl.apply(lambda r: (r["Convertis"] / r["Cotations"]) if r["Cotations"] else 0.0, axis=1)
    type_tbl = type_tbl.reset_index().rename(columns={q_type_col: "Type"})
    type_tbl = type_tbl[type_tbl["Type"].astype("string").str.strip().ne("")].copy()

st.dataframe(
    type_tbl.style.format({
        "Cotations": lambda v: f"{int(v):,}".replace(",", " "),
        "Convertis": lambda v: f"{int(v):,}".replace(",", " "),
        "Montant": lambda v: f"{float(v):,.0f}".replace(",", " "),
        "Taux": lambda v: f"{float(v)*100:.1f}%".replace(".", ","),
    }),
    use_container_width=True,
    hide_index=True,
)

st.markdown("### Export")
sections_pdf = [
    (
        "Mes Cotations",
        [
            ("Cotations", _fmt_int(year_count)),
            ("Montant cotations", _fmt_money(quote_amount)),
            ("Cotations moyenne", _fmt_int(round(avg_cotations))),
        ],
    ),
    (
        "Police",
        [
            ("Polices créées", _fmt_int(polices_creees)),
            ("Chiffres d'affaires", _fmt_money(ca_total)),
            ("Part réalisée", f"{part_realisee*100:.1f}%".replace(".", ",")),
        ],
    ),
    (
        "Paiements",
        [
            ("Montant à payer", _fmt_money(pay_year_val)),
            ("Montant payé", _fmt_money(pay_paid_month_val)),
            ("Montant restant", _fmt_money(pay_remaining_val)),
        ],
    ),
    (
        "Commissions",
        [
            ("Commissions année", _fmt_money(year_val)),
            ("Commissions mois", _fmt_money(month_val)),
            ("Commission moyenne", _fmt_money(mean_val)),
        ],
    ),
    (
        "Renouvellements",
        [
            ("Polices à renouveler", _fmt_int(renew_ech)),
            ("Polices renouvelées", _fmt_int(renew_done)),
            ("Nouvelles affaires", _fmt_int(new_business_count)),
        ],
    ),
]
if not prod_tbl.empty:
    prod_rows_pdf = []
    for _, r in prod_tbl.head(6).iterrows():
        prod_name = str(r.get("Produit", "")).strip()
        if not prod_name:
            continue
        prod_rows_pdf.append(
            (
                f"{prod_name} (Renouvellement / Nouvelles affaires)",
                f"{_fmt_int(r.get('Renouvellement', 0))} / {_fmt_int(r.get('Nouvelles affaires', 0))}",
            )
        )
    if prod_rows_pdf:
        sections_pdf.append(("Produit", prod_rows_pdf))
if not df_period_totals.empty:
    period_rows_pdf = []
    for _, r in df_period_totals.iterrows():
        typ = str(r.get("Type", "")).strip()
        if typ.lower() == "compagnie":
            continue
        period_rows_pdf.append(
            (
                f"{r.get('Période', '')} - {typ}",
                f"Cot:{_fmt_int(r.get('Cotations', 0))} | Conv:{_fmt_int(r.get('Convertis', 0))} | Mont:{_fmt_int(r.get('Montant Total', 0))}",
            )
        )
    if period_rows_pdf:
        sections_pdf.append(("Totaux par Période", period_rows_pdf))
if not type_tbl.empty:
    type_rows_pdf = []
    for _, r in type_tbl.head(10).iterrows():
        tname = str(r.get("Type", "")).strip()
        if not tname:
            continue
        tlabel = tname if len(tname) <= 34 else (tname[:31] + "...")
        type_rows_pdf.append(
            (
                tlabel,
                f"Cot:{_fmt_int(r.get('Cotations', 0))} | Conv:{_fmt_int(r.get('Convertis', 0))} | Taux:{float(r.get('Taux', 0))*100:.1f}%".replace(".", ","),
            )
        )
    if type_rows_pdf:
        sections_pdf.append(("Répartition par Type d'Affaires", type_rows_pdf))
dr_txt = "Non définie"
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    dr_txt = f"{date_range[0].isoformat()} -> {date_range[1].isoformat()}"
meta_pdf = [
    ("Courtier", broker_filter_label),
    ("Période", dr_txt),
]
recap_lines_pdf: list[str] = []
if not broker_perf.empty and broker_scope == "Courtier" and broker_choice != "Tous courtiers":
    rec_sel = broker_perf[broker_perf["Courtier"] == broker_choice]
    if not rec_sel.empty:
        rr = rec_sel.iloc[0]
        total_montant_all = float(broker_perf["Montant"].sum()) if not broker_perf.empty else 0.0
        part_rr = (float(rr["Montant"]) / total_montant_all) if total_montant_all else 0.0
        rank_rr = int(broker_perf.index[broker_perf["Courtier"] == broker_choice][0]) + 1
        total_rr = int(len(broker_perf))
        recap_lines_pdf = [
            f"Le chiffre d'affaires realise sur la periode est {_fmt_money(float(rr['Montant']))} pour {broker_choice}.",
            f"Ta part de realisation dans le chiffre d'affaires global de tous les courtiers est {part_rr*100:.1f}%.".replace(".", ","),
            f"Le taux de conversion par rapport au nombre de convertis est {float(rr['Taux'])*100:.1f}%.".replace(".", ","),
            f"Ton rang global par rapport au nombre de courtiers est {rank_rr}/{total_rr}.",
        ]
else:
    recap_lines_pdf = ["Selectionne un courtier specifique pour afficher son recapitulatif detaille."]

logo_pdf_bytes: bytes | None = None
logo_pdf_source = "Aucun"
logo_pdf_align = "right"
if pdf_logo_mode == "Upload":
    if pdf_logo_file is not None:
        try:
            logo_pdf_bytes = pdf_logo_file.getvalue()
            logo_pdf_source = f"Upload: {pdf_logo_file.name}"
            if "assur" in pdf_logo_file.name.lower() and "defender" in pdf_logo_file.name.lower():
                logo_pdf_align = "left"
                logo_pdf_bytes = _remove_dark_background(logo_pdf_bytes)
        except Exception:
            logo_pdf_bytes = None
elif pdf_logo_mode == "Leadway":
    for _logo_path in [Path("agent_report/assets/logo_leadway.jpg"), Path("assets/logo_leadway.jpg"), Path("logo_leadway.jpg")]:
        if _logo_path.exists():
            try:
                logo_pdf_bytes = _logo_path.read_bytes()
                logo_pdf_source = f"Fichier: {_logo_path}"
                logo_pdf_align = "right"
                break
            except Exception:
                logo_pdf_bytes = None
elif pdf_logo_mode == "Assur Defender":
    for _logo_path in [Path("assets/logo_assur_defender.png"), Path("assets/logo_assur_defender.jpg"), Path("logo_assur_defender.png"), Path("logo_assur_defender.jpg")]:
        if _logo_path.exists():
            try:
                logo_pdf_bytes = _remove_dark_background(_logo_path.read_bytes())
                logo_pdf_source = f"Fichier: {_logo_path}"
                logo_pdf_align = "left"
                break
            except Exception:
                logo_pdf_bytes = None
else:
    for _logo_path in [
        Path("assets/logo_assur_defender.png"),
        Path("assets/logo_assur_defender.jpg"),
        Path("logo_assur_defender.png"),
        Path("logo_assur_defender.jpg"),
        Path("agent_report/assets/logo_leadway.jpg"),
        Path("assets/logo_leadway.jpg"),
        Path("logo_leadway.jpg"),
    ]:
        if _logo_path.exists():
            try:
                logo_pdf_bytes = _logo_path.read_bytes()
                logo_pdf_source = f"Fichier: {_logo_path}"
                if "assur" in _logo_path.name.lower() and "defender" in _logo_path.name.lower():
                    logo_pdf_align = "left"
                    logo_pdf_bytes = _remove_dark_background(logo_pdf_bytes)
                else:
                    logo_pdf_align = "right"
                break
            except Exception:
                logo_pdf_bytes = None
if is_admin:
    if logo_pdf_bytes is not None:
        st.sidebar.caption(f"Logo PDF actif: {logo_pdf_source}")
    else:
        st.sidebar.caption("Logo PDF actif: aucun")

payment_paid_donut_png = _build_ratio_donut_png(pay_ratio)
payment_remaining_donut_png = _build_ratio_donut_png(pay_remaining_ratio)
pdf_bytes = _build_dashboard_pdf(
    "RAPPORT DE VOTRE ACTIVITE",
    meta_pdf,
    sections_pdf,
    payment_paid_donut_png=payment_paid_donut_png,
    payment_remaining_donut_png=payment_remaining_donut_png,
    recap_lines=recap_lines_pdf,
    logo_png=logo_pdf_bytes,
    logo_align=logo_pdf_align,
)
st.download_button(
    "Télécharger le rapport PDF",
    data=pdf_bytes,
    file_name=f"rapport_plateforme_courtier_{date.today().isoformat()}.pdf",
    mime="application/pdf",
)
