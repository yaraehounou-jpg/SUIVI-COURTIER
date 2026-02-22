# ===============================================================
# auto_send.py – Envoi automatique des échéances (Leadway IARD)
# ===============================================================
# Auteur : Yara Kouadio
# Objectif :
#   - Le 23 → envoi automatique des renouvellements du mois suivant
#   - Le 1er du mois suivant → relance automatique des renouvellements
# ===============================================================

# ===============================================================
# auto_send.py – Envoi automatique des échéances + confirmation mail
# ===============================================================
import os
import datetime
import subprocess
import smtplib
import ssl
from email.message import EmailMessage

# --- 1️⃣ Paramètres globaux ---
PROJECT_DIR = "/Users/y-kouadio/Documents/WORKSPACE/agent_report"
MAIN_SCRIPT = os.path.join(PROJECT_DIR, "main.py")
DATA_FILE = os.path.join(PROJECT_DIR, "data", "data_to_generate.xlsx")

# --- ⚙️ SMTP Configuration (ton compte Leadway) ---
SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = "leadvie@leadway.com"   # <-- ton email expéditeur
SMTP_PASS = "bM_BS[EZc7/m"    # <-- ton mot de passe ou mot de passe d’application
ADMIN_EMAIL = "y-kouadio@leadway.com" # <-- ton adresse de réception du rapport

# --- 2️⃣ Vérifier la date du jour ---
today = datetime.date.today()
day = today.day
month = today.month
year = today.year

# --- 3️⃣ Calcul du mois "cible" ---
if day == 23:
    target_month = (month % 12) + 1
    target_year = year + (1 if month == 12 else 0)
    mode = "PREPARATION"
elif day == 1:
    target_month = month
    target_year = year
    mode = "RELANCE"
else:
    print(f"⏸️ {today} – Pas une date d'envoi automatique (seulement le 23 et le 1er).")
    exit(0)

# --- 4️⃣ Vérifier le fichier de données ---
if not os.path.exists(DATA_FILE):
    msg = f"⚠️ Fichier introuvable : {DATA_FILE}\n➡️ Dépose le fichier à jour avant le 23."
    print(msg)
    exit(1)

# --- 5️⃣ Exécution principale ---
month_label = datetime.date(target_year, target_month, 1).strftime("%B %Y").capitalize()
print("===============================================================")
print(f"📅 {today} — Exécution automatique ({mode})")
print(f"🎯 Mois cible : {month_label}")
print(f"📦 Données utilisées : {DATA_FILE}")
print("===============================================================")

log_message = ""
success = False

try:
    subprocess.run(["python3", MAIN_SCRIPT], cwd=PROJECT_DIR, check=True)
    success = True
    log_message = f"✅ Envoi automatique terminé avec succès ({mode.lower()} pour {month_label})."
    print(log_message)
except subprocess.CalledProcessError as e:
    log_message = f"❌ Erreur lors de l’envoi ({mode}) : {e}"
    print(log_message)

# --- 6️⃣ Envoi du mail de confirmation ---
def send_confirmation_email(subject, body):
    """Envoie un email de confirmation via SMTP Office365."""
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls(context=ctx)
            server.login(SMTP_USER, SMTP_PASS)
            msg = EmailMessage()
            msg["From"] = f"Leadway AutoBot <{SMTP_USER}>"
            msg["To"] = ADMIN_EMAIL
            msg["Subject"] = subject
            msg.set_content(body)
            server.send_message(msg)
        print(f"📧 Mail de confirmation envoyé à {ADMIN_EMAIL}")
    except Exception as e:
        print(f"⚠️ Erreur lors de l’envoi du mail de confirmation : {e}")

# Corps du message
subject = f"Rapport automatique – {mode} {month_label}"
body = (
    f"Bonjour Yara,\n\n"
    f"Exécution automatique du script d’envoi des échéances effectuée.\n"
    f"📅 Date : {today}\n"
    f"⚙️ Mode : {mode}\n"
    f"🎯 Mois cible : {month_label}\n"
    f"📦 Fichier utilisé : {DATA_FILE}\n\n"
    f"{log_message}\n\n"
    f"Cordialement,\n"
    f"Le Robot Leadway Auto-Renouvellement 🤖"
)

send_confirmation_email(subject, body)
