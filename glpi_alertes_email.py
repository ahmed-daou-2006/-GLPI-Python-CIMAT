# glpi_alertes_email.py — Notifications tickets non résolus
# Projet CIMAT Béni Mellal

import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timezone

# ── Config GLPI ───────────────────────────────────────────
GLPI_URL   = "http://192.168.112.129/glpi"
APP_TOKEN  = "Qzqd9JxK0k9vcF9daXxT4zIJKZGht2r20pHgaYR4"
USER_TOKEN = "gMJqxaSdpy79c9HQout8reOo1PBY0UfM15FkQGYs"

# ── Config Email ──────────────────────────────────────────
SMTP_SERVER   = "smtp.gmail.com"
SMTP_PORT     = 587
EMAIL_EXPEDITEUR  = "ahmeddaou2006@gmail.com"       # ← Ton Gmail
EMAIL_MOT_PASSE   = "edmq idff wylc rrva"    # ← App Password Gmail
EMAIL_DESTINATAIRE = "ahmeddaou2006@gmail.com"     # ← Email du responsable

# Seuil : tickets ouverts depuis plus de X heures → alerte
SEUIL_HEURES = 0

PRIORITES = {1:"Très basse", 2:"Basse", 3:"Moyenne",
             4:"Haute", 5:"Très haute", 6:"Majeure"}
STATUTS   = {1:"Nouveau", 2:"En cours (Assigné)",
             3:"En cours (Planifié)", 4:"En attente"}

# ── Connexion GLPI ────────────────────────────────────────
def connect_glpi():
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"user_token {USER_TOKEN}",
        "App-Token":     APP_TOKEN
    }
    r = requests.get(f"{GLPI_URL}/apirest.php/initSession", headers=headers)
    data = r.json()
    if isinstance(data, list) or "session_token" not in data:
        print(f"❌ Erreur connexion : {data}")
        return None
    print(f" Connecté — Session : {data['session_token'][:12]}…")
    return data["session_token"]

def disconnect_glpi(session_token):
    headers = {
        "Content-Type":  "application/json",
        "Session-Token": session_token,
        "App-Token":     APP_TOKEN
    }
    requests.get(f"{GLPI_URL}/apirest.php/killSession", headers=headers)
    print(" Session fermée.")

# ── Récupération tickets non résolus ─────────────────────
def get_tickets_non_resolus(session_token):
    headers = {
        "Content-Type":  "application/json",
        "Session-Token": session_token,
        "App-Token":     APP_TOKEN
    }
    params = {
        "expand_dropdowns": True,
        "range": "0-999",
        "sort": "date",
        "order": "ASC"
    }
    r = requests.get(
        f"{GLPI_URL}/apirest.php/Ticket",
        headers=headers,
        params=params
    )
    if r.status_code != 200:
        print(f"❌ Erreur : {r.text}")
        return []

    tous_tickets = r.json()
    maintenant   = datetime.now(timezone.utc)
    non_resolus  = []

    for t in tous_tickets:
        # On garde statuts : Nouveau(1), En cours(2,3), En attente(4)
        if t.get("status") not in [1, 2, 3, 4]:
            continue

        # Calcul ancienneté
        date_str = t.get("date", "")
        if date_str:
            try:
                date_ouverture = datetime.fromisoformat(
                    date_str.replace(" ", "T")
                ).replace(tzinfo=timezone.utc)
                heures_ouvert = (maintenant - date_ouverture).total_seconds() / 3600
                t["heures_ouvert"] = round(heures_ouvert, 1)

                if heures_ouvert >= SEUIL_HEURES:
                    non_resolus.append(t)
            except:
                t["heures_ouvert"] = "?"
                non_resolus.append(t)

    print(f"  {len(non_resolus)} ticket(s) non résolu(s) depuis +{SEUIL_HEURES}h.")
    return non_resolus

# ── Construction du corps HTML de l'email ────────────────
def construire_email_html(tickets):
    date_rapport = datetime.now().strftime("%d/%m/%Y à %H:%M")

    lignes_tickets = ""
    for t in tickets:
        priorite_id = t.get("priority", 3)
        priorite    = PRIORITES.get(priorite_id, "?")
        statut      = STATUTS.get(t.get("status"), "?")
        heures      = t.get("heures_ouvert", "?")

        # Couleur selon priorité
        if priorite_id >= 5:
            couleur = "#ff4d4d"
        elif priorite_id == 4:
            couleur = "#ff9900"
        else:
            couleur = "#2E75B6"

        lignes_tickets += f"""
        <tr>
          <td style="padding:8px;border:1px solid #ddd;">{t.get('id','')}</td>
          <td style="padding:8px;border:1px solid #ddd;">{t.get('name','')}</td>
          <td style="padding:8px;border:1px solid #ddd;">{statut}</td>
          <td style="padding:8px;border:1px solid #ddd;
                     color:{couleur};font-weight:bold;">{priorite}</td>
          <td style="padding:8px;border:1px solid #ddd;">{heures}h</td>
        </tr>"""

    html = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333;">
      <div style="background:#1F4E79;color:white;padding:20px;border-radius:8px;">
        <h2 style="margin:0;"> Alerte GLPI — CIMAT Béni Mellal</h2>
        <p style="margin:5px 0 0;">Rapport du {date_rapport}</p>
      </div>

      <div style="padding:20px;">
        <p>Bonjour,</p>
        <p>Les tickets suivants sont <strong>non résolus depuis plus
           de {SEUIL_HEURES} heures</strong> :</p>

        <table style="width:100%;border-collapse:collapse;margin-top:10px;">
          <thead>
            <tr style="background:#2E75B6;color:white;">
              <th style="padding:10px;border:1px solid #ddd;">ID</th>
              <th style="padding:10px;border:1px solid #ddd;">Titre</th>
              <th style="padding:10px;border:1px solid #ddd;">Statut</th>
              <th style="padding:10px;border:1px solid #ddd;">Priorité</th>
              <th style="padding:10px;border:1px solid #ddd;">Ancienneté</th>
            </tr>
          </thead>
          <tbody>{lignes_tickets}</tbody>
        </table>

        <div style="margin-top:20px;padding:15px;
                    background:#fff3cd;border-radius:6px;
                    border-left:4px solid #ff9900;">
          <strong>Total : {len(tickets)} ticket(s) en attente</strong><br>
          Veuillez traiter ces tickets dès que possible.
        </div>

        <p style="margin-top:20px;">
           <a href="{GLPI_URL}/front/ticket.php">
             Accéder à GLPI</a>
        </p>

        <p style="color:#888;font-size:12px;margin-top:30px;">
          — Email automatique généré par le système GLPI CIMAT
        </p>
      </div>
    </body></html>"""

    return html

# ── Envoi de l'email ──────────────────────────────────────
def envoyer_email(tickets):
    if not tickets:
        print(" Aucun ticket en retard — aucun email envoyé.")
        return

    html = construire_email_html(tickets)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f" GLPI CIMAT — {len(tickets)} ticket(s) non résolu(s)"
    msg["From"]    = EMAIL_EXPEDITEUR
    msg["To"]      = EMAIL_DESTINATAIRE
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as serveur:
            serveur.starttls()
            serveur.login(EMAIL_EXPEDITEUR, EMAIL_MOT_PASSE)
            serveur.sendmail(EMAIL_EXPEDITEUR, EMAIL_DESTINATAIRE, msg.as_string())
        print(f" Email envoyé à {EMAIL_DESTINATAIRE}")
    except Exception as e:
        print(f" Erreur envoi email : {e}")

# ── MAIN ──────────────────────────────────────────────────
if __name__ == "__main__":
    session = connect_glpi()
    if session:
        tickets = get_tickets_non_resolus(session)
        envoyer_email(tickets)
        disconnect_glpi(session)