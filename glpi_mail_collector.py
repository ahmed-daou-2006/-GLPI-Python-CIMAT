# glpi_mail_collector.py
# Collecteur email → Création automatique de tickets GLPI
# Projet CIMAT Béni Mellal

import imaplib
import email
import requests
import time
from email.header import decode_header
from datetime import datetime

# ===== CONFIG GLPI =====
GLPI_URL   = "http://192.168.112.129/glpi"
APP_TOKEN  = "Qzqd9JxK0k9vcF9daXxT4zIJKZGht2r20pHgaYR4"
USER_TOKEN = "gMJqxaSdpy79c9HQout8reOo1PBY0UfM15FkQGYs"

# ===== CONFIG GMAIL =====
GMAIL_USER     = "ahmeddaou2006@gmail.com"
GMAIL_PASSWORD = "zkactlknxhsmoztu" 

# ===== CONNEXION GLPI =====
def connect_glpi():
    r = requests.get(
        f"{GLPI_URL}/apirest.php/initSession",
        headers={
            "App-Token": APP_TOKEN,
            "Authorization": f"user_token {USER_TOKEN}"
        }
    )
    return r.json()["session_token"]

# ===== CRÉER TICKET =====
def create_ticket(session_token, subject, body, sender):
    headers = {
        "App-Token": APP_TOKEN,
        "Session-Token": session_token,
        "Content-Type": "application/json"
    }
    payload = {"input": {
        "name":        f"[EMAIL] {subject[:200]}",
        "content":     f"De: {sender}\n\n{body}",
        "type":        1,
        "status":      1,
        "urgency":     3,
        "priority":    3,
        "entities_id": 1
    }}
    r = requests.post(
        f"{GLPI_URL}/apirest.php/Ticket",
        headers=headers, json=payload
    )
    return r.json()

# ===== COLLECTER EMAILS =====
def collect_emails():
    print(f"\n{'='*50}")
    print(f" Collecte — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}")

    try:
        # Connexion Gmail
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(GMAIL_USER, GMAIL_PASSWORD)
        mail.select("INBOX")

        #  Filtre — seulement emails avec [TICKET] dans le sujet
        _, messages = mail.search(None, 'UNSEEN SUBJECT "[TICKET]"')
        email_ids = messages[0].split()

        if not email_ids:
            print(" Aucun email [TICKET] non lu.")
            mail.logout()
            return

        session_token = connect_glpi()
        print(f" GLPI connecté | {len(email_ids)} email(s) trouvé(s)")

        for num in email_ids:
            _, msg_data = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            # Décoder sujet
            raw_subject, enc = decode_header(msg["Subject"])[0]
            subject = raw_subject.decode(enc or "utf-8") if isinstance(raw_subject, bytes) else raw_subject

            sender = msg.get("From", "Inconnu")

            # Extraire corps
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

            # Créer ticket GLPI
            result = create_ticket(session_token, subject, body, sender)
            print(f" Ticket #{result.get('id')} créé — {subject[:50]}")
            print(f" Expéditeur : {sender}")

            # Marquer comme lu
            mail.store(num, "+FLAGS", "\\Seen")

        mail.logout()

    except Exception as e:
        print(f" Erreur : {e}")

# ===== LANCEMENT =====
if __name__ == "__main__":
    print(" Collecteur GLPI démarré — vérification toutes les 5 min")
    print(" Filtre actif : Objet contient [TICKET]")
    print("  Ctrl+C pour arrêter\n")
    while True:
        collect_emails()
        time.sleep(300)