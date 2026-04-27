# glpi_connect.py — Connexion GLPI
# Projet CIMAT Béni Mellal

import requests

GLPI_URL   = "http://192.168.112.129/glpi"
APP_TOKEN  = "Qzqd9JxK0k9vcF9daXxT4zIJKZGht2r20pHgaYR4"
USER_TOKEN = "gMJqxaSdpy79c9HQout8reOo1PBY0UfM15FkQGYs"

def connect_glpi():
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"user_token {USER_TOKEN}",
        "App-Token":     APP_TOKEN
    }
    response = requests.get(
        f"{GLPI_URL}/apirest.php/initSession",
        headers=headers
    )

    data = response.json()

    # ── Diagnostic : affiche ce que GLPI renvoie réellement ──
    print(f"Status code : {response.status_code}")
    print(f"Réponse brute : {data}")

    # ── Si GLPI renvoie une erreur sous forme de liste ────────
    if isinstance(data, list):
        print(f" GLPI a retourné une erreur : {data}")
        return None

    # ── Connexion réussie ─────────────────────────────────────
    if "session_token" in data:
        print(f" Connexion réussie ! Session : {data['session_token'][:12]}…")
        return data["session_token"]
    else:
        print(f" Réponse inattendue : {data}")
        return None