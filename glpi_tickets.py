# glpi_tickets.py — Création automatique de tickets
# Projet CIMAT Béni Mellal

import requests
from glpi_connect import connect_glpi, GLPI_URL, APP_TOKEN

# ─── Référence rapide ───────────────────────────────────────
# type     : 1=Incident       2=Demande
# status   : 1=Nouveau        2=En cours   5=Résolu   6=Clos
# priority : 1=Très basse     3=Moyenne    5=Très haute
# urgency  : 1=Très basse     3=Moyenne    5=Très haute
# ────────────────────────────────────────────────────────────

def create_ticket(session_token, titre, description,
                  priorite=3, type_ticket=1, urgence=3):
    """Crée un ticket et retourne son ID."""
    headers = {
        "Content-Type":  "application/json",
        "Session-Token": session_token,
        "App-Token":     APP_TOKEN
    }
    data = {
        "input": {
            "name":        titre,
            "content":     description,
            "type":        type_ticket,
            "priority":    priorite,
            "urgency":     urgence,
            "impact":      3,
            "status":      1,
            "entities_id": 0
        }
    }
    response = requests.post(
        f"{GLPI_URL}/apirest.php/Ticket",
        headers=headers,
        json=data
    )
    if response.status_code == 201:
        ticket_id = response.json()["id"]
        print(f" Ticket créé  —  ID : {ticket_id}  |  {titre}")
        return ticket_id
    else:
        print(f" Erreur {response.status_code} : {response.text}")
        return None


def create_plusieurs_tickets(session_token, liste):
    """
    Crée plusieurs tickets depuis une liste de dicts.

    Exemple de liste :
    [
        {"titre": "Panne réseau",  "description": "Switch HS salle 3", "priorite": 5},
        {"titre": "Demande accès", "description": "Nouveau employé",   "priorite": 2},
    ]
    """
    ids = []
    print(f" Création de {len(liste)} ticket(s)...\n")
    for t in liste:
        tid = create_ticket(
            session_token,
            titre       = t.get("titre",       "Sans titre"),
            description = t.get("description", ""),
            priorite    = t.get("priorite",    3),
            type_ticket = t.get("type",        1),
            urgence     = t.get("urgence",     3)
        )
        if tid:
            ids.append(tid)
    print(f"\n {len(ids)}/{len(liste)} ticket(s) créé(s) avec succès.")
    return ids


# ─── TEST ────────────────────────────────────────────────────
if __name__ == "__main__":

    session = connect_glpi()

    if session:
        # ── Test 1 : ticket unique ──────────────────────────
        create_ticket(
            session,
            titre       = "Panne imprimante atelier",
            description = "L'imprimante HP du bureau 12 ne répond plus.",
            priorite    = 4,
            type_ticket = 1    # Incident
        )

        # ── Test 2 : plusieurs tickets en une fois ──────────
        tickets_a_creer = [
            {
                "titre":       "Problème réseau salle serveur",
                "description": "Connexion instable depuis ce matin.",
                "priorite":    5,
                "type":        1
            },
            {
                "titre":       "Demande installation logiciel",
                "description": "Besoin d'AutoCAD sur poste bureau 7.",
                "priorite":    2,
                "type":        2    # Demande
            },
            {
                "titre":       "Mise à jour antivirus",
                "description": "Antivirus expiré sur 3 postes.",
                "priorite":    3,
                "type":        2
            },
        ]

        create_plusieurs_tickets(session, tickets_a_creer)