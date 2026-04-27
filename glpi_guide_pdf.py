# glpi_guide_pdf.py — Guide d'installation GLPI CIMAT
# Projet CIMAT Béni Mellal

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, HRFlowable, PageBreak)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from datetime import datetime
import os

# ── Couleurs CIMAT ────────────────────────────────────────
BLEU_FONCE  = HexColor("#1F4E79")
BLEU_MOY    = HexColor("#2E75B6")
BLEU_CLAIR  = HexColor("#D6E4F0")
VERT        = HexColor("#1D6A2A")
VERT_CLAIR  = HexColor("#E2EFDA")
GRIS        = HexColor("#F2F2F2")
GRIS_FONCE  = HexColor("#595959")
ORANGE      = HexColor("#C55A11")

def creer_styles():
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="Titre1",
        fontSize=22, textColor=white, fontName="Helvetica-Bold",
        alignment=TA_CENTER, spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="Titre2",
        fontSize=16, textColor=BLEU_FONCE, fontName="Helvetica-Bold",
        spaceBefore=16, spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="Titre3",
        fontSize=13, textColor=BLEU_MOY, fontName="Helvetica-Bold",
        spaceBefore=10, spaceAfter=4
    ))
    styles.add(ParagraphStyle(
        name="Corps",
        fontSize=10, textColor=black, fontName="Helvetica",
        spaceAfter=6, leading=16, alignment=TA_JUSTIFY
    ))
    styles.add(ParagraphStyle(
        name="CorpsBlock",
        fontSize=9, textColor=HexColor("#1a1a1a"),
        fontName="Courier", backColor=HexColor("#F5F5F5"),
        spaceAfter=6, leading=14, leftIndent=12,
        borderPad=6
    ))
    styles.add(ParagraphStyle(
        name="Note",
        fontSize=9, textColor=ORANGE, fontName="Helvetica-Oblique",
        spaceAfter=4, leftIndent=12
    ))
    styles.add(ParagraphStyle(
        name="Puce",
        fontSize=10, textColor=black, fontName="Helvetica",
        spaceAfter=3, leftIndent=20, leading=14
    ))
    return styles

def page_de_garde(elements, styles):
    # Fond titre
    data = [[Paragraph(
        "<font color='white'><b>GUIDE D'INSTALLATION ET D'UTILISATION</b></font>",
        styles["Titre1"]
    )]]
    t = Table(data, colWidths=[17*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BLEU_FONCE),
        ("ROUNDEDCORNERS", [8]),
        ("TOPPADDING",    (0,0), (-1,-1), 18),
        ("BOTTOMPADDING", (0,0), (-1,-1), 18),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 0.4*cm))

    # Sous-titre
    data2 = [[Paragraph(
        "<font color='white'>Système GLPI 10.0.17 — Automatisation Python</font>",
        ParagraphStyle("st", fontSize=13, textColor=white,
                       fontName="Helvetica", alignment=TA_CENTER)
    )]]
    t2 = Table(data2, colWidths=[17*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BLEU_MOY),
        ("TOPPADDING",    (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
    ]))
    elements.append(t2)
    elements.append(Spacer(1, 1*cm))

    # Infos projet
    infos = [
        ["Projet",      "Système GLPI — Usine CIMAT"],
        ["Client",      "CIMAT Béni Mellal, Maroc"],
        ["Version",     "GLPI 10.0.17 + Python 3.14"],
        ["Auteur",      "Ahmed Daou"],
        ["Date",        datetime.now().strftime("%d/%m/%Y")],
        ["Statut",      "Production"],
    ]
    t3 = Table(infos, colWidths=[5*cm, 12*cm])
    t3.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (0,-1), BLEU_CLAIR),
        ("BACKGROUND",  (1,0), (1,-1), GRIS),
        ("FONTNAME",    (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 10),
        ("TEXTCOLOR",   (0,0), (0,-1), BLEU_FONCE),
        ("GRID",        (0,0), (-1,-1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING",  (0,0), (-1,-1), 7),
        ("BOTTOMPADDING",(0,0),(-1,-1), 7),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
    ]))
    elements.append(t3)
    elements.append(PageBreak())

def table_des_matieres(elements, styles):
    elements.append(Paragraph("Table des matières", styles["Titre2"]))
    elements.append(HRFlowable(width="100%", thickness=1,
                               color=BLEU_MOY, spaceAfter=8))
    chapitres = [
        ("1.", "Infrastructure et prérequis",          "3"),
        ("2.", "Installation GLPI 10.0.17",             "3"),
        ("3.", "Configuration initiale",                "4"),
        ("4.", "API REST — Activation et tokens",       "4"),
        ("5.", "Scripts Python — Vue d'ensemble",       "5"),
        ("6.", "Script 1 — Création automatique tickets","5"),
        ("7.", "Script 2 — Export Excel hebdomadaire",  "6"),
        ("8.", "Script 3 — Alertes email",              "6"),
        ("9.", "Script 4 — Inventaire réseau",          "7"),
        ("10.","Planification automatique (Windows)",   "7"),
        ("11.","Dépannage et erreurs courantes",        "8"),
    ]
    for num, titre, page in chapitres:
        elements.append(Paragraph(
            f"<b>{num}</b> &nbsp;&nbsp; {titre} "
            f"<font color='#888888'>{'.' * (55 - len(titre))} {page}</font>",
            styles["Corps"]
        ))
    elements.append(PageBreak())

def section_infrastructure(elements, styles):
    elements.append(Paragraph("1. Infrastructure et prérequis", styles["Titre2"]))
    elements.append(HRFlowable(width="100%", thickness=1,
                               color=BLEU_MOY, spaceAfter=8))

    elements.append(Paragraph("Environnement technique", styles["Titre3"]))

    data = [
        ["Composant",      "Détail",                    "Statut"],
        ["Serveur VM",     "Ubuntu 24 — VirtualBox",    " Opérationnel"],
        ["GLPI",           "Version 10.0.17",           " Installé"],
        ["IP Serveur",     "192.168.112.129",           " Configurée"],
        ["HTTPS",          "Certificat SSL actif",      " Activé"],
        ["API REST",       "Activée + tokens générés",  " Fonctionnelle"],
        ["Python",         "Version 3.14 — Windows",    " Installé"],
        ["Plugin Inventory","GLPI Inventory",           " Installé"],
    ]
    t = Table(data, colWidths=[5*cm, 8*cm, 4*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), BLEU_FONCE),
        ("TEXTCOLOR",    (0,0), (-1,0), white),
        ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 9),
        ("BACKGROUND",   (0,1), (-1,-1), GRIS),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [white, GRIS]),
        ("GRID",         (0,0), (-1,-1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING",   (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 0.3*cm))

    elements.append(Paragraph("Bibliothèques Python requises", styles["Titre3"]))
    elements.append(Paragraph(
        "pip install requests openpyxl reportlab",
        styles["Corps"]
    ))

def section_api(elements, styles):
    elements.append(Spacer(1, 0.3*cm))
    elements.append(Paragraph(
        "4. API REST — Activation et tokens", styles["Titre2"]))
    elements.append(HRFlowable(width="100%", thickness=1,
                               color=BLEU_MOY, spaceAfter=8))

    elements.append(Paragraph(
        "Activer l'API REST dans GLPI :", styles["Titre3"]))
    steps = [
        "Administration → Configuration générale → API",
        "Activer l'API REST → Oui",
        "Activer la connexion avec credentials → Oui",
        "Sauvegarder",
    ]
    for i, s in enumerate(steps, 1):
        elements.append(Paragraph(f"  {i}. {s}", styles["Puce"]))

    elements.append(Spacer(1, 0.3*cm))
    elements.append(Paragraph("Tokens de connexion :", styles["Titre3"]))

    data = [
        ["Token",        "Valeur",                              "Usage"],
        ["APP_TOKEN",    "Qzqd9JxK0k9vcF9daXxT4zIJKZGht2r20…", "Identifie l'application"],
        ["USER_TOKEN",   "gMJqxaSdpy79c9HQout8reOo1PBY0UfM…",  "Authentifie l'utilisateur"],
    ]
    t = Table(data, colWidths=[3.5*cm, 9*cm, 4.5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), BLEU_FONCE),
        ("TEXTCOLOR",    (0,0), (-1,0), white),
        ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [white, GRIS]),
        ("GRID",         (0,0), (-1,-1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING",   (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("FONTNAME",     (0,1), (1,-1), "Courier"),
    ]))
    elements.append(t)

def section_scripts(elements, styles):
    elements.append(PageBreak())
    elements.append(Paragraph(
        "5. Scripts Python — Vue d'ensemble", styles["Titre2"]))
    elements.append(HRFlowable(width="100%", thickness=1,
                               color=BLEU_MOY, spaceAfter=8))

    data = [
        ["#", "Fichier",                  "Fonction",                        "Commande"],
        ["1", "glpi_tickets.py",          "Création automatique de tickets", "python glpi_tickets.py"],
        ["2", "glpi_export_excel.py",     "Export Excel hebdomadaire",       "python glpi_export_excel.py"],
        ["3", "glpi_alertes_email.py",    "Alertes tickets non résolus",     "python glpi_alertes_email.py"],
        ["4", "glpi_inventaire.py",       "Inventaire enrichi réseau",       "python glpi_inventaire.py"],
    ]
    t = Table(data, colWidths=[0.8*cm, 4.5*cm, 6.5*cm, 5.2*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), BLEU_FONCE),
        ("TEXTCOLOR",    (0,0), (-1,0), white),
        ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[VERT_CLAIR, white]),
        ("GRID",         (0,0), (-1,-1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING",   (0,0), (-1,-1), 7),
        ("BOTTOMPADDING",(0,0), (-1,-1), 7),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("FONTNAME",     (0,1), (1,-1), "Courier"),
    ]))
    elements.append(t)

    # Scripts détails
    scripts_info = [
        ("6. Script 1 — Création automatique de tickets",
         "Permet de créer des tickets GLPI automatiquement via l'API REST Python. "
         "Supporte la création unitaire et en lot avec gestion des priorités.",
         ["connect_glpi() — Ouvre la session API",
          "create_ticket() — Crée un ticket unique",
          "create_plusieurs_tickets() — Crée plusieurs tickets depuis une liste"]),

        ("7. Script 2 — Export Excel hebdomadaire",
         "Récupère tous les tickets GLPI et génère un rapport Excel formaté "
         "avec 2 feuilles : liste complète et statistiques résumées.",
         ["get_tickets() — Récupère tous les tickets",
          "export_excel() — Génère le fichier .xlsx formaté",
          "Feuille 1 : Liste tickets colorisée par statut/priorité",
          "Feuille 2 : Statistiques par statut et priorité"]),

        ("8. Script 3 — Alertes email",
         "Envoie automatiquement un email HTML aux responsables listant "
         "tous les tickets non résolus depuis plus de X heures.",
         ["get_tickets_non_resolus() — Filtre les tickets en retard",
          "construire_email_html() — Génère l'email HTML formaté",
          "envoyer_email() — Envoi via SMTP Gmail (App Password)"]),

        ("9. Script 4 — Inventaire réseau",
         "Récupère l'inventaire matériel depuis GLPI et l'enrichit avec "
         "des données réseau réelles : ping, résolution hostname.",
         ["get_computers() — Liste les ordinateurs GLPI",
          "ping_host() — Teste la disponibilité réseau",
          "resolve_hostname() — Résout IP → hostname",
          "export_inventaire_excel() — Export Excel avec statut réseau"]),
    ]

    for titre, desc, fonctions in scripts_info:
        elements.append(Spacer(1, 0.3*cm))
        elements.append(Paragraph(titre, styles["Titre3"]))
        elements.append(Paragraph(desc, styles["Corps"]))
        for f in fonctions:
            elements.append(Paragraph(f"  • {f}", styles["Puce"]))

def section_planification(elements, styles):
    elements.append(PageBreak())
    elements.append(Paragraph(
        "10. Planification automatique — Windows Task Scheduler",
        styles["Titre2"]))
    elements.append(HRFlowable(width="100%", thickness=1,
                               color=BLEU_MOY, spaceAfter=8))

    elements.append(Paragraph(
        "Pour automatiser l'exécution des scripts sans intervention manuelle :",
        styles["Corps"]))

    elements.append(Paragraph("Ouvrir Task Scheduler :", styles["Titre3"]))
    steps = [
        "Windows + R → taskschd.msc → Entrée",
        "Action → Créer une tâche de base",
        "Nom : GLPI_Export_Excel  |  Déclencheur : Hebdomadaire",
        "Action : Démarrer un programme",
        "Programme : C:\\Users\\pc\\AppData\\Local\\Programs\\Python\\Python314\\python.exe",
        "Arguments : C:\\Users\\pc\\Desktop\\GLPI_Python\\glpi_export_excel.py",
    ]
    for i, s in enumerate(steps, 1):
        elements.append(Paragraph(f"  {i}. {s}", styles["Puce"]))

    elements.append(Spacer(1, 0.3*cm))
    elements.append(Paragraph("Planification recommandée :", styles["Titre3"]))
    data = [
        ["Script",               "Fréquence",     "Heure"],
        ["glpi_export_excel.py", "Chaque lundi",  "08:00"],
        ["glpi_alertes_email.py","Chaque jour",   "09:00"],
        ["glpi_inventaire.py",   "Chaque semaine","07:00"],
    ]
    t = Table(data, colWidths=[7*cm, 5*cm, 5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), BLEU_FONCE),
        ("TEXTCOLOR",    (0,0), (-1,0), white),
        ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 10),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[white, GRIS]),
        ("GRID",         (0,0), (-1,-1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING",   (0,0), (-1,-1), 7),
        ("BOTTOMPADDING",(0,0), (-1,-1), 7),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
    ]))
    elements.append(t)

def section_depannage(elements, styles):
    elements.append(Spacer(1, 0.4*cm))
    elements.append(Paragraph(
        "11. Dépannage et erreurs courantes", styles["Titre2"]))
    elements.append(HRFlowable(width="100%", thickness=1,
                               color=BLEU_MOY, spaceAfter=8))

    erreurs = [
        ("TypeError: list indices must be integers",
         "Token expiré ou incorrect.",
         "Régénérer APP_TOKEN et USER_TOKEN dans Administration → API"),
        ("ConnectionRefusedError",
         "Serveur GLPI inaccessible.",
         "Vérifier que la VM Ubuntu est démarrée et GLPI actif"),
        ("SMTPAuthenticationError",
         "App Password Gmail incorrect.",
         "Régénérer le App Password sur myaccount.google.com/apppasswords"),
        ("Aucun ordinateur trouvé",
         "Parc GLPI vide.",
         "Ajouter des équipements dans Parc → Ordinateurs"),
        ("Status Corps 401",
         "Non autorisé — token invalide.",
         "Vérifier APP_TOKEN et USER_TOKEN dans le script"),
    ]

    for erreur, cause, solution in erreurs:
        data = [
            [Paragraph(f" {erreur}", ParagraphStyle(
                "err", fontSize=9, fontName="Courier",
                textColor=HexColor("#C00000")))],
            [Paragraph(f"Cause : {cause}", styles["Note"])],
            [Paragraph(f" Solution : {solution}", styles["Corps"])],
        ]
        t = Table(data, colWidths=[17*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0), (-1,0), HexColor("#FFF2F2")),
            ("BACKGROUND",   (0,2), (-1,2), VERT_CLAIR),
            ("GRID",         (0,0), (-1,-1), 0.5, HexColor("#DDDDDD")),
            ("TOPPADDING",   (0,0), (-1,-1), 5),
            ("BOTTOMPADDING",(0,0), (-1,-1), 5),
            ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 0.2*cm))

# ── MAIN ──────────────────────────────────────────────────
def generer_guide():
    nom_fichier = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        f"Guide_GLPI_CIMAT_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    )
    doc = SimpleDocTemplate(
        nom_fichier,
        pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm,   bottomMargin=2*cm,
        title="Guide GLPI CIMAT",
        author="Ahmed Daou"
    )
    styles   = creer_styles()
    elements = []

    print(" Génération du guide PDF GLPI CIMAT…")
    page_de_garde(elements, styles)
    table_des_matieres(elements, styles)
    section_infrastructure(elements, styles)
    section_api(elements, styles)
    section_scripts(elements, styles)
    section_planification(elements, styles)
    section_depannage(elements, styles)

    doc.build(elements)
    print(f" Guide PDF créé : {nom_fichier}")
    return nom_fichier

if __name__ == "__main__":
    generer_guide()