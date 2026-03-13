"""
Agent IA — Décorations Noël
Connexion Google Sheets + génération de dossier PDF

Structure de fichiers attendue :
  agent_deco_noel/
  ├── agent.py              ← ce fichier
  ├── credentials.json      ← votre clé Google (NE PAS PARTAGER)
  ├── requirements.txt
  └── dossiers_generes/     ← PDFs générés ici (créé automatiquement)

Usage :
  python agent.py
"""

import os, json, re
from datetime import datetime

# ── Dépendances ───────────────────────────────────────────────────────────────
try:
    import gspread
    from google.oauth2.service_account import Credentials
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, PageBreak,
                                     KeepTogether, Flowable)
    from reportlab.pdfgen import canvas as rl_canvas
except ImportError:
    print("Installation des dépendances...")
    os.system("pip install gspread google-auth reportlab --break-system-packages -q")
    import gspread
    from google.oauth2.service_account import Credentials
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, PageBreak,
                                     KeepTogether, Flowable)
    from reportlab.pdfgen import canvas as rl_canvas

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — modifier ces valeurs
# ════════════════════════════════════════════════════════════════════════════
CREDENTIALS_FILE = "credentials.json"          # chemin vers votre fichier JSON
SPREADSHEET_ID   = "1jUqtAX65_CCSKtyjwsYFNrdb_AdvFJHur2RbvCmLe4M"  # ID de votre Google Sheets
SHEET_DECOS      = "DÉCORATIONS"               # nom exact de l'onglet décorations
SHEET_MODULES    = "MODULES"                   # nom exact de l'onglet modules
OUTPUT_DIR       = "dossiers_generes"          # dossier de sortie des PDFs

# ════════════════════════════════════════════════════════════════════════════
# CONNEXION GOOGLE SHEETS
# ════════════════════════════════════════════════════════════════════════════
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

def connect_sheets():
    """Connexion à Google Sheets via le compte de service."""
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def load_stock(spreadsheet):
    """
    Charge les décorations et modules depuis Google Sheets.
    Retourne une liste de dicts prêts à l'emploi.
    """
    # ── Charger les décorations ──────────────────────────────────────────────
    ws_dec = spreadsheet.worksheet(SHEET_DECOS)
    dec_rows = ws_dec.get_all_records()

    # ── Charger les modules ──────────────────────────────────────────────────
    ws_mod = spreadsheet.worksheet(SHEET_MODULES)
    mod_rows = ws_mod.get_all_records()

    # Indexer les modules par ID décoration parente
    modules_by_dec = {}
    for m in mod_rows:
        dec_id = str(m.get("ID_Décoration parente", "")).strip()
        if dec_id:
            modules_by_dec.setdefault(dec_id, []).append({
                "type": str(m.get("Type de Module", "")).strip(),
                "nom":  str(m.get("Nom du Module", "")).strip(),
                "hauteur": _to_int(m.get("Hauteur (cm)", 0)),
                "largeur": _to_int(m.get("Largeur (cm)", 0)),
                "qty":  _to_int(m.get("Qté en Stock", 0)),
                "dim":  f"{m.get('Largeur (cm)','')} x {m.get('Hauteur (cm)','')} cm",
                "desc": str(m.get("Description Module", "")).strip(),
            })

    # Assembler les décorations avec leurs modules
    decorations = []
    for d in dec_rows:
        dec_id = str(d.get("ID_Décoration", "")).strip()
        if not dec_id or str(d.get("Actif", "")).strip().upper() != "OUI":
            continue

        mods = modules_by_dec.get(dec_id, [])
        qty  = min((m["qty"] for m in mods), default=0) if mods else 0
        hmax = max((m["hauteur"] for m in mods), default=0) if mods else 0
        lmax = max((m["largeur"] for m in mods), default=0) if mods else 0

        decorations.append({
            "id":       dec_id,
            "nom":      str(d.get("Nom de la Décoration", "")).strip(),
            "style":    str(d.get("Style", "")).strip(),
            "couleur1": str(d.get("Couleur Principale", "")).strip(),
            "couleur2": str(d.get("Couleur Secondaire", "")).strip(),
            "tags":     str(d.get("Tags / Thèmes", "")).strip().lower(),
            "desc":     str(d.get("Description Courte", "")).strip(),
            "photo_url":str(d.get("URL Photo assemblée", "")).strip(),
            "qty":      qty,
            "hauteur_max": hmax,
            "largeur_max": lmax,
            "modules":  mods,
        })

    print(f"  → {len(decorations)} décorations actives chargées")
    print(f"  → {sum(len(m) for m in modules_by_dec.values())} modules au total")
    return decorations

def _to_int(val):
    try:    return int(str(val).strip())
    except: return 0

# ════════════════════════════════════════════════════════════════════════════
# MOTEUR DE SÉLECTION
# ════════════════════════════════════════════════════════════════════════════
def score_decoration(d, params):
    """
    Score une décoration selon les critères de l'utilisateur.
    Retourne un entier (plus élevé = meilleur match).
    """
    s = 0
    # Disponibilité obligatoire
    if d["qty"] <= 0:
        return -1

    # Style (critère fort)
    if params.get("style") and d["style"] == params["style"]:
        s += 40

    # Couleur principale (critère fort)
    c = params.get("couleur", "")
    if c and (d["couleur1"] == c or d["couleur2"] == c):
        s += 30

    # Hauteur sous plafond (contrainte technique)
    if params.get("hauteur"):
        if d["hauteur_max"] <= int(params["hauteur"]):
            s += 15
        else:
            s -= 30  # pénalité si trop grand

    # Tags / mots-clés libres
    if params.get("notes"):
        notes_lower = params["notes"].lower()
        for tag in d["tags"].split(","):
            if tag.strip() and tag.strip() in notes_lower:
                s += 5

    # Bonus disponibilité (plusieurs exemplaires = plus fiable)
    if d["qty"] >= 3: s += 10
    elif d["qty"] >= 2: s += 5

    return s

def select_decorations(stock, params):
    """Sélectionne et trie les décorations selon les critères."""
    nb = int(params.get("nb", 3))
    scored = [
        {**d, "score": score_decoration(d, params)}
        for d in stock
    ]
    scored = [d for d in scored if d["score"] >= 0]
    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:nb]

def coherence_score(decos):
    """Calcule un score de cohérence 0-100 entre les décorations sélectionnées."""
    if len(decos) <= 1:
        return 100
    matches, total = 0, 0
    for i in range(len(decos)):
        for j in range(i+1, len(decos)):
            total += 1
            a, b = decos[i], decos[j]
            if (a["style"] == b["style"] or
                a["couleur1"] == b["couleur1"] or
                a["couleur1"] == b["couleur2"] or
                a["couleur2"] == b["couleur1"]):
                matches += 1
    return round((matches / total) * 100)

# ════════════════════════════════════════════════════════════════════════════
# GÉNÉRATION PDF
# ════════════════════════════════════════════════════════════════════════════
W, H = A4
C_DARK_GREEN = colors.HexColor("#1A3A2A")
C_BLUE       = colors.HexColor("#185FA5")
C_BLUE_LIGHT = colors.HexColor("#E6F1FB")
C_GREY_BG    = colors.HexColor("#F9F4EE")
C_GREY_TEXT  = colors.HexColor("#5F5E5A")
C_BORDER     = colors.HexColor("#CCCCCC")
C_WHITE      = colors.white
C_BLACK      = colors.HexColor("#2C2C2A")

def s(name, **kw):
    return ParagraphStyle(name, **kw)

ST = {
    "title":    s("t",  fontName="Helvetica-Bold", fontSize=18, textColor=C_DARK_GREEN, leading=24, spaceAfter=4),
    "section":  s("s",  fontName="Helvetica-Bold", fontSize=12, textColor=C_DARK_GREEN, leading=16, spaceBefore=12, spaceAfter=6),
    "body":     s("b",  fontName="Helvetica",      fontSize=10, textColor=C_BLACK,      leading=16, spaceAfter=4),
    "small":    s("sm", fontName="Helvetica",      fontSize=9,  textColor=C_GREY_TEXT,  leading=13),
    "label":    s("lb", fontName="Helvetica-Bold", fontSize=9,  textColor=C_GREY_TEXT,  leading=12),
    "card_ttl": s("ct", fontName="Helvetica-Bold", fontSize=14, textColor=C_WHITE,      leading=18),
    "card_id":  s("ci", fontName="Helvetica",      fontSize=10, textColor=colors.HexColor("#9FE1CB"), leading=14),
}

class CoherenceBar(Flowable):
    def __init__(self, pct, width=15*cm):
        self.pct, self.width = pct, width
    def draw(self):
        self.canv.setFillColor(C_BORDER)
        self.canv.roundRect(0, 0, self.width, 8, 4, fill=1, stroke=0)
        self.canv.setFillColor(C_DARK_GREEN)
        fw = max(8, self.width * self.pct / 100)
        self.canv.roundRect(0, 0, fw, 8, 4, fill=1, stroke=0)
    def wrap(self, *a): return self.width, 8

class PhotoBlock(Flowable):
    def __init__(self, url, dec_id, w=15*cm, h=5*cm):
        self.url, self.dec_id, self.w, self.h = url, dec_id, w, h
    def draw(self):
        c = self.canv
        c.setFillColor(C_GREY_BG)
        c.setStrokeColor(C_BORDER)
        c.setLineWidth(0.5)
        c.roundRect(0, 0, self.w, self.h, 6, fill=1, stroke=1)
        c.setFillColor(C_GREY_TEXT)
        c.setFont("Helvetica", 9)
        if self.url and self.url.startswith("http"):
            c.drawCentredString(self.w/2, self.h/2+4, f"Photo : {self.dec_id}")
            c.setFont("Helvetica", 8)
            c.setFillColor(colors.HexColor("#185FA5"))
            # Truncate URL for display
            display_url = self.url[:60] + "..." if len(self.url) > 60 else self.url
            c.drawCentredString(self.w/2, self.h/2-12, display_url)
        else:
            c.drawCentredString(self.w/2, self.h/2, f"Photo {self.dec_id} — URL Drive à renseigner")
    def wrap(self, *a): return self.w, self.h

class NumberedCanvas(rl_canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved = []
    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()
    def save(self):
        n = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            pg = self._pageNumber
            if pg > 1:
                self.setFont("Helvetica", 8)
                self.setFillColor(C_GREY_TEXT)
                self.drawCentredString(W/2, 18*cm/10*3,
                    f"Dossier Décorations Noël  —  Page {pg} / {n}")
                self.setStrokeColor(C_BORDER)
                self.setLineWidth(0.5)
                self.line(2*cm, 2.4*cm, W-2*cm, 2.4*cm)
            super().showPage()
        super().save()

def draw_cover(c, doc, params, decos, coh):
    c.saveState()
    c.setFillColor(C_DARK_GREEN)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_BLUE)
    c.rect(0, H-10*cm/10*12, W, 10*cm/10*12, fill=1, stroke=0)
    c.rect(0, 0, 8*cm/10, H, fill=1, stroke=0)
    # Decorative circles
    c.setFillColor(colors.HexColor("#FFFFFF08"))
    for x,y,r in [(W*.78,H*.55,110),(W*.88,H*.3,70),(W*.65,H*.72,55)]:
        c.circle(x,y,r,fill=1,stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 30)
    c.drawString(3*cm, H-6*cm, "Dossier de Présentation")
    c.setFont("Helvetica-Bold", 24)
    c.drawString(3*cm, H-8*cm, "Décorations de Noël")
    c.setFillColor(C_BLUE)
    c.rect(3*cm, H-8.8*cm, 7*cm, 3*cm/10, fill=1, stroke=0)
    # Theme badge
    style_label = params.get("style","Tous styles")
    couleur_label = params.get("couleur","Toutes couleurs")
    badge_text = f"Thème : {style_label} / {couleur_label}"
    c.setFillColor(C_BLUE)
    c.roundRect(3*cm, H-11*cm, len(badge_text)*0.18*cm+1*cm, 1*cm, 5, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(3.3*cm, H-10.55*cm, badge_text)
    # Stats
    coh_label = "Excellente" if coh>=80 else "Bonne" if coh>=55 else "Partielle"
    nb_modules = sum(len(d["modules"]) for d in decos)
    c.setFont("Helvetica", 11)
    c.setFillColor(colors.HexColor("#9FE1CB"))
    lines = [
        f"Sélection personnalisée — {len(decos)} décoration(s)",
        f"Cohérence de l'ensemble : {coh_label} ({coh}%)",
        f"Modules inclus : {nb_modules} modules au total",
        f"Généré le : {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
    ]
    for i, line in enumerate(lines):
        c.drawString(3*cm, H-13.5*cm - i*1.2*cm, line)
    c.setStrokeColor(colors.HexColor("#FFFFFF25"))
    c.setLineWidth(0.5)
    c.line(3*cm, H*.26, W-3*cm, H*.26)
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.HexColor("#9FE1CB"))
    c.drawString(3*cm, H*.22, "Document généré par l'agent IA — Usage commercial interne")
    c.restoreState()

def generate_pdf(decos, params, output_path):
    """Génère le PDF complet du dossier."""
    coh = coherence_score(decos)
    coh_label = "Excellente" if coh >= 80 else "Bonne" if coh >= 55 else "Partielle"

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=2.2*cm, rightMargin=2.2*cm,
        topMargin=2.5*cm, bottomMargin=3*cm,
    )

    story = [PageBreak()]  # Page 1 = cover dessinée manuellement

    # ── Page 2 : Synthèse ────────────────────────────────────────────────────
    story.append(Paragraph("Synthèse de la sélection", ST["title"]))
    story.append(HRFlowable(width="100%", thickness=0.5, color=C_BORDER, spaceAfter=10))

    summary = [
        [Paragraph("<b>Critère</b>", ST["label"]),  Paragraph("<b>Valeur</b>", ST["label"])],
        [Paragraph("Style",          ST["small"]),   Paragraph(params.get("style","Non précisé"), ST["body"])],
        [Paragraph("Couleur",        ST["small"]),   Paragraph(params.get("couleur","Non précisé"), ST["body"])],
        [Paragraph("Taille du site", ST["small"]),   Paragraph(params.get("taille","Non précisé"), ST["body"])],
        [Paragraph("Hauteur plafond",ST["small"]),   Paragraph(f"{params.get('hauteur','-')} cm" if params.get("hauteur") else "Non précisé", ST["body"])],
        [Paragraph("Budget",         ST["small"]),   Paragraph(f"{params.get('budget','Non précisé')} €" if params.get("budget") else "Non précisé", ST["body"])],
        [Paragraph("Décorations",    ST["small"]),   Paragraph(str(len(decos)), ST["body"])],
        [Paragraph("Cohérence",      ST["small"]),   Paragraph(f"{coh_label} ({coh}%)", ST["body"])],
    ]
    t = Table(summary, colWidths=[6*cm, 9.1*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,0), C_DARK_GREEN),
        ("TEXTCOLOR",  (0,0),(-1,0), C_WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[C_GREY_BG, C_WHITE]),
        ("GRID",       (0,0),(-1,-1), 0.5, C_BORDER),
        ("TOPPADDING", (0,0),(-1,-1), 7),
        ("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("LEFTPADDING",(0,0),(-1,-1),10),
        ("VALIGN",     (0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(t)
    story.append(Spacer(1, 14))
    story.append(Paragraph("Cohérence visuelle de l'ensemble", ST["section"]))
    story.append(CoherenceBar(coh))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "L'indice de cohérence mesure l'harmonie entre les décorations sélectionnées "
        "(style, couleurs, ambiance). Un score élevé garantit un rendu visuel homogène "
        "sur l'ensemble du site.", ST["body"]))

    if params.get("notes"):
        story.append(Spacer(1, 8))
        story.append(Paragraph("Notes du projet", ST["section"]))
        story.append(Paragraph(params["notes"], ST["body"]))

    story.append(PageBreak())

    # ── Fiches décorations ────────────────────────────────────────────────────
    for i, d in enumerate(decos):
        # En-tête coloré
        hdr = Table([[
            Paragraph(d["nom"], ST["card_ttl"]),
            Paragraph(d["id"],  ST["card_id"]),
        ]], colWidths=[12*cm, 3.1*cm])
        hdr.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), C_BLUE if d["style"] in ["Moderne","Luxe"] else C_DARK_GREEN),
            ("TOPPADDING",    (0,0),(-1,-1), 14),
            ("BOTTOMPADDING", (0,0),(-1,-1), 14),
            ("LEFTPADDING",   (0,0),(0,-1),  14),
            ("RIGHTPADDING",  (-1,0),(-1,-1),14),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
        ]))
        story.append(KeepTogether([hdr]))
        story.append(Spacer(1, 10))

        # Photo
        story.append(PhotoBlock(d["photo_url"], d["id"]))
        story.append(Spacer(1, 12))

        # Méta
        meta = [[
            Paragraph(f"<b>Style</b><br/>{d['style']}", ST["body"]),
            Paragraph(f"<b>Couleurs</b><br/>{d['couleur1']} / {d['couleur2']}", ST["body"]),
            Paragraph(f"<b>Disponible</b><br/>{d['qty']} ex.", ST["body"]),
            Paragraph(f"<b>Haut. max</b><br/>{d['hauteur_max']} cm", ST["body"]),
            Paragraph(f"<b>Larg. max</b><br/>{d['largeur_max']} cm", ST["body"]),
        ]]
        mt = Table(meta, colWidths=[3.02*cm]*5)
        mt.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), C_GREY_BG),
            ("GRID",         (0,0),(-1,-1), 0.5, C_BORDER),
            ("TOPPADDING",   (0,0),(-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 8),
            ("LEFTPADDING",  (0,0),(-1,-1), 8),
            ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ]))
        story.append(mt)
        story.append(Spacer(1, 12))

        # Description
        story.append(Paragraph("Description", ST["section"]))
        story.append(Paragraph(d["desc"] or "Description à renseigner dans Google Sheets.", ST["body"]))
        story.append(Spacer(1, 8))

        # Modules
        story.append(Paragraph(f"Composition — {len(d['modules'])} module(s)", ST["section"]))
        mod_rows = [[
            Paragraph("<b>Type</b>",        ST["label"]),
            Paragraph("<b>Nom du module</b>",ST["label"]),
            Paragraph("<b>Dimensions</b>",  ST["label"]),
            Paragraph("<b>Qté</b>",         ST["label"]),
        ]]
        for m in d["modules"]:
            mod_rows.append([
                Paragraph(m["type"], ST["small"]),
                Paragraph(m["nom"],  ST["body"]),
                Paragraph(m["dim"],  ST["small"]),
                Paragraph(str(m["qty"]), ST["body"]),
            ])
        mt2 = Table(mod_rows, colWidths=[2.8*cm, 7.3*cm, 3.5*cm, 1.5*cm])
        mt2.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0), C_DARK_GREEN),
            ("TEXTCOLOR",     (0,0),(-1,0), C_WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[C_GREY_BG, C_WHITE]),
            ("GRID",          (0,0),(-1,-1), 0.5, C_BORDER),
            ("TOPPADDING",    (0,0),(-1,-1), 7),
            ("BOTTOMPADDING", (0,0),(-1,-1), 7),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("ALIGN",         (3,0),(3,-1),  "CENTER"),
        ]))
        story.append(mt2)

        if d["tags"]:
            story.append(Spacer(1, 8))
            story.append(Paragraph(f"<b>Mots-clés :</b> {d['tags']}", ST["small"]))

        if i < len(decos) - 1:
            story.append(PageBreak())

    def _cover(c, doc):
        draw_cover(c, doc, params, decos, coh)
    def _later(c, doc):
        pass

    doc.build(story, onFirstPage=_cover, onLaterPages=_later,
              canvasmaker=NumberedCanvas)
    print(f"  → PDF généré : {output_path}")

# ════════════════════════════════════════════════════════════════════════════
# INTERFACE EN LIGNE DE COMMANDE
# ════════════════════════════════════════════════════════════════════════════
def ask(prompt, choices=None, default=""):
    """Pose une question à l'utilisateur en ligne de commande."""
    if choices:
        print(f"\n{prompt}")
        for i, c in enumerate(choices, 1):
            print(f"  {i}. {c}")
        val = input(f"  Votre choix (1-{len(choices)}, Entrée pour ignorer) : ").strip()
        if val.isdigit() and 1 <= int(val) <= len(choices):
            return choices[int(val)-1]
        return default
    else:
        val = input(f"\n{prompt} : ").strip()
        return val or default

def main():
    print("\n" + "="*55)
    print("  AGENT IA — DÉCORATIONS NOËL")
    print("="*55)

    # Connexion
    print("\nConnexion à Google Sheets...")
    try:
        spreadsheet = connect_sheets()
        print(f"  → Connecté à : {spreadsheet.title}")
    except FileNotFoundError:
        print(f"\nERREUR : Fichier '{CREDENTIALS_FILE}' introuvable.")
        print("Placez votre fichier credentials.json dans le même dossier que agent.py")
        return
    except Exception as e:
        print(f"\nERREUR de connexion : {e}")
        return

    # Chargement du stock
    print("\nChargement du stock...")
    try:
        stock = load_stock(spreadsheet)
    except Exception as e:
        print(f"\nERREUR lors du chargement : {e}")
        print("Vérifiez que les onglets 'DÉCORATIONS' et 'MODULES' existent dans votre Sheets.")
        return

    if not stock:
        print("\nAucune décoration active trouvée dans le stock.")
        return

    # Collecte des critères
    print("\n" + "-"*55)
    print("  CRITÈRES DU PROJET")
    print("-"*55)

    params = {}
    params["style"]   = ask("Style souhaité",
                            ["Traditionnel","Moderne","Nature","Luxe","Enfantin"])
    params["couleur"] = ask("Couleur principale",
                            ["Rouge","Or","Argent","Blanc","Bleu","Vert","Naturel"])
    params["taille"]  = ask("Taille du site",
                            ["Petit","Moyen","Grand"])
    params["hauteur"] = ask("Hauteur sous plafond en cm (ex: 350)")
    params["budget"]  = ask("Budget indicatif en € (ex: 5000)")
    nb_str            = ask("Nombre de décorations souhaitées (ex: 3)")
    params["nb"]      = int(nb_str) if nb_str.isdigit() else 3
    params["notes"]   = ask("Informations complémentaires (optionnel)")

    # Sélection
    print("\n" + "-"*55)
    print("  SÉLECTION EN COURS...")
    print("-"*55)
    selected = select_decorations(stock, params)

    if not selected:
        print("\nAucune décoration ne correspond aux critères.")
        return

    coh = coherence_score(selected)
    coh_label = "Excellente" if coh >= 80 else "Bonne" if coh >= 55 else "Partielle"

    print(f"\n{len(selected)} décoration(s) sélectionnée(s) :")
    for d in selected:
        print(f"  • {d['id']} — {d['nom']} (score: {d['score']}, stock: {d['qty']})")
    print(f"\nCohérence de l'ensemble : {coh_label} ({coh}%)")

    # Génération PDF
    confirm = input("\nGénérer le dossier PDF ? (O/n) : ").strip().lower()
    if confirm in ("", "o", "oui", "y", "yes"):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        style_slug = re.sub(r'[^a-zA-Z0-9]', '_', params.get("style","") or "Tous")
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"Dossier_Noel_{style_slug}_{ts}.pdf"
        output_path = os.path.join(OUTPUT_DIR, filename)

        print("\nGénération du PDF...")
        generate_pdf(selected, params, output_path)
        print(f"\nDossier disponible dans : {OUTPUT_DIR}/{filename}")
    else:
        print("\nGénération annulée.")

    print("\n" + "="*55 + "\n")

if __name__ == "__main__":
    main()
