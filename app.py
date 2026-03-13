"""
Agent IA — Décorations Noël
Interface web Streamlit
"""

import os, io, re
from datetime import datetime
import streamlit as st

# ── Config page ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Agent IA — Décorations Noël",
    page_icon="🎄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personnalisé ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1A3A2A 0%, #185FA5 100%);
        padding: 2rem 2rem 1.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p  { color: #9FE1CB; margin: 0.3rem 0 0; font-size: 0.95rem; }

    .deco-card {
        background: white;
        border: 1px solid #E0E0E0;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .deco-card-header {
        background: #1A3A2A;
        color: white;
        padding: 0.8rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
    }
    .badge-blue    { background: #E6F1FB; color: #0C447C; }
    .badge-green   { background: #EAF3DE; color: #27500A; }
    .badge-amber   { background: #FAEEDA; color: #633806; }
    .badge-red     { background: #FCEBEB; color: #791F1F; }

    .metric-row {
        display: flex;
        gap: 12px;
        margin: 0.8rem 0;
    }
    .metric-box {
        flex: 1;
        background: #F9F4EE;
        border: 1px solid #E8E0D5;
        border-radius: 8px;
        padding: 0.6rem 0.8rem;
        text-align: center;
    }
    .metric-box .label { font-size: 0.72rem; color: #888; margin-bottom: 2px; }
    .metric-box .value { font-size: 1rem; font-weight: 600; color: #1A3A2A; }

    .module-row {
        display: flex;
        align-items: center;
        padding: 5px 0;
        border-bottom: 1px solid #F0F0F0;
        font-size: 0.85rem;
        gap: 8px;
    }
    .module-row:last-child { border-bottom: none; }
    .module-type  { color: #888; min-width: 80px; }
    .module-name  { flex: 1; color: #333; }
    .module-dim   { color: #888; font-size: 0.8rem; }

    .stock-row-ok   { color: #27500A; }
    .stock-row-low  { color: #633806; }
    .stock-row-none { color: #791F1F; }

    .coherence-container {
        background: #F9F4EE;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .stDownloadButton > button {
        background: #1A3A2A !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        font-size: 1rem !important;
        width: 100%;
    }
    .stDownloadButton > button:hover { background: #185FA5 !important; }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# CONNEXION GOOGLE SHEETS (avec cache 5 minutes)
# ════════════════════════════════════════════════════════════════════════════
@st.cache_resource(ttl=300)
def get_spreadsheet():
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    # Sur Streamlit Cloud : credentials dans st.secrets
    # En local : fichier credentials.json
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    except Exception:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)

    client = gspread.authorize(creds)
    return client.open_by_key(st.secrets.get("SPREADSHEET_ID", os.getenv("SPREADSHEET_ID", "")))

@st.cache_data(ttl=300)
def load_stock():
    try:
        spreadsheet = get_spreadsheet()
    except Exception as e:
        st.error(f"Erreur de connexion Google Sheets : {e}")
        return []

    dec_rows = spreadsheet.worksheet("DÉCORATIONS").get_all_records()
    mod_rows = spreadsheet.worksheet("MODULES").get_all_records()

    modules_by_dec = {}
    for m in mod_rows:
        did = str(m.get("ID_Décoration parente", "")).strip()
        if did:
            modules_by_dec.setdefault(did, []).append({
                "type":    str(m.get("Type de Module", "")).strip(),
                "nom":     str(m.get("Nom du Module", "")).strip(),
                "hauteur": _int(m.get("Hauteur (cm)", 0)),
                "largeur": _int(m.get("Largeur (cm)", 0)),
                "qty":     _int(m.get("Qté en Stock", 0)),
                "dim":     f"{m.get('Largeur (cm)','')} x {m.get('Hauteur (cm)','')} cm",
                "desc":    str(m.get("Description Module", "")).strip(),
            })

    decorations = []
    for d in dec_rows:
        did = str(d.get("ID_Décoration", "")).strip()
        if not did or str(d.get("Actif", "")).strip().upper() != "OUI":
            continue
        mods = modules_by_dec.get(did, [])
        decorations.append({
            "id":          did,
            "nom":         str(d.get("Nom de la Décoration", "")).strip(),
            "style":       str(d.get("Style", "")).strip(),
            "couleur1":    str(d.get("Couleur Principale", "")).strip(),
            "couleur2":    str(d.get("Couleur Secondaire", "")).strip(),
            "tags":        str(d.get("Tags / Thèmes", "")).strip().lower(),
            "desc":        str(d.get("Description Courte", "")).strip(),
            "photo_url":   str(d.get("URL Photo assemblée", "")).strip(),
            "qty":         min((m["qty"] for m in mods), default=0) if mods else 0,
            "hauteur_max": max((m["hauteur"] for m in mods), default=0) if mods else 0,
            "largeur_max": max((m["largeur"] for m in mods), default=0) if mods else 0,
            "modules":     mods,
        })
    return decorations

def _int(v):
    try: return int(str(v).strip())
    except: return 0

# ════════════════════════════════════════════════════════════════════════════
# MOTEUR DE SÉLECTION
# ════════════════════════════════════════════════════════════════════════════
def score_deco(d, params):
    if d["qty"] <= 0: return -1
    s = 0
    if params["style"]   and d["style"]    == params["style"]:   s += 40
    if params["couleur"] and params["couleur"] in (d["couleur1"], d["couleur2"]): s += 30
    if params["hauteur"] and d["hauteur_max"] <= params["hauteur"]: s += 15
    elif not params["hauteur"]: s += 15
    if params["notes"]:
        nl = params["notes"].lower()
        for t in d["tags"].split(","):
            if t.strip() and t.strip() in nl: s += 5
    s += 10 if d["qty"] >= 3 else 5 if d["qty"] >= 2 else 0
    return s

def select_decos(stock, params):
    scored = [{**d, "score": score_deco(d, params)} for d in stock]
    scored = [d for d in scored if d["score"] >= 0]
    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:params["nb"]]

def coherence(decos):
    if len(decos) <= 1: return 100
    m, t = 0, 0
    for i in range(len(decos)):
        for j in range(i+1, len(decos)):
            t += 1
            a, b = decos[i], decos[j]
            if (a["style"] == b["style"] or
                a["couleur1"] == b["couleur1"] or
                a["couleur1"] == b["couleur2"] or
                a["couleur2"] == b["couleur1"]):
                m += 1
    return round((m/t)*100)

# ════════════════════════════════════════════════════════════════════════════
# GÉNÉRATION PDF
# ════════════════════════════════════════════════════════════════════════════
def generate_pdf(decos, params):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable,
                                     PageBreak, KeepTogether, Flowable)
    from reportlab.pdfgen import canvas as rl_canvas

    W, H = A4
    CDG  = rl_colors.HexColor("#1A3A2A")
    CBL  = rl_colors.HexColor("#185FA5")
    CBLL = rl_colors.HexColor("#E6F1FB")
    CGR  = rl_colors.HexColor("#F9F4EE")
    CGT  = rl_colors.HexColor("#5F5E5A")
    CBR  = rl_colors.HexColor("#CCCCCC")

    def ps(name, **kw): return ParagraphStyle(name, **kw)
    ST = {
        "title":   ps("t",  fontName="Helvetica-Bold", fontSize=16, textColor=CDG, leading=22, spaceAfter=4),
        "section": ps("s",  fontName="Helvetica-Bold", fontSize=11, textColor=CDG, leading=15, spaceBefore=10, spaceAfter=5),
        "body":    ps("b",  fontName="Helvetica",      fontSize=10, textColor=rl_colors.HexColor("#2C2C2A"), leading=15, spaceAfter=3),
        "small":   ps("sm", fontName="Helvetica",      fontSize=9,  textColor=CGT, leading=12),
        "label":   ps("lb", fontName="Helvetica-Bold", fontSize=9,  textColor=CGT, leading=12),
        "card_t":  ps("ct", fontName="Helvetica-Bold", fontSize=13, textColor=rl_colors.white, leading=17),
        "card_id": ps("ci", fontName="Helvetica",      fontSize=9,  textColor=rl_colors.HexColor("#9FE1CB"), leading=13),
    }

    class CohBar(Flowable):
        def __init__(self, pct, w=14*cm):
            self.pct, self.w = pct, w
        def draw(self):
            self.canv.setFillColor(CBR)
            self.canv.roundRect(0,0,self.w,8,4,fill=1,stroke=0)
            self.canv.setFillColor(CDG)
            self.canv.roundRect(0,0,max(8,self.w*self.pct/100),8,4,fill=1,stroke=0)
        def wrap(self,*a): return self.w, 8

    class PhotoBlock(Flowable):
        def __init__(self, url, did, w=15*cm, h=4.5*cm):
            self.url,self.did,self.w,self.h = url,did,w,h
        def draw(self):
            c = self.canv
            c.setFillColor(CGR); c.setStrokeColor(CBR); c.setLineWidth(0.5)
            c.roundRect(0,0,self.w,self.h,6,fill=1,stroke=1)
            c.setFillColor(CGT); c.setFont("Helvetica",9)
            if self.url and self.url.startswith("http"):
                c.drawCentredString(self.w/2,self.h/2+4,f"Photo : {self.did}")
                c.setFont("Helvetica",7); c.setFillColor(CBL)
                url_s = self.url[:65]+"..." if len(self.url)>65 else self.url
                c.drawCentredString(self.w/2,self.h/2-10,url_s)
            else:
                c.drawCentredString(self.w/2,self.h/2,f"Photo {self.did} — URL Drive à renseigner")
        def wrap(self,*a): return self.w,self.h

    class NumCanvas(rl_canvas.Canvas):
        def __init__(self,*a,**kw):
            super().__init__(*a,**kw); self._saved=[]
        def showPage(self):
            self._saved.append(dict(self.__dict__)); self._startPage()
        def save(self):
            n=len(self._saved)
            for state in self._saved:
                self.__dict__.update(state)
                pg=self._pageNumber
                if pg>1:
                    self.setFont("Helvetica",8); self.setFillColor(CGT)
                    self.drawCentredString(W/2,2.2*cm,f"Dossier Décorations Noël  —  Page {pg} / {n}")
                    self.setStrokeColor(CBR); self.setLineWidth(0.5)
                    self.line(2*cm,2.6*cm,W-2*cm,2.6*cm)
                super().showPage()
            super().save()

    coh = coherence(decos)
    coh_lbl = "Excellente" if coh>=80 else "Bonne" if coh>=55 else "Partielle"
    buf = io.BytesIO()

    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=2.2*cm, rightMargin=2.2*cm,
        topMargin=2.5*cm, bottomMargin=3.2*cm)

    def draw_cover(c, doc):
        c.saveState()
        c.setFillColor(CDG); c.rect(0,0,W,H,fill=1,stroke=0)
        c.setFillColor(CBL); c.rect(0,H-1.5*cm,W,1.5*cm,fill=1,stroke=0)
        c.rect(0,0,0.8*cm,H,fill=1,stroke=0)
        c.setFillColor(rl_colors.HexColor("#FFFFFF08"))
        for x,y,r in [(W*.78,H*.55,110),(W*.88,H*.32,72),(W*.63,H*.73,55)]:
            c.circle(x,y,r,fill=1,stroke=0)
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold",28); c.drawString(3*cm,H-5.5*cm,"Dossier de Présentation")
        c.setFont("Helvetica-Bold",22); c.drawString(3*cm,H-7.2*cm,"Décorations de Noël")
        c.setFillColor(CBL); c.rect(3*cm,H-7.9*cm,6*cm,0.25*cm,fill=1,stroke=0)
        sl = params.get("style","Tous") or "Tous"
        cl = params.get("couleur_label","Toutes") or "Toutes"
        badge = f"Thème : {sl} / {cl}"
        bw = len(badge)*0.175*cm+1*cm
        c.setFillColor(CBL); c.roundRect(3*cm,H-10.2*cm,bw,1*cm,5,fill=1,stroke=0)
        c.setFillColor(rl_colors.white); c.setFont("Helvetica-Bold",10)
        c.drawString(3.3*cm,H-9.75*cm,badge)
        nb_mod = sum(len(d["modules"]) for d in decos)
        c.setFont("Helvetica",11); c.setFillColor(rl_colors.HexColor("#9FE1CB"))
        for i,line in enumerate([
            f"Sélection personnalisée — {len(decos)} décoration(s)",
            f"Cohérence de l'ensemble : {coh_lbl} ({coh}%)",
            f"Modules inclus : {nb_mod} au total",
            f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
        ]):
            c.drawString(3*cm, H-13*cm - i*1.15*cm, line)
        c.setStrokeColor(rl_colors.HexColor("#FFFFFF20")); c.setLineWidth(0.5)
        c.line(3*cm,H*.26,W-3*cm,H*.26)
        c.setFont("Helvetica",9); c.setFillColor(rl_colors.HexColor("#9FE1CB"))
        c.drawString(3*cm,H*.22,"Document généré par l'agent IA — Usage commercial interne")
        c.restoreState()

    story = [PageBreak()]

    # Synthèse
    story.append(Paragraph("Synthèse de la sélection", ST["title"]))
    story.append(HRFlowable(width="100%",thickness=0.5,color=CBR,spaceAfter=10))
    rows = [
        [Paragraph("<b>Critère</b>",ST["label"]),  Paragraph("<b>Valeur</b>",ST["label"])],
        [Paragraph("Style",         ST["small"]),  Paragraph(params.get("style","—") or "—",ST["body"])],
        [Paragraph("Couleur",       ST["small"]),  Paragraph(params.get("couleur_label","—") or "—",ST["body"])],
        [Paragraph("Taille site",   ST["small"]),  Paragraph(params.get("taille","—") or "—",ST["body"])],
        [Paragraph("Haut. plafond", ST["small"]),  Paragraph(f"{params['hauteur']} cm" if params.get("hauteur") else "—",ST["body"])],
        [Paragraph("Budget",        ST["small"]),  Paragraph(f"{params['budget']} €" if params.get("budget") else "—",ST["body"])],
        [Paragraph("Décorations",   ST["small"]),  Paragraph(str(len(decos)),ST["body"])],
        [Paragraph("Cohérence",     ST["small"]),  Paragraph(f"{coh_lbl} ({coh}%)",ST["body"])],
    ]
    t = Table(rows, colWidths=[6*cm,9.1*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),CDG),("TEXTCOLOR",(0,0),(-1,0),rl_colors.white),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[CGR,rl_colors.white]),
        ("GRID",(0,0),(-1,-1),0.5,CBR),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("LEFTPADDING",(0,0),(-1,-1),10),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story += [t, Spacer(1,12),
              Paragraph("Cohérence visuelle de l'ensemble", ST["section"]),
              CohBar(coh), Spacer(1,5),
              Paragraph("L'indice mesure l'harmonie entre les décorations sélectionnées (style, couleurs, ambiance).",ST["body"]),
              PageBreak()]

    # Fiches
    for i,d in enumerate(decos):
        hdr = Table([[Paragraph(d["nom"],ST["card_t"]), Paragraph(d["id"],ST["card_id"])]],
                    colWidths=[12*cm,3.1*cm])
        hdr.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),CBL if d["style"] in ["Moderne","Luxe"] else CDG),
            ("TOPPADDING",(0,0),(-1,-1),13),("BOTTOMPADDING",(0,0),(-1,-1),13),
            ("LEFTPADDING",(0,0),(0,-1),14),("RIGHTPADDING",(-1,0),(-1,-1),14),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(1,0),(1,-1),"RIGHT"),
        ]))
        story.append(KeepTogether([hdr]))
        story.append(Spacer(1,10))
        story.append(PhotoBlock(d["photo_url"],d["id"]))
        story.append(Spacer(1,10))
        meta = [[
            Paragraph(f"<b>Style</b><br/>{d['style']}",ST["body"]),
            Paragraph(f"<b>Couleurs</b><br/>{d['couleur1']} / {d['couleur2']}",ST["body"]),
            Paragraph(f"<b>Disponible</b><br/>{d['qty']} ex.",ST["body"]),
            Paragraph(f"<b>Haut. max</b><br/>{d['hauteur_max']} cm",ST["body"]),
            Paragraph(f"<b>Larg. max</b><br/>{d['largeur_max']} cm",ST["body"]),
        ]]
        mt = Table(meta,colWidths=[3.02*cm]*5)
        mt.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),CGR),("GRID",(0,0),(-1,-1),0.5,CBR),
            ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("LEFTPADDING",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        story += [mt, Spacer(1,10),
                  Paragraph("Description", ST["section"]),
                  Paragraph(d["desc"] or "À renseigner dans Google Sheets.", ST["body"]),
                  Spacer(1,8),
                  Paragraph(f"Composition — {len(d['modules'])} module(s)", ST["section"])]
        mod_rows = [[
            Paragraph("<b>Type</b>",ST["label"]),
            Paragraph("<b>Nom du module</b>",ST["label"]),
            Paragraph("<b>Dimensions</b>",ST["label"]),
            Paragraph("<b>Qté</b>",ST["label"]),
        ]]
        for m in d["modules"]:
            mod_rows.append([
                Paragraph(m["type"],ST["small"]),
                Paragraph(m["nom"],ST["body"]),
                Paragraph(m["dim"],ST["small"]),
                Paragraph(str(m["qty"]),ST["body"]),
            ])
        mt2 = Table(mod_rows,colWidths=[2.8*cm,7.3*cm,3.5*cm,1.5*cm])
        mt2.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),CDG),("TEXTCOLOR",(0,0),(-1,0),rl_colors.white),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[CGR,rl_colors.white]),
            ("GRID",(0,0),(-1,-1),0.5,CBR),
            ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("LEFTPADDING",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ALIGN",(3,0),(3,-1),"CENTER"),
        ]))
        story.append(mt2)
        if d["tags"]:
            story += [Spacer(1,6), Paragraph(f"<b>Mots-clés :</b> {d['tags']}",ST["small"])]
        if i < len(decos)-1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=draw_cover, onLaterPages=lambda c,d: None,
              canvasmaker=NumCanvas)
    buf.seek(0)
    return buf

# ════════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ════════════════════════════════════════════════════════════════════════════

# En-tête
st.markdown("""
<div class="main-header">
  <h1>🎄 Agent IA — Décorations Noël</h1>
  <p>Générez automatiquement vos dossiers de présentation à partir de votre stock en temps réel</p>
</div>
""", unsafe_allow_html=True)

# Chargement du stock
with st.spinner("Chargement du stock depuis Google Sheets..."):
    stock = load_stock()

if not stock:
    st.error("Impossible de charger le stock. Vérifiez la connexion Google Sheets.")
    st.stop()

# ── Sidebar : formulaire ──────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Critères du projet")
    st.caption(f"Stock chargé : {len(stock)} décorations actives")

    style_opts   = [""] + sorted(set(d["style"]    for d in stock if d["style"]))
    couleur_opts = [""] + sorted(set(d["couleur1"] for d in stock if d["couleur1"]))
    taille_opts  = ["", "Petit", "Moyen", "Grand"]

    style_sel   = st.selectbox("Style / thème",    style_opts,
                               format_func=lambda x: "— Indifférent —" if x=="" else x)
    couleur_sel = st.selectbox("Couleur principale", couleur_opts,
                               format_func=lambda x: "— Indifférent —" if x=="" else x)
    taille_sel  = st.selectbox("Taille du site",   taille_opts,
                               format_func=lambda x: "— Non précisé —" if x=="" else x)
    hauteur_sel = st.number_input("Hauteur sous plafond (cm)", min_value=0, max_value=2000,
                                   value=0, step=10,
                                   help="0 = non précisé")
    budget_sel  = st.number_input("Budget indicatif (€)", min_value=0, max_value=500000,
                                   value=0, step=500,
                                   help="0 = non précisé")
    nb_sel      = st.slider("Nombre de décorations", min_value=1, max_value=min(8,len(stock)),
                             value=3)
    notes_sel   = st.text_area("Informations complémentaires",
                                placeholder="Ex : galerie couverte 80m, clientèle familiale…",
                                height=80)
    st.divider()
    generate_btn = st.button("🎄 Générer le dossier", use_container_width=True, type="primary")
    st.button("🔄 Rafraîchir le stock", use_container_width=True,
              on_click=lambda: st.cache_data.clear())

# ── Onglets principaux ────────────────────────────────────────────────────────
tab_dossier, tab_stock = st.tabs(["📄 Dossier généré", "📦 Stock disponible"])

# ── Onglet Stock ──────────────────────────────────────────────────────────────
with tab_stock:
    st.markdown("### Stock des décorations actives")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        filtre_style = st.selectbox("Filtrer par style", ["Tous"] +
                                    sorted(set(d["style"] for d in stock if d["style"])),
                                    key="f_style")
    with col_f2:
        filtre_dispo = st.checkbox("Afficher uniquement les disponibles (qty > 0)", value=False)

    stock_filtered = [d for d in stock
                      if (filtre_style == "Tous" or d["style"] == filtre_style)
                      and (not filtre_dispo or d["qty"] > 0)]

    for d in stock_filtered:
        badge_color = ("🟢" if d["qty"] >= 3 else "🟡" if d["qty"] >= 1 else "🔴")
        with st.expander(f"{badge_color}  {d['id']} — {d['nom']}  |  {d['style']}  |  {d['couleur1']} / {d['couleur2']}  |  Qté : {d['qty']}"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Quantité disponible", d["qty"])
            c2.metric("Hauteur max",         f"{d['hauteur_max']} cm")
            c3.metric("Largeur max module",  f"{d['largeur_max']} cm")
            if d["desc"]:
                st.caption(d["desc"])
            if d["modules"]:
                st.markdown("**Modules :**")
                rows = [{"Type": m["type"], "Nom": m["nom"],
                         "Dimensions": m["dim"], "Qté stock": m["qty"]}
                        for m in d["modules"]]
                st.dataframe(rows, use_container_width=True, hide_index=True)

# ── Onglet Dossier ────────────────────────────────────────────────────────────
with tab_dossier:
    if not generate_btn:
        st.info("👈 Renseignez vos critères dans le panneau gauche et cliquez sur **Générer le dossier**.")
    else:
        params = {
            "style":         style_sel,
            "couleur":       couleur_sel,
            "couleur_label": couleur_sel,
            "taille":        taille_sel,
            "hauteur":       hauteur_sel if hauteur_sel > 0 else None,
            "budget":        budget_sel  if budget_sel  > 0 else None,
            "nb":            nb_sel,
            "notes":         notes_sel,
        }

        with st.spinner("L'agent sélectionne les décorations..."):
            selected = select_decos(stock, params)

        if not selected:
            st.warning("Aucune décoration ne correspond aux critères. Essayez d'élargir votre sélection.")
        else:
            coh     = coherence(selected)
            coh_lbl = "Excellente" if coh>=80 else "Bonne" if coh>=55 else "Partielle"
            coh_col = "green"      if coh>=80 else "orange" if coh>=55 else "red"
            nb_mod  = sum(len(d["modules"]) for d in selected)

            # Résumé
            m1,m2,m3,m4 = st.columns(4)
            m1.metric("Décorations sélectionnées", len(selected))
            m2.metric("Modules au total",          nb_mod)
            m3.metric("Cohérence",                 f"{coh}%")
            m4.metric("Stock min disponible",      min(d["qty"] for d in selected))

            st.markdown(f"**Cohérence de l'ensemble :** :{coh_col}[{coh_lbl} ({coh}%)]")
            st.progress(coh / 100)
            st.divider()

            # Cartes décorations
            for d in selected:
                badge = ("🔵 Moderne" if d["style"]=="Moderne" else
                         "🟤 Traditionnel" if d["style"]=="Traditionnel" else
                         "🟢 Nature" if d["style"]=="Nature" else
                         "⭐ Luxe" if d["style"]=="Luxe" else d["style"])
                with st.expander(f"**{d['id']}** — {d['nom']}  |  {badge}  |  Qté : {d['qty']}", expanded=True):
                    col_info, col_photo = st.columns([1.6, 1])
                    with col_info:
                        ci1,ci2,ci3 = st.columns(3)
                        ci1.metric("Couleurs",      f"{d['couleur1']} / {d['couleur2']}")
                        ci2.metric("Haut. max",     f"{d['hauteur_max']} cm")
                        ci3.metric("Larg. max",     f"{d['largeur_max']} cm")
                        if d["desc"]:
                            st.markdown(f"_{d['desc']}_")
                        st.markdown(f"**Modules ({len(d['modules'])}) :**")
                        rows = [{"Type": m["type"], "Nom": m["nom"],
                                 "Dimensions": m["dim"], "Qté": m["qty"]}
                                for m in d["modules"]]
                        st.dataframe(rows, use_container_width=True, hide_index=True)
                    with col_photo:
                        if d["photo_url"] and d["photo_url"].startswith("http"):
                            st.markdown(f"[📷 Voir la photo]({d['photo_url']})")
                        else:
                            st.caption("📷 Photo à renseigner dans Google Sheets")
                        if d["tags"]:
                            st.caption(f"🏷️ {d['tags']}")

            st.divider()

            # Génération PDF
            st.markdown("### Télécharger le dossier PDF")
            with st.spinner("Génération du PDF en cours..."):
                pdf_buf = generate_pdf(selected, params)

            ts       = datetime.now().strftime("%Y%m%d_%H%M")
            style_slug = re.sub(r'[^a-zA-Z0-9]','_', style_sel or "Tous")
            filename = f"Dossier_Noel_{style_slug}_{ts}.pdf"

            st.download_button(
                label="⬇️ Télécharger le dossier PDF",
                data=pdf_buf,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True,
            )
            st.caption(f"Fichier : {filename}")
