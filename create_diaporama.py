from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import copy, math

# ── Palette ──────────────────────────────────────────────────────────────────
INK       = RGBColor(0x0D, 0x1B, 0x3E)   # bleu marine profond
COBALT    = RGBColor(0x1A, 0x4F, 0x8A)   # bleu cobalt
SKY       = RGBColor(0x4A, 0x90, 0xD9)   # bleu ciel
MINT      = RGBColor(0x00, 0xC9, 0xA7)   # vert menthe
CORAL     = RGBColor(0xFF, 0x6B, 0x6B)   # rouge corail
GOLD      = RGBColor(0xFF, 0xC2, 0x00)   # jaune or
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xF5, 0xF7, 0xFF)
LIGHT_BG  = RGBColor(0xEA, 0xF2, 0xFB)
SLATE     = RGBColor(0x4A, 0x5A, 0x7A)
PALE_MINT = RGBColor(0xD4, 0xF5, 0xEF)
PALE_CORAL= RGBColor(0xFF, 0xE5, 0xE5)

def hex2rgb(h): return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ── Primitives ────────────────────────────────────────────────────────────────
def rect(sl, x,y,w,h, fill, radius=0):
    sp = sl.shapes.add_shape(1, Inches(x),Inches(y),Inches(w),Inches(h))
    sp.line.fill.background()
    if fill:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    else:
        sp.fill.background()
    return sp

def tb(sl, text, x,y,w,h, size=16, bold=False, color=INK,
       align=PP_ALIGN.LEFT, italic=False, wrap=True, spacing=1.0):
    box = sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    box.word_wrap = wrap
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size); run.font.bold = bold
    run.font.italic = italic; run.font.color.rgb = color
    return box

def mtb(sl, lines, x,y,w,h, size=15, color=INK, bold_first=False, spacing=Pt(4),
        align=PP_ALIGN.LEFT, line_bold=False):
    """Multi-line textbox. lines = list of str."""
    box = sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    box.word_wrap = True
    tf = box.text_frame; tf.word_wrap = True
    for i,line in enumerate(lines):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment = align; p.space_after = spacing
        r = p.add_run(); r.text = line
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.bold = (bold_first and i==0) or line_bold
    return box

def gradient_bg(sl, c1, c2):
    """Fake gradient: two rects side by side."""
    rect(sl,0,0,6.67,7.5,c1)
    rect(sl,6.67,0,6.66,7.5,c2)

# ── Card helper ───────────────────────────────────────────────────────────────
def card(sl, x,y,w,h, bg, title=None, title_color=WHITE,
         title_bg=None, icon=None, lines=None, line_size=14,
         line_color=None, title_size=16):
    rect(sl,x,y,w,h,bg)
    lc = line_color or title_color
    cy = y+0.12
    if title_bg:
        rect(sl,x,y,w,0.48,title_bg)
    if icon:
        tb(sl,icon,x+0.1,cy,0.55,0.45,size=22,align=PP_ALIGN.CENTER,color=title_color)
        ix=x+0.6
    else:
        ix=x+0.18
    if title:
        tb(sl,title,ix,cy,w-(ix-x)-0.1,0.42,size=title_size,bold=True,
           color=title_color,align=PP_ALIGN.LEFT)
        cy+=0.5
    if lines:
        mtb(sl,lines,x+0.18,cy,w-0.3,h-(cy-y)-0.1,size=line_size,color=lc)

def kpi_card(sl, value, label, x,y,w=3.0,h=1.5, bg=INK, val_color=GOLD, lbl_color=WHITE):
    rect(sl,x,y,w,h,bg)
    tb(sl,value,x,y+0.12,w,0.75,size=36,bold=True,color=val_color,align=PP_ALIGN.CENTER)
    tb(sl,label,x+0.1,y+0.82,w-0.2,0.6,size=13,color=lbl_color,align=PP_ALIGN.CENTER)

def divider(sl, y, color=GOLD, x=0.5, w=12.33, h=0.05):
    rect(sl,x,y,w,h,color)

def section_tag(sl, label, color=COBALT):
    rect(sl,10.5,0.12,2.7,0.42,color)
    tb(sl,label,10.5,0.12,2.7,0.42,size=12,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

def slide_title(sl, title, subtitle=None, y=0.22):
    tb(sl,title,0.5,y,9.8,0.75,size=30,bold=True,color=WHITE,align=PP_ALIGN.LEFT)
    if subtitle:
        tb(sl,subtitle,0.5,y+0.72,9.8,0.45,size=16,color=SKY,italic=True)

def dark_header(sl, title, subtitle=None, tag=None):
    rect(sl,0,0,13.33,1.55,INK)
    rect(sl,0,1.55,13.33,5.95,OFF_WHITE)
    rect(sl,0,1.55,0.08,5.95,COBALT)
    if tag: section_tag(sl,tag)
    slide_title(sl,title,subtitle)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 1 — COUVERTURE
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
rect(sl,0,0,13.33,7.5,INK)
# Bande diagonale décorative
rect(sl,0,4.5,13.33,0.08,GOLD)
rect(sl,0,4.58,13.33,2.92,COBALT)
# Cercles décoratifs (simulés avec des carrés colorés semi-transparents)
rect(sl,9.5,-0.5,4.5,4.5,RGBColor(0x1E,0x3A,0x6E))
rect(sl,10.2,0.2,3.0,3.0,RGBColor(0x26,0x4F,0x90))

tb(sl,"STYLO ORTHO",0.7,0.9,12,1.5,size=72,bold=True,color=WHITE,align=PP_ALIGN.LEFT)
rect(sl,0.7,2.35,4.5,0.08,GOLD)
tb(sl,"Stylo correcteur d'orthographe",0.7,2.5,10,0.65,
   size=26,color=SKY,align=PP_ALIGN.LEFT)
tb(sl,"Assemblé en France  ·  Papeterie intelligente",
   0.7,3.15,10,0.5,size=18,color=RGBColor(0x80,0xA8,0xD8),italic=True)

# 3 KPIs en bas
for val,lbl,x in [("214","répondants",0.6),("82,3%","intéressés",4.8),("99,90 €","prix public",9.0)]:
    tb(sl,val,x,4.85,3.5,0.9,size=38,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    tb(sl,lbl,x,5.65,3.5,0.45,size=16,color=WHITE,align=PP_ALIGN.CENTER)

tb(sl,"Analyse du Marché",0.7,6.3,8,0.5,size=16,color=RGBColor(0x80,0xA8,0xD8),bold=True)
tb(sl,"Mathéo Guzzi  ·  Pierre Lachat  ·  Valentin Imbault-Casset",
   0.7,6.8,11,0.4,size=13,color=RGBColor(0x60,0x88,0xB8))

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 2 — INTRODUCTION
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"1. Introduction",tag="INTRODUCTION")

# Grand texte accroche
rect(sl,0.5,1.7,12.33,1.3,COBALT)
tb(sl,"Le premier stylo qui corrige l'orthographe\nen temps réel — sur papier, sans internet.",
   0.8,1.82,11.5,1.1,size=22,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

# 4 KPI cards
kpi_card(sl,"214","répondants à l'enquête",0.5,3.25,2.9,1.6)
kpi_card(sl,"82,3 %","Oui + Peut-être",3.55,3.25,2.9,1.6)
kpi_card(sl,"86,4 %","jugent le produit utile",6.6,3.25,2.9,1.6)
kpi_card(sl,"99,90 €","prix de vente public",9.65,3.25,2.9,1.6)

# Ligne contexte
rect(sl,0.5,5.1,12.33,1.2,PALE_MINT)
mtb(sl,[
    "📍  Marché papeterie France : 4,6 Mds€ en 2025  ·  107 millions d'articles vendus à la rentrée 2024",
    "🚀  Segment stylos intelligents quasi inexistant en France → opportunité de positionnement pionnier"
],0.7,5.18,12,1.05,size=15,color=INK,bold_first=False)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 3 — DEMANDE QUANTITATIVE
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"2.1 Analyse quantitative de la demande",
            subtitle="Poids en volume et en valeur · Taux de croissance · Tendances",tag="DEMANDE")

# 4 grandes stats visuelles en cartes
stats = [
    ("4,6 Mds €","Marché papeterie\nFrance 2025",COBALT),
    ("147,5 Mds $","Marché mondial 2024\n+3,8 % / an jusqu'en 2034",INK),
    ("107 M","Articles vendus\nrentrée scolaire 2024",RGBColor(0x0E,0x6B,0x5E)),
    ("45 %","des répondants font\ndes fautes souvent",RGBColor(0xC0,0x3A,0x2B)),
]
for i,(val,lbl,bg) in enumerate(stats):
    x = 0.4 + i*3.2
    kpi_card(sl,val,lbl,x,1.75,3.0,1.9,bg=bg)

# Insight encadré
rect(sl,0.4,3.95,12.53,1.25,LIGHT_BG)
rect(sl,0.4,3.95,0.12,1.25,GOLD)
tb(sl,"📊 Tri à plat — Fréquence des fautes (n=214)",
   0.65,4.05,12,0.42,size=16,bold=True,color=INK)
tb(sl,"45 % des répondants font des fautes souvent ou très souvent. "
   "Ce chiffre valide le besoin concret et quotidien du produit sur le marché. "
   "C'est la cible directe et fonctionnelle du Stylo Ortho.",
   0.65,4.5,12,0.62,size=15,color=SLATE)

# Barre visuelle proportion
rect(sl,0.4,5.4,12.53,0.9,WHITE)
tb(sl,"Fréquence des fautes d'orthographe (enquête n=214)",
   0.6,5.44,12,0.38,size=14,bold=True,color=INK)
# Barres proportionnelles
bar_data = [("Très souvent","22%",0.22,CORAL),("Souvent","23%",0.23,RGBColor(0xFF,0x9F,0x6B)),
            ("Parfois","35%",0.35,SKY),("Rarement","20%",0.20,LIGHT_BG)]
bx = 0.6; bw_total = 12.1; by = 5.88
for lbl,pct,ratio,col in bar_data:
    bw = bw_total * ratio
    rect(sl,bx,by,bw,0.34,col)
    if ratio > 0.1:
        tb(sl,f"{lbl}\n{pct}",bx+0.05,by,bw-0.1,0.34,size=11,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
    bx += bw

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 4 — DEMANDE QUALITATIVE
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"2.2 Analyse qualitative — Les acteurs",
            subtitle="Qui achète ? Qui utilise ? Qui recommande ?",tag="DEMANDE")

# 4 colonnes acteurs avec design carte
actors = [
    ("🛒","Acheteur",COBALT,["Parents d'élèves\n(30–50 ans)","Élèves eux-mêmes\n(dès le collège)","Entreprises","Écoles et centres\nde formation"]),
    ("✏️","Utilisateur",RGBColor(0x0E,0x6B,0x5E),["Élèves et étudiants\n(dès le collège)","Professionnels\nécrivant à la main","Personnes avec\ndifficultés d'écriture","Apprenants d'une\nlangue étrangère"]),
    ("🎓","Prescripteur",RGBColor(0x7B,0x3F,0x9E),["Enseignants et\néducateurs spécialisés","Orthophonistes","Conseillers\nd'orientation"]),
    ("📣","Influenceur",RGBColor(0xC0,0x3A,0x2B),["Influenceurs éducatifs\nTikTok / Instagram","Médias parentaux\net scolaires","Bouche-à-oreille\nscolaire"]),
]
for i,(icon,title,bg,items) in enumerate(actors):
    x = 0.35 + i*3.25
    rect(sl,x,1.72,3.05,5.5,bg)
    tb(sl,icon,x,1.8,3.05,0.65,size=28,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,title,x,2.38,3.05,0.52,size=19,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    rect(sl,x+0.2,2.88,2.65,0.05,RGBColor(0xFF,0xFF,0xFF))
    for j,item in enumerate(items):
        tb(sl,f"• {item}",x+0.2,3.0+j*0.62,2.65,0.58,size=13,color=WHITE)

# Stat bas
rect(sl,0.35,7.0,12.53,0.38,GOLD)
tb(sl,"46,7 % des répondants ont déjà ressenti de la gêne à cause de leurs fautes — le Stylo Ortho lève ce frein.",
   0.5,7.03,12.2,0.32,size=13,bold=True,color=INK,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 5 — FACTEURS INDIVIDUELS
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"Facteurs individuels",
            subtitle="Ce qui motive — ou freine — l'achat du Stylo Ortho",tag="DEMANDE")

factors = [
    ("💡","Les besoins",COBALT,
     "Besoin préexistant et latent de corriger ses fautes. Le Stylo Ortho le rend explicite et actionnable."),
    ("🚀","Les motivations",RGBColor(0x0E,0x6B,0x5E),
     "Hédoniste : réussir à l'école, écrire sans honte.\nRationnelle : gagner du temps, éviter les erreurs en manuscrit."),
    ("🚧","Les freins",RGBColor(0xC0,0x3A,0x2B),
     "Prix de 99,90 € élevé vs stylo classique (1–5 €).\nCrainte de paraître assisté ou dépendant d'un outil."),
    ("👁","La perception",RGBColor(0x7B,0x3F,0x9E),
     "Design classique avec écran discret → image de produit simple et non intrusif."),
    ("👤","Profil cible",RGBColor(0x1A,0x4F,0x8A),
     "30–50 ans (parents acheteurs) · 11–25 ans (élèves et étudiants utilisateurs)"),
]
y = 1.75
for icon,title,bg,text in factors:
    h = 0.94
    rect(sl,0.4,y,12.53,h,bg)
    tb(sl,icon,0.55,y+(h-0.5)/2,0.6,0.5,size=22,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,title,1.25,y+0.08,2.8,0.42,size=16,bold=True,color=GOLD)
    tb(sl,text,4.2,y+0.08,8.6,h-0.18,size=14,color=WHITE)
    y += h+0.06

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 6 — FACTEURS COLLECTIFS
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"Facteurs collectifs",
            subtitle="Culture, groupes de référence, cycle de vie familiale",tag="DEMANDE")

# 3 grandes cartes
collective = [
    ("🇫🇷","La culture","En France, l'orthographe est un marqueur social fort. Faire des fautes est perçu comme une lacune. "
     "Le Stylo Ortho répond directement à cette pression culturelle et scolaire.",COBALT),
    ("👥","Groupes de référence","Appartenance : camarades de classe, collègues.\n"
     "Aspiration : bons élèves, professionnels valorisant l'écrit soigné.",RGBColor(0x0E,0x6B,0x5E)),
    ("👨‍👩‍👧","Cycle de vie familiale","Famille avec enfants scolarisés : cible prioritaire — la rentrée déclenche l'achat.\n"
     "Jeune actif entrant dans la vie professionnelle : cible secondaire.",RGBColor(0x7B,0x3F,0x9E)),
]
for i,(icon,title,text,bg) in enumerate(collective):
    x = 0.4 + i*4.3
    rect(sl,x,1.72,4.1,3.8,bg)
    tb(sl,icon,x,1.88,4.1,0.7,size=34,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,title,x+0.15,2.6,3.8,0.52,size=17,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    rect(sl,x+0.4,3.1,3.3,0.05,WHITE)
    tb(sl,text,x+0.2,3.22,3.7,2.2,size=14,color=WHITE,wrap=True)

# Tri croisé encadré
rect(sl,0.4,5.72,12.53,1.55,LIGHT_BG)
rect(sl,0.4,5.72,0.12,1.55,MINT)
tb(sl,"📊 Tri croisé — Intérêt par profil (enquête n=214)",
   0.65,5.82,12,0.42,size=16,bold=True,color=INK)
tb(sl,"Les personnes faisant le plus souvent des fautes sont aussi les plus intéressées par le produit. "
   "46,9 % des parents sont prêts à acheter le Stylo Ortho pour leur enfant. "
   "75,2 % des professionnels voient son utilité pour leurs équipes ou élèves.",
   0.65,6.28,12,0.85,size=14,color=SLATE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 7 — SITUATIONS D'ACHAT
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"Situations d'achat",
            subtitle="Quand, comment et pourquoi le consommateur achète",tag="DEMANDE")

# Timeline saisonnière
rect(sl,0.4,1.75,12.53,1.6,WHITE)
tb(sl,"📅  Saisonnalité de l'achat",0.6,1.82,12,0.42,size=16,bold=True,color=INK)
rect(sl,0.6,2.3,12.1,0.5,LIGHT_BG)
months = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
heights = [0.15,0.12,0.12,0.12,0.15,0.12,0.2,0.5,0.5,0.2,0.18,0.35]  # proportionnel
for i,(m,h) in enumerate(zip(months,heights)):
    bx = 0.6 + i*(12.1/12)
    bw = 12.1/12 - 0.04
    col = CORAL if h>=0.4 else (GOLD if h>=0.2 else SKY)
    rect(sl,bx,2.3+(0.5-h),bw,h,col)
    tb(sl,m,bx,2.82,bw,0.25,size=9,color=SLATE,align=PP_ALIGN.CENTER)
tb(sl,"🔴 Pic rentrée : +52,1 % CA en août vs décembre",0.6,3.28,12,0.35,size=13,color=CORAL,bold=True)

# 3 cartes types d'achat
buy_types = [
    ("🆕","Achat de nouveauté","Produit inconnu — forte implication.\nDémonstration en magasin indispensable.",COBALT),
    ("🤔","Achat réfléchi","Prix 99,90 € — achat planifié.\nJustifié par l'USP unique sur le marché.",RGBColor(0x0E,0x6B,0x5E)),
    ("📋","Achat prescrit","Recommandation d'un enseignant\nou d'un parent — effet bouche-à-oreille.",RGBColor(0x7B,0x3F,0x9E)),
]
for i,(icon,title,text,bg) in enumerate(buy_types):
    x = 0.4+i*4.32
    rect(sl,x,3.72,4.1,2.5,bg)
    tb(sl,icon,x,3.82,4.1,0.62,size=28,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,title,x+0.15,4.46,3.8,0.46,size=16,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    tb(sl,text,x+0.2,4.98,3.7,1.1,size=14,color=WHITE)

# Insight bas
rect(sl,0.4,6.35,12.53,0.9,INK)
tb(sl,"📊 56,6 % utilisent le correcteur téléphone · 43,4 % le correcteur Word — "
   "mais aucun ne fonctionne sur papier. Le Stylo Ortho est le seul outil qui fait ça.",
   0.6,6.5,12.1,0.65,size=15,color=WHITE,bold=True,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 8 — OFFRE : PRODUITS ET SEGMENTS
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"3.1 Produits et segments du marché",
            subtitle="Où se positionne le Stylo Ortho dans le marché papeterie ?",tag="OFFRE")

segments = [
    ("📏","Stylos classiques\n(BIC, Pilot…)","≈ 60–65 %\ndu marché","Stable / léger recul","Concurrent\nindirect",SLATE),
    ("✏️","Stylos effaçables\n(FriXion…)","≈ 10–15 %","En croissance","Complémentaire\n(encre effaçable)",RGBColor(0x2E,0x7D,0x52)),
    ("💎","Stylos premium\n/ luxe","≈ 5–8 %","Stable — niche","Inspiration\npositionnement",RGBColor(0x7B,0x3F,0x9E)),
    ("🧠","Instruments\nintelligents","Quasi\ninexistant","Fort\npotentiel","🎯 Terrain\nvierge",CORAL),
]
for i,(icon,name,pdm,evol,pos,bg) in enumerate(segments):
    x = 0.4+i*3.22
    rect(sl,x,1.75,3.05,4.4,bg)
    tb(sl,icon,x,1.85,3.05,0.7,size=30,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,name,x+0.1,2.58,2.85,0.62,size=14,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    rect(sl,x+0.3,3.2,2.45,0.05,WHITE)
    tb(sl,"Part de marché",x+0.15,3.32,2.75,0.35,size=11,color=RGBColor(0xC8,0xD8,0xE8),italic=True)
    tb(sl,pdm,x+0.15,3.62,2.75,0.55,size=15,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
    tb(sl,"Évolution",x+0.15,4.2,2.75,0.32,size=11,color=RGBColor(0xC8,0xD8,0xE8),italic=True)
    tb(sl,evol,x+0.15,4.5,2.75,0.45,size=13,color=WHITE,align=PP_ALIGN.CENTER)
    tb(sl,pos,x+0.15,4.98,2.75,0.55,size=13,bold=True,color=GOLD,align=PP_ALIGN.CENTER)

rect(sl,0.4,6.28,12.53,0.98,PALE_MINT)
rect(sl,0.4,6.28,0.12,0.98,MINT)
tb(sl,"📊 Tri à plat — 56,6 % utilisent le correcteur téléphone et 43,4 % le correcteur Word, "
   "mais aucun ne fonctionne sur papier. "
   "Le Stylo Ortho occupe un territoire inexploité.",
   0.65,6.38,12,0.8,size=14,color=INK)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 9 — CONCURRENTS
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"3.2 Analyse concurrentielle",
            subtitle="Concurrents directs (numérique) et indirects (papeterie)",tag="OFFRE")

# Split gauche/droite
rect(sl,0.4,1.72,5.9,5.55,RGBColor(0xE8,0xF0,0xFB))
rect(sl,6.45,1.72,6.48,5.55,RGBColor(0xFF,0xED,0xED))
rect(sl,0.4,1.72,5.9,0.52,COBALT)
rect(sl,6.45,1.72,6.48,0.52,RGBColor(0xC0,0x3A,0x2B))
tb(sl,"INDIRECTS — Papeterie",0.4,1.77,5.9,0.42,size=15,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"DIRECTS — Numérique",6.45,1.77,6.48,0.42,size=15,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

indirect = [
    ("BIC","≈ 20–25 % marché stylos","Notoriété mondiale, prix très bas","❌ Aucune correction intégrée"),
    ("Pilot / Stabilo / Pentel","≈ 15–20 % cumulés","Innovation stylos effaçables","❌ Pas de correction ortho"),
]
for i,(name,pdm,force,limite) in enumerate(indirect):
    y = 2.42+i*1.35
    rect(sl,0.55,y,5.6,1.22,WHITE)
    tb(sl,name,0.72,y+0.1,3.0,0.42,size=15,bold=True,color=COBALT)
    tb(sl,pdm,0.72,y+0.52,3.0,0.35,size=12,color=SLATE,italic=True)
    tb(sl,f"✅ {force}",0.72,y+0.78,2.6,0.35,size=12,color=RGBColor(0x0E,0x6B,0x5E))
    tb(sl,limite,3.4,y+0.1,2.6,0.75,size=12,color=CORAL,bold=True)

direct = [
    ("Correcteur téléphone","≈ 56,6 % utilisateurs","Gratuit, partout, instantané","❌ Hors papier"),
    ("Microsoft Word /\nGoogle Docs","≈ 43,4 % utilisateurs","Intégré, gratuit","❌ Inutilisable manuscrit"),
    ("Grammarly / Antidote","Niche","Très performant à l'écrit","❌ Réservé numérique"),
]
for i,(name,pdm,force,limite) in enumerate(direct):
    y = 2.42+i*1.0
    rect(sl,6.6,y,6.15,0.88,WHITE)
    tb(sl,name,6.75,y+0.07,3.2,0.42,size=14,bold=True,color=RGBColor(0xC0,0x3A,0x2B))
    tb(sl,pdm,6.75,y+0.5,2.0,0.3,size=11,color=SLATE,italic=True)
    tb(sl,limite,9.1,y+0.07,3.5,0.75,size=12,color=CORAL,bold=True)
    tb(sl,f"✅ {force}",6.75,y+0.57,3.0,0.28,size=11,color=RGBColor(0x0E,0x6B,0x5E))

rect(sl,0.4,6.42,12.53,0.85,INK)
tb(sl,"🎯 Le Stylo Ortho occupe un territoire inexploité : "
   "correction en temps réel sur papier, sans internet — aucun concurrent direct.",
   0.6,6.57,12.1,0.6,size=15,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 10 — DISTRIBUTEURS
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"3.3 Stratégie de distribution",
            subtitle="Priorités et circuits de mise sur le marché",tag="OFFRE")

distrib = [
    ("1","Bureau Vallée · Fnac · Cultura","Circuit spécialisé (court)","≈ 15–20 %",
     "Démonstration produit possible en rayon.\nEssentiel pour un produit innovant → Priorité 1.",COBALT),
    ("2","Amazon · Fnac.com · stylo-ortho.fr","E-commerce (ultra-court)","≈ 49 %",
     "Parents qui achètent en ligne.\nMarge maîtrisée — forte progression +15 %/an → Priorité 2.",RGBColor(0x0E,0x6B,0x5E)),
    ("⏳","Carrefour · E.Leclerc","Grandes surfaces (circuit long)","332 M€ rentrée",
     "Phase 2 uniquement — après notoriété établie.\nRecul –7 % en 2024 → approche prudente.",SLATE),
    ("🤝","Lycées · Collèges · Entreprises","Circuit direct B2B","Fort potentiel",
     "Partenariat Ministère Éducation nationale.\nObjectif : 500 stylos An 1, 2 000+ An 2.",RGBColor(0x7B,0x3F,0x9E)),
]
y = 1.75
for badge,name,circuit,pdm,text,bg in distrib:
    h = 1.28
    rect(sl,0.4,y,12.53,h,LIGHT_BG)
    rect(sl,0.4,y,1.5,h,bg)
    tb(sl,badge,0.4,y+(h-0.55)/2,1.5,0.55,size=28,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
    tb(sl,name,2.0,y+0.1,4.5,0.48,size=16,bold=True,color=INK)
    tb(sl,circuit,2.0,y+0.58,3.5,0.35,size=13,color=SLATE,italic=True)
    rect(sl,6.6,y+0.2,0.05,h-0.4,COBALT)
    tb(sl,f"PDM : {pdm}",6.75,y+0.1,2.0,0.42,size=13,bold=True,color=bg)
    tb(sl,text,6.75,y+0.52,5.8,0.72,size=13,color=SLATE)
    y += h+0.07

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 11 — PESTEL
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"3.4 Analyse PESTEL",
            subtitle="Environnement macro-économique du Stylo Ortho",tag="OFFRE")

pestel = [
    ("🏛","Politique",COBALT,
     "Politiques éducation et lutte contre\nl'illettrisme favorables",
     "Réglementations électronique scolaire\nCertifications obligatoires"),
    ("💶","Économique",RGBColor(0x0E,0x6B,0x5E),
     "Marché innovation en croissance\nBudget scolaire familles stable",
     "Pouvoir d'achat sous pression\n(99,90 € = achat réfléchi)"),
    ("👥","Socioculturel",RGBColor(0x7B,0x3F,0x9E),
     "Orthographe = marqueur social en France\n300 M francophones · Dyslexie 5–10 %",
     "Résistance enseignants\nHabitude stylo classique"),
    ("🔬","Technologique",RGBColor(0xC0,0x3A,0x2B),
     "Progrès composants électroniques\nStylos connectés en développement",
     "Évolution rapide IA\nRisque substitution à terme"),
    ("🌱","Écologique",RGBColor(0x1A,0x7A,0x5E),
     "Produit rechargeable USB-C\nFaible empreinte vs stylos jetables",
     "Normes recyclabilité Europe\nDemande réduction plastique"),
    ("⚖","Légal",SLATE,
     "Normes EU maîtrisables\nAssemblage France → conformité CE",
     "Conformité produits enfants\nDépôt de brevet nécessaire"),
]
cols = 3
for i,(icon,title,bg,opp,men) in enumerate(pestel):
    row = i//cols; col = i%cols
    x = 0.4 + col*4.3
    y = 1.72 + row*2.72
    rect(sl,x,y,4.1,2.55,bg)
    tb(sl,icon+" "+title,x+0.15,y+0.1,3.8,0.52,size=16,bold=True,color=GOLD)
    rect(sl,x+0.15,y+0.65,3.8,0.05,WHITE)
    tb(sl,"✅ "+opp,x+0.15,y+0.78,3.75,0.9,size=12,color=WHITE)
    tb(sl,"⚠️ "+men,x+0.15,y+1.62,3.75,0.82,size=12,color=RGBColor(0xFF,0xD0,0xB0))

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 12 — OPPORTUNITÉS & MENACES
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"Opportunités & Menaces",
            subtitle="Synthèse stratégique du marché",tag="SYNTHÈSE")

rect(sl,0.4,1.72,6.1,5.55,RGBColor(0xE8,0xF9,0xF2))
rect(sl,6.65,1.72,6.28,5.55,RGBColor(0xFF,0xED,0xED))
rect(sl,0.4,1.72,6.1,0.55,RGBColor(0x0E,0x6B,0x5E))
rect(sl,6.65,1.72,6.28,0.55,RGBColor(0xC0,0x3A,0x2B))
tb(sl,"✅  OPPORTUNITÉS",0.4,1.78,6.1,0.44,size=17,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"⚠️  MENACES",6.65,1.78,6.28,0.44,size=17,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

opps = [
    "Besoin croissant d'outils pédagogiques innovants",
    "Segment stylos intelligents quasi inexistant en France\n→ positionnement pionnier",
    "82,3 % des répondants intéressés (n=214)",
    "Assemblage France (Nicomatic) → argument marketing et conformité CE",
    "49 % des transactions papeterie en e-commerce\n→ canal direct fort",
    "Rentrée scolaire : pic achat prévisible et massif",
]
threats = [
    "Concurrence outils numériques (correcteurs, IA, GPT)",
    "Prix 99,90 € perçu comme élevé vs stylo classique (1–5 €)",
    "Recul marché papeterie : –7 % en 2024 vs 2023",
    "Résistance culturelle de certains enseignants",
    "Évolution rapide de l'IA — risque substitution à moyen terme",
    "Certifications et réglementations produits électroniques enfants",
]
y = 2.45
for opp,thr in zip(opps,threats):
    tb(sl,f"  • {opp}",0.55,y,5.8,0.72,size=13,color=RGBColor(0x0E,0x4A,0x2E))
    tb(sl,f"  • {thr}",6.8,y,6.0,0.72,size=13,color=RGBColor(0x7A,0x10,0x10))
    y += 0.74

rect(sl,0.4,6.88,12.53,0.5,INK)
tb(sl,"Le Stylo Ortho est positionné pour être le leader d'un marché pionnier grâce à son USP unique : "
   "correction en temps réel sur papier, sans internet.",
   0.6,6.9,12.1,0.45,size=14,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 13 — CONCEPT PRODUIT
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
rect(sl,0,0,13.33,7.5,INK)
rect(sl,0,0,5.5,7.5,COBALT)
section_tag(sl,"CONCEPT PRODUIT")

tb(sl,"Concept\nProduit",0.4,1.0,4.8,1.8,size=38,bold=True,color=WHITE)
rect(sl,0.4,2.78,4.5,0.08,GOLD)
tb(sl,"Un stylo classique enrichi d'une\ntechnologie discrète de correction\northographique en temps réel.",
   0.4,2.95,4.7,1.5,size=17,color=SKY)

# Partenaires bas gauche
tb(sl,"Partenaires :",0.4,5.2,4.7,0.38,size=13,bold=True,color=GOLD)
mtb(sl,["🇨🇳 Shenzhen Apec — électronique",
        "🇯🇵 Pentel / Aihao — encre effaçable",
        "🇫🇷 Nicomatic — assemblage final"],
    0.4,5.58,4.8,1.4,size=13,color=WHITE)

# Fonctionnalités droite
features = [
    ("⚡","Recharge USB-C","15 min de charge — pratique au quotidien",COBALT),
    ("🖊","Écran intégré","Correction en rouge, orthographe proposée",RGBColor(0xC0,0x3A,0x2B)),
    ("🔄","Encre rechargeable","Économique vs stylos jetables",RGBColor(0x0E,0x6B,0x5E)),
    ("✏️","Gomme effaçable","Usage fluide, pas de correcteur blanc",RGBColor(0x7B,0x3F,0x9E)),
    ("🎨","4 couleurs","Noir · Bleu · Rouge · Vert",SLATE),
    ("🏃","Design ergonomique","Léger, toutes durées d'utilisation",RGBColor(0x2E,0x5F,0x9E)),
]
fy=1.5
for i,(icon,title,desc,bg) in enumerate(features):
    row=i//2; col=i%2
    x=5.8+col*3.7; y=fy+row*1.82
    rect(sl,x,y,3.45,1.65,bg)
    tb(sl,icon,x+0.12,y+0.15,0.55,0.5,size=24,color=WHITE)
    tb(sl,title,x+0.75,y+0.15,2.55,0.45,size=14,bold=True,color=GOLD)
    tb(sl,desc,x+0.75,y+0.6,2.55,0.8,size=13,color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 14 — SEGMENTS
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"Segmentation du marché",
            subtitle="Qui sont nos clients ? Qui cibler en priorité ?",tag="MARKETING")

segs = [
    ("🎯","Parents d'élèves","30–50 ans — acheteurs principaux","46,9 % prêts à acheter\npour leur enfant (enquête)",COBALT,"CIBLE PRINCIPALE"),
    ("✏️","Élèves & Étudiants","11–25 ans — utilisateurs directs","45 % font des fautes\nsouvent ou très souvent",RGBColor(0x0E,0x6B,0x5E),"CIBLE SECONDAIRE"),
    ("🎓","Professionnels & Enseignants","25–50 ans — prescripteurs","75,2 % voient l'utilité\npour élèves / collègues",RGBColor(0x7B,0x3F,0x9E),"PRESCRIPTEURS"),
]
for i,(icon,name,profil,chiffre,bg,badge) in enumerate(segs):
    y = 1.72 + i*1.62
    rect(sl,0.4,y,12.53,1.52,bg)
    tb(sl,icon,0.55,y+0.45,0.8,0.7,size=30,color=WHITE,align=PP_ALIGN.CENTER)
    tb(sl,name,1.5,y+0.12,4.5,0.55,size=20,bold=True,color=GOLD)
    tb(sl,profil,1.5,y+0.68,4.5,0.45,size=14,color=WHITE,italic=True)
    rect(sl,6.2,y+0.2,0.05,1.1,RGBColor(0xFF,0xFF,0xFF))
    tb(sl,"📊 "+chiffre,6.4,y+0.18,4.2,0.8,size=16,bold=True,color=WHITE)
    rect(sl,10.8,y+0.08,1.95,0.45,GOLD)
    tb(sl,badge,10.8,y+0.1,1.95,0.42,size=11,bold=True,color=INK,align=PP_ALIGN.CENTER)

rect(sl,0.4,6.72,12.53,0.65,LIGHT_BG)
rect(sl,0.4,6.72,0.12,0.65,MINT)
tb(sl,"Stratégie indifférenciée : un seul produit pour tous les profils. "
   "Communication en priorité sur les parents et élèves lors de la rentrée scolaire.",
   0.65,6.8,12,0.52,size=14,bold=True,color=INK)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 15 — CIBLAGE & POSITIONNEMENT
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
rect(sl,0,0,13.33,7.5,OFF_WHITE)
rect(sl,0,0,13.33,1.55,INK)
section_tag(sl,"MARKETING")
slide_title(sl,"Ciblage & Positionnement")

# Ciblage — 3 bulles
rect(sl,0.4,1.68,5.9,3.6,LIGHT_BG)
rect(sl,0.4,1.68,5.9,0.5,COBALT)
tb(sl,"CIBLAGE",0.4,1.72,5.9,0.42,size=16,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
targets = [
    ("🥇 Parents d'élèves","46,9 % prêts à acheter",COBALT),
    ("🥈 Élèves & étudiants","45 % font souvent des fautes",RGBColor(0x0E,0x6B,0x5E)),
    ("🥉 Professionnels","75,2 % voient l'utilité",RGBColor(0x7B,0x3F,0x9E)),
]
for i,(name,stat,col) in enumerate(targets):
    y2 = 2.28+i*0.97
    rect(sl,0.55,y2,5.6,0.84,col)
    tb(sl,name,0.72,y2+0.1,3.5,0.42,size=15,bold=True,color=WHITE)
    tb(sl,stat,4.3,y2+0.1,1.7,0.42,size=13,color=GOLD,bold=True,align=PP_ALIGN.RIGHT)

# Positionnement — grande carte
rect(sl,6.5,1.68,6.48,3.6,INK)
rect(sl,6.5,1.68,6.48,0.5,COBALT)
tb(sl,"POSITIONNEMENT",6.5,1.72,6.48,0.42,size=16,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"Dans la tête du consommateur :",6.7,2.28,6.1,0.4,size=14,color=SKY,italic=True)
rect(sl,6.7,2.75,6.0,0.06,GOLD)
tb(sl,"« Le stylo intelligent\nqui corrige mes fautes\nquand j'écris à la main »",
   6.7,2.88,6.0,1.55,size=22,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"Axe fonctionnel · Émotionnel · Made in France",6.7,4.45,6.0,0.42,
   size=13,color=SKY,italic=True,align=PP_ALIGN.CENTER)

# Axes de communication
rect(sl,0.4,5.45,12.53,1.85,WHITE)
rect(sl,0.4,5.45,0.12,1.85,COBALT)
tb(sl,"Axes de communication",0.65,5.52,12,0.42,size=16,bold=True,color=INK)
axes = [
    ("🎯 Fonctionnel","Correction orthographique\nen temps réel sur papier",COBALT),
    ("❤️ Émotionnel","Confiance en soi\nRéussite scolaire et pro",RGBColor(0xC0,0x3A,0x2B)),
    ("🏅 Symbolique","Produit made in France\nDiscrèt et innovant",RGBColor(0x7B,0x3F,0x9E)),
    ("📱 Digital","Aucun concurrent\nsur papier — USP unique",RGBColor(0x0E,0x6B,0x5E)),
]
for i,(title,text,col) in enumerate(axes):
    x=0.55+i*3.1
    rect(sl,x,5.98,2.95,1.22,col)
    tb(sl,title,x+0.1,6.05,2.75,0.42,size=13,bold=True,color=GOLD)
    tb(sl,text,x+0.1,6.48,2.75,0.65,size=13,color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 16 — BÉNÉFICE CONSOMMATEUR
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
rect(sl,0,0,13.33,7.5,INK)
section_tag(sl,"MARKETING")

tb(sl,"Bénéfice\nConsommateur",0.5,0.35,10,1.5,size=42,bold=True,color=WHITE)
rect(sl,0.5,1.85,12.33,0.08,GOLD)

benefits = [
    ("🔧","Bénéfice\nFonctionnel","Ce que le produit fait",
     "Corriger les fautes d'orthographe en temps réel sur papier, "
     "sans connexion, en situation de manuscrit (école, bureau, examen).",
     COBALT),
    ("❤️","Bénéfice\nÉmotionnel","Ce que l'utilisateur ressent",
     "Confiance en soi, réduction de la honte liée aux fautes, "
     "sérénité lors des examens et prises de notes importantes.",
     RGBColor(0xC0,0x3A,0x2B)),
    ("🏆","Bénéfice\nSymbolique","Ce que le produit dit de soi",
     "Utiliser un produit made in France innovant. "
     "Montrer qu'on valorise l'écrit soigné tout en adoptant les nouvelles technologies.",
     RGBColor(0x7B,0x3F,0x9E)),
]
for i,(icon,title,subtitle,text,bg) in enumerate(benefits):
    x=0.4+i*4.3
    rect(sl,x,2.1,4.1,3.55,bg)
    tb(sl,icon,x,2.2,4.1,0.75,size=32,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,title,x+0.15,2.97,3.8,0.7,size=18,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    tb(sl,subtitle,x+0.15,3.65,3.8,0.38,size=12,color=RGBColor(0xC8,0xD8,0xE8),
       italic=True,align=PP_ALIGN.CENTER)
    rect(sl,x+0.35,4.02,3.4,0.05,WHITE)
    tb(sl,text,x+0.2,4.15,3.7,1.38,size=13,color=WHITE,wrap=True)

# Synthèse dorée
rect(sl,0.4,5.85,12.53,1.45,GOLD)
tb(sl,"Bénéfice synthétique :",0.7,5.95,12,0.42,size=15,bold=True,color=INK)
tb(sl,"« Avec le Stylo Ortho, j'écris à la main sans stresser pour mes fautes — "
   "en cours, au bureau ou en examen, c'est le seul outil qui fait ça sur papier. »",
   0.7,6.38,11.9,0.82,size=18,bold=True,color=INK,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 17 — MARKETING MIX : PRODUIT
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"P1 — Produit",
            subtitle="Le Stylo Ortho : un stylo intelligent qui aide à écrire sans fautes",
            tag="MARKETING MIX")

# Visuel central + caractéristiques en orbite
rect(sl,4.8,1.72,3.75,3.1,COBALT)
tb(sl,"✏️",4.8,1.82,3.75,1.5,size=62,align=PP_ALIGN.CENTER,color=WHITE)
tb(sl,"Stylo Ortho\nStandard",4.8,3.3,3.75,0.85,size=18,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
tb(sl,"Assemblé en France 🇫🇷",4.8,4.1,3.75,0.5,size=13,color=WHITE,italic=True,align=PP_ALIGN.CENTER)

# Gauche
for j,(feat,desc) in enumerate([
    ("⚡ Recharge USB-C","15 minutes de charge"),
    ("🔄 Encre rechargeable","Économique et moderne"),
    ("✏️ Gomme effaçable","Usage fluide, sans correcteur"),
]):
    y2=1.82+j*1.02
    rect(sl,0.4,y2,4.1,0.88,RGBColor(0x1E,0x3A,0x6E))
    tb(sl,feat,0.55,y2+0.08,3.8,0.38,size=14,bold=True,color=GOLD)
    tb(sl,desc,0.55,y2+0.48,3.8,0.32,size=13,color=WHITE)

# Droite
for j,(feat,desc) in enumerate([
    ("📺 Écran intégré","Correction en rouge en temps réel"),
    ("🎨 4 couleurs","Noir · Bleu · Rouge · Vert"),
    ("🏃 Design ergonomique","Léger, toutes durées d'utilisation"),
]):
    y2=1.82+j*1.02
    rect(sl,8.85,y2,4.1,0.88,RGBColor(0x1E,0x3A,0x6E))
    tb(sl,feat,9.0,y2+0.08,3.8,0.38,size=14,bold=True,color=GOLD)
    tb(sl,desc,9.0,y2+0.48,3.8,0.32,size=13,color=WHITE)

rect(sl,0.4,5.05,12.53,0.65,PALE_MINT)
rect(sl,0.4,5.05,0.12,0.65,MINT)
tb(sl,"86,4 % des répondants jugent le Stylo Ortho utile ou très utile (enquête n=214) — "
   "le produit répond à un besoin réel et quotidien.",
   0.65,5.12,12,0.52,size=14,bold=True,color=INK)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 18 — MARKETING MIX : PRIX
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"P2 — Prix",subtitle="Stratégie de prix et justification de la valeur",tag="MARKETING MIX")

# 2 grands prix visuels
rect(sl,0.5,1.72,5.8,3.3,COBALT)
rect(sl,6.6,1.72,5.83,3.3,RGBColor(0x0E,0x6B,0x5E))
tb(sl,"99,90 €",0.5,1.88,5.8,1.4,size=56,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
tb(sl,"Prix de vente public",0.5,3.2,5.8,0.48,size=17,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"Standard · tous publics",0.5,3.65,5.8,0.38,size=14,color=SKY,italic=True,align=PP_ALIGN.CENTER)
tb(sl,"44,90 €",6.6,1.88,5.83,1.4,size=56,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
tb(sl,"Prix promotionnel rentrée",6.6,3.2,5.83,0.48,size=17,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"Déclencheur des premiers achats",6.6,3.65,5.83,0.38,size=14,color=RGBColor(0x90,0xE8,0xD0),
   italic=True,align=PP_ALIGN.CENTER)

# Justifications
rect(sl,0.5,5.15,12.43,2.12,LIGHT_BG)
rect(sl,0.5,5.15,0.12,2.12,COBALT)
tb(sl,"Pourquoi 99,90 € ?",0.75,5.22,12,0.42,size=16,bold=True,color=INK)
points = [
    "✅ 82,3 % des répondants sont intéressés ou potentiellement intéressés — le prix est accepté",
    "✅ USP unique sur le marché papier → valeur fonctionnelle qui justifie le prix premium",
    "✅ Comparable à une calculatrice scolaire (30–80 €) ou une montre connectée d'entrée de gamme",
    "✅ Prix psychologique 44,90 € utilisé à la rentrée pour déclencher les premiers achats",
]
for j,pt in enumerate(points):
    tb(sl,pt,0.75,5.7+j*0.36,11.8,0.34,size=13,color=SLATE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 19 — MARKETING MIX : DISTRIBUTION
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
dark_header(sl,"P3 — Distribution",
            subtitle="Une stratégie en 3 phases : sélective → directe → intensive",
            tag="MARKETING MIX")

# Timeline horizontale
rect(sl,0.4,1.72,12.53,0.72,COBALT)
phases = [("Phase 1 — Lancement",0.5,4.3),("Phase 2 — Croissance",4.85,4.3),("Phase 3 — Maturité",9.2,3.7)]
for label,x,w in phases:
    tb(sl,label,x,1.8,w,0.55,size=15,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

distrib2 = [
    ("🏪","Magasins spécialisés","Bureau Vallée · Fnac\nCultura · Librairies",
     "Démonstration produit\nen rayon — essentiel pour\nun produit innovant",
     "PDM ≈ 15–20 %\n+3 % / an",COBALT,0.5,4.05),
    ("🛒","E-commerce","stylo-ortho.fr\nAmazon · Fnac.com",
     "49 % des transactions\npapeterie en ligne\nMarge maîtrisée",
     "PDM ≈ 49 %\n+15 % / an",RGBColor(0x0E,0x6B,0x5E),4.85,4.05),
    ("🏬","Grandes surfaces\nB2B",
     "Carrefour · E.Leclerc\nLycées · Collèges",
     "Phase 2 après notoriété\nPartenariat Ministère\nÉducation Nationale",
     "332 M€ rentrée\nFort potentiel B2B",SLATE,9.2,3.55),
]
for icon,title,pts,why,stats,bg,x,w in distrib2:
    rect(sl,x,2.58,w,4.62,bg)
    tb(sl,icon,x,2.65,w,0.82,size=34,align=PP_ALIGN.CENTER,color=WHITE)
    tb(sl,title,x+0.15,3.5,w-0.3,0.7,size=16,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    rect(sl,x+0.3,4.22,w-0.6,0.05,WHITE)
    tb(sl,pts,x+0.2,4.35,w-0.4,0.85,size=13,color=WHITE,align=PP_ALIGN.CENTER)
    rect(sl,x+0.3,5.22,w-0.6,0.05,RGBColor(0xFF,0xFF,0xFF))
    tb(sl,why,x+0.2,5.35,w-0.4,0.95,size=12,color=RGBColor(0xC8,0xD8,0xE8),italic=True)
    rect(sl,x+0.15,6.38,w-0.3,0.65,RGBColor(0,0,0))
    tb(sl,stats,x+0.15,6.44,w-0.3,0.55,size=12,bold=True,color=GOLD,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 20 — CONCLUSION & SLOGAN
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
rect(sl,0,0,13.33,7.5,INK)
rect(sl,0,5.6,13.33,1.9,COBALT)
# Décoration
rect(sl,0,3.45,13.33,0.08,GOLD)

tb(sl,"CONCLUSION",0.5,0.35,12.33,0.65,size=22,bold=True,color=GOLD,align=PP_ALIGN.CENTER)

tb(sl,"Le Stylo Ortho est le premier stylo correcteur d'orthographe "
   "en temps réel sur papier.\nIl répond à un besoin réel, validé par 214 répondants, "
   "sur un marché encore vierge.",
   0.5,1.05,12.33,1.2,size=19,color=WHITE,align=PP_ALIGN.CENTER)

# 3 chiffres clés
for val,lbl,x in [("82,3 %","intéressés ou\npotentiellement",0.5),
                   ("86,4 %","jugent le produit\nutile",4.8),
                   ("99,90 €","prix de lancement\npublic",9.1)]:
    tb(sl,val,x,2.38,3.6,0.88,size=38,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
    tb(sl,lbl,x,3.18,3.6,0.55,size=15,color=SKY,align=PP_ALIGN.CENTER)

# Slogan central
rect(sl,0.5,3.62,12.33,1.75,GOLD)
tb(sl,"Notre slogan :",0.7,3.72,12,0.42,size=16,color=INK)
tb(sl,"« Écrire juste, à la main. »",0.5,4.05,12.33,1.15,
   size=44,bold=True,color=INK,align=PP_ALIGN.CENTER)

# Pied de page
tb(sl,"Stylo Ortho  ·  Assemblé en France  ·  stylo-ortho.fr",
   0.5,5.72,12.33,0.48,size=18,color=WHITE,align=PP_ALIGN.CENTER)
tb(sl,"Mathéo Guzzi  ·  Pierre Lachat  ·  Valentin Imbault-Casset",
   0.5,6.28,12.33,0.42,size=14,color=SKY,align=PP_ALIGN.CENTER)
tb(sl,"Merci de votre attention",0.5,6.75,12.33,0.42,
   size=14,color=RGBColor(0x60,0x88,0xB8),italic=True,align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
output = "/home/user/Stylo/Analyse_du_marche_StyloOrtho.pptx"
prs.save(output)
print(f"✅ Saved: {output}  ({len(prs.slides)} slides)")
