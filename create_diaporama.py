from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
DARK_BLUE   = RGBColor(0x0A, 0x29, 0x4B)   # fond titre
MID_BLUE    = RGBColor(0x1A, 0x52, 0x76)   # accent
LIGHT_BLUE  = RGBColor(0xD0, 0xE8, 0xF5)   # fond contenu
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE      = RGBColor(0xFF, 0x7A, 0x00)   # chiffres clés
GRAY_LIGHT  = RGBColor(0xF0, 0xF4, 0xF8)
TEXT_DARK   = RGBColor(0x1A, 0x1A, 0x2E)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]   # completement vide

def add_rect(slide, x, y, w, h, fill_color, alpha=None):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    return shape

def add_textbox(slide, text, x, y, w, h,
                font_size=18, bold=False, color=TEXT_DARK,
                align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.word_wrap = wrap
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb

def slide_header(slide, title, subtitle=None, section_label=None):
    """Bande bleue en haut avec titre."""
    add_rect(slide, 0, 0, 13.33, 1.4, DARK_BLUE)
    add_textbox(slide, title, 0.4, 0.08, 11, 0.9,
                font_size=32, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_textbox(slide, subtitle, 0.4, 0.85, 11, 0.45,
                    font_size=16, color=RGBColor(0xB0, 0xD0, 0xE8), align=PP_ALIGN.LEFT)
    if section_label:
        add_textbox(slide, section_label, 10.8, 0.1, 2.3, 0.5,
                    font_size=12, color=ORANGE, bold=True, align=PP_ALIGN.RIGHT)
    # fond gris clair pour le corps
    add_rect(slide, 0, 1.4, 13.33, 6.1, GRAY_LIGHT)

def add_bullet_box(slide, bullets, x, y, w, h,
                   title=None, title_color=DARK_BLUE, bg=WHITE,
                   font_size=16, title_size=18):
    add_rect(slide, x, y, w, h, bg)
    offset = y + 0.12
    if title:
        add_textbox(slide, title, x+0.15, offset, w-0.3, 0.4,
                    font_size=title_size, bold=True, color=title_color)
        offset += 0.42
    tb = slide.shapes.add_textbox(Inches(x+0.15), Inches(offset),
                                   Inches(w-0.3), Inches(h-(offset-y)-0.15))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(3)
        run = p.add_run()
        run.text = b
        run.font.size = Pt(font_size)
        run.font.color.rgb = TEXT_DARK

def add_kpi(slide, value, label, x, y, w=2.6, h=1.2):
    add_rect(slide, x, y, w, h, DARK_BLUE)
    add_textbox(slide, value, x, y+0.05, w, 0.65,
                font_size=30, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    add_textbox(slide, label, x, y+0.65, w, 0.5,
                font_size=13, color=WHITE, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 1 — Couverture
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, DARK_BLUE)
add_rect(sl, 0, 5.2, 13.33, 2.3, MID_BLUE)
# Logo fictif / titre produit
add_textbox(sl, "STYLO ORTHO", 1, 1.2, 11.33, 1.5,
            font_size=60, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(sl, "Stylo correcteur d'orthographe", 1, 2.7, 11.33, 0.8,
            font_size=28, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(sl, "Assemblé en France  •  Marché de la papeterie intelligente", 1, 3.4, 11.33, 0.6,
            font_size=18, color=RGBColor(0xB0, 0xD0, 0xE8), align=PP_ALIGN.CENTER)
add_rect(sl, 4.5, 4.1, 4.33, 0.06, ORANGE)
add_textbox(sl, "ANALYSE DU MARCHÉ", 1, 5.4, 11.33, 0.7,
            font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(sl, "Mathéo Guzzi  •  Pierre Lachat  •  Valentin Imbault-Casset", 1, 6.1, 11.33, 0.5,
            font_size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 2 — Introduction
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "1. Introduction", section_label="INTRODUCTION")

add_textbox(sl,
    "Le stylo correcteur d'orthographe Stylo Ortho apporte une solution innovante "
    "au besoin quotidien de corriger ses fautes d'orthographe — en temps réel, sur papier, "
    "sans connexion internet.",
    0.5, 1.6, 12.33, 1.0, font_size=19, color=TEXT_DARK)

# 4 KPI
add_kpi(sl, "214", "répondants à l'enquête", 0.5, 2.8)
add_kpi(sl, "82,3 %", "Oui + Peut-être", 3.3, 2.8)
add_kpi(sl, "86,4 %", "jugent le produit utile", 6.1, 2.8)
add_kpi(sl, "99,90 €", "prix de vente public", 8.9, 2.8)

add_bullet_box(sl,
    ["Marché de la papeterie en France : 4,6 milliards d'euros en 2025",
     "107 millions d'articles vendus à la rentrée scolaire 2024",
     "Segment des stylos intelligents quasi inexistant en France → opportunité pionnière",
     "Produit assemblé en France (Nicomatic) — conformité normes européennes"],
    0.5, 4.3, 12.33, 2.9, title="Contexte marché", bg=WHITE, font_size=17)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 3 — Demande : Analyse quantitative
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "2.1 Analyse quantitative de la demande", section_label="DEMANDE")

add_textbox(sl,
    "Chiffres généraux sur le marché : poids en volume et en valeur, taux de croissance, tendance et quantités moyennes consommées.",
    0.5, 1.5, 12.33, 0.55, font_size=16, italic=True, color=MID_BLUE)

# Tableau chiffres marché
headers = ["Indicateur", "En valeur / Volume"]
rows = [
    ("Marché papeterie France (2025)", "4,6 milliards d'euros — +52,1 % de CA en août vs décembre (forte saisonnalité rentrée)"),
    ("Papeterie hors papier bureautique (2024)", "1,3 milliard d'euros — 107 millions d'articles vendus à la rentrée scolaire 2024"),
    ("Marché mondial (2024)", "147,5 milliards de dollars — TCAC +3,8 % jusqu'en 2034"),
    ("Fournitures scolaires mondiales", "40 % de la demande mondiale"),
]
y = 2.15
add_rect(sl, 0.5, y, 12.33, 0.4, DARK_BLUE)
add_textbox(sl, "Indicateur", 0.55, y+0.04, 5.5, 0.35, font_size=14, bold=True, color=WHITE)
add_textbox(sl, "En valeur / Volume", 6.1, y+0.04, 6.5, 0.35, font_size=14, bold=True, color=WHITE)
colors = [WHITE, GRAY_LIGHT]
for i, (ind, val) in enumerate(rows):
    y += 0.52
    add_rect(sl, 0.5, y, 12.33, 0.5, colors[i%2])
    add_textbox(sl, ind, 0.55, y+0.04, 5.5, 0.45, font_size=13, bold=True, color=DARK_BLUE)
    add_textbox(sl, val, 6.1, y+0.04, 6.5, 0.45, font_size=13, color=TEXT_DARK)

# Chiffre clé + tri à plat
y += 0.65
add_rect(sl, 0.5, y, 12.33, 1.15, MID_BLUE)
add_textbox(sl, "📊 Tri à plat — Fréquence des fautes d'orthographe (enquête n=214)",
            0.7, y+0.05, 12, 0.4, font_size=15, bold=True, color=WHITE)
add_textbox(sl,
    "45 % des répondants font des fautes souvent ou très souvent. "
    "C'est la cible directe et fonctionnelle du Stylo Ortho. "
    "Ce chiffre valide le besoin concret et quotidien du produit sur le marché.",
    0.7, y+0.45, 12, 0.6, font_size=14, color=LIGHT_BLUE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 4 — Demande : Analyse qualitative
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "2.2 Analyse qualitative de la demande", section_label="DEMANDE")

add_textbox(sl,
    "Analyse du comportement d'achat du consommateur en s'appuyant sur les facteurs explicatifs.",
    0.5, 1.5, 12.33, 0.5, font_size=16, italic=True, color=MID_BLUE)

# 4 colonnes acteurs
cols = [
    ("Acheteur", ["• Parents d'élèves (30–50 ans)", "• Élèves eux-mêmes (dès le collège)", "• Écoles et centres de formation", "• Entreprises"]),
    ("Utilisateur", ["• Élèves et étudiants (dès le collège)", "• Professionnels écrivant à la main", "• Personnes avec difficultés d'écriture", "• Apprenants d'une langue étrangère"]),
    ("Prescripteur", ["• Enseignants et éducateurs spécialisés"]),
    ("Influenceur", ["• Influenceurs éducatifs TikTok/Instagram", "• Médias parentaux et scolaires", "• Bouche-à-oreille scolaire"]),
]
col_w = 3.0
for i, (title, items) in enumerate(cols):
    x = 0.4 + i * 3.2
    add_bullet_box(sl, items, x, 2.1, col_w, 3.0,
                   title=title, title_color=WHITE,
                   bg=DARK_BLUE if i%2==0 else MID_BLUE,
                   font_size=14, title_size=16)
    # override title color to white done in add_bullet_box title_color param
    # but text also needs to be white
    # redo with white text
    # (already passed bg=DARK_BLUE, title_color=WHITE — text color still TEXT_DARK)

# fix text color — redo the boxes properly
# Clear and redo
# (we'll just rebuild with a helper)

def add_actor_box(slide, title, items, x, y, w=3.0, h=3.0, dark=True):
    bg = DARK_BLUE if dark else MID_BLUE
    add_rect(slide, x, y, w, h, bg)
    add_textbox(slide, title, x+0.1, y+0.1, w-0.2, 0.42,
                font_size=17, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    tb = slide.shapes.add_textbox(Inches(x+0.1), Inches(y+0.6),
                                   Inches(w-0.2), Inches(h-0.7))
    tb.word_wrap = True
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(5)
        r = p.add_run(); r.text = item
        r.font.size = Pt(14); r.font.color.rgb = WHITE

# Remove shapes added above and redo — simpler: just add on top
# (pptx layers, so just add the correct boxes)
for i, (title, items) in enumerate(cols):
    x = 0.4 + i * 3.2
    add_actor_box(sl, title, items, x, 2.1, col_w, 3.1, dark=(i%2==0))

# Bottom note
add_rect(sl, 0.4, 5.35, 12.53, 0.65, LIGHT_BLUE)
add_textbox(sl,
    "46,7 % des répondants ont déjà ressenti de la gêne à cause de leurs fautes (enquête). "
    "Le Stylo Ortho répond à un besoin fonctionnel concret qui lève ce frein.",
    0.6, 5.4, 12.1, 0.55, font_size=14, color=DARK_BLUE, bold=True)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 5 — Facteurs individuels
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Facteurs individuels", section_label="DEMANDE")

factors = [
    ("Les besoins",
     "Le besoin de corriger ses fautes est un besoin préexistant, souvent implicite et latent chez l'utilisateur. "
     "Le Stylo Ortho le rend explicite et actionnable."),
    ("Les motivations",
     "• Motivation hédoniste : réussir à l'école ou au travail, écrire sans honte\n"
     "• Motivation rationnelle : gagner du temps, éviter les erreurs en situation de manuscrit"),
    ("Les freins",
     "• Frein rationnel : prix de 99,90 € élevé par rapport à un stylo classique (1–5 €)\n"
     "• Frein psychologique : crainte de paraître assisté ou de dépendre d'un outil"),
    ("La perception",
     "Le design du Stylo Ortho (stylo classique avec écran discret) crée une image de produit simple et non intrusif."),
    ("Profil socio-démographique",
     "Âge cible : 30–50 ans (parents acheteurs) et 11–25 ans (élèves et étudiants utilisateurs)"),
]
y = 1.55
for i, (title, text) in enumerate(factors):
    add_rect(sl, 0.4, y, 12.53, 0.95, WHITE if i%2==0 else GRAY_LIGHT)
    add_textbox(sl, title, 0.5, y+0.05, 3.5, 0.42, font_size=15, bold=True, color=DARK_BLUE)
    add_textbox(sl, text, 4.0, y+0.05, 9.0, 0.85, font_size=14, color=TEXT_DARK)
    y += 0.98

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 6 — Facteurs collectifs
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Facteurs collectifs", section_label="DEMANDE")

factors_c = [
    ("La culture",
     "En France, la maîtrise de l'orthographe est un marqueur social fort. "
     "Faire des fautes est perçu comme une lacune — le Stylo Ortho répond directement à cette pression culturelle."),
    ("Les groupes de référence",
     "• Groupe d'appartenance : camarades de classe, collègues\n"
     "• Groupe d'aspiration : bons élèves, professionnels valorisant l'écrit soigné"),
    ("Le cycle de vie familiale",
     "• Famille avec enfants scolarisés : cible prioritaire — la rentrée scolaire déclenche l'achat\n"
     "• Jeune actif entrant dans la vie professionnelle : cible secondaire"),
]
y = 1.6
for i, (title, text) in enumerate(factors_c):
    h = 1.6
    add_rect(sl, 0.4, y, 12.53, h, WHITE if i%2==0 else GRAY_LIGHT)
    add_textbox(sl, title, 0.5, y+0.1, 3.5, 0.45, font_size=16, bold=True, color=DARK_BLUE)
    add_textbox(sl, text, 4.0, y+0.08, 9.0, h-0.2, font_size=15, color=TEXT_DARK)
    y += h + 0.1

# Tri croisé
add_rect(sl, 0.4, y, 12.53, 1.0, MID_BLUE)
add_textbox(sl, "📊 Tri croisé — Intérêt pour le produit selon le profil (enquête n=214)",
            0.6, y+0.05, 12, 0.38, font_size=15, bold=True, color=WHITE)
add_textbox(sl,
    "Les personnes faisant le plus souvent des fautes sont aussi les plus intéressées par le produit. "
    "46,9 % des parents sont prêts à acheter le Stylo Ortho pour leur enfant.",
    0.6, y+0.45, 12, 0.45, font_size=14, color=LIGHT_BLUE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 7 — Situations d'achat
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Situations d'achat", section_label="DEMANDE")

add_bullet_box(sl,
    ["Facteurs situationnels",
     "• La rentrée scolaire (août–septembre) : pic d'achat principal — saisonnalité forte (+52,1 % CA en août vs décembre)",
     "• Examens et concours : besoin ponctuel de correction en situation de manuscrit",
     "• Cadeau de Noël ou anniversaire : achat impulsion pour un proche"],
    0.4, 1.55, 12.53, 2.35,
    title="Contexte situationnel", bg=WHITE, font_size=16, title_size=18)

add_bullet_box(sl,
    ["Types d'achat",
     "• Achat de nouveauté : le Stylo Ortho est un produit inconnu — forte implication à l'achat",
     "• Achat réfléchi : prix de 99,90 € justifié par la valeur perçue (USP unique)",
     "• Achat prescrit : recommandation d'un enseignant ou d'un parent"],
    0.4, 4.05, 12.53, 2.2,
    title="Types d'achat", bg=GRAY_LIGHT, font_size=16, title_size=18)

add_rect(sl, 0.4, 6.35, 12.53, 0.7, DARK_BLUE)
add_textbox(sl,
    "📊 Tri à plat : 56,6 % utilisent le correcteur téléphone et 43,4 % le correcteur Word — "
    "mais aucun de ces outils ne fonctionne sur papier. Le Stylo Ortho est le seul outil qui fait ça sur papier.",
    0.6, 6.4, 12.1, 0.6, font_size=14, color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 8 — Offre : Produits et segments
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "3.1 Analyse des produits et segments", section_label="OFFRE")

# En-tête tableau
y = 1.55
add_rect(sl, 0.4, y, 12.53, 0.42, DARK_BLUE)
for x, w, label in [(0.42,2.8,"Segment"),(3.25,2.4,"Part de marché"),(5.68,2.5,"Évolution"),(8.2,4.7,"Position Stylo Ortho")]:
    add_textbox(sl, label, x, y+0.04, w, 0.35, font_size=13, bold=True, color=WHITE)

rows = [
    ("Stylos classiques (BIC, Pilot…)", "≈ 60–65 %", "Stable à légèrement en recul (–2 à –3 % / an)", "Concurrent indirect — segment classique"),
    ("Stylos effaçables (FriXion…)", "≈ 10–15 %", "En croissance — en vogue dans le scolaire", "Complémentaire : Stylo Ortho intègre encre effaçable"),
    ("Stylos premium / luxe", "≈ 5–8 % (valeur)", "Stable — niche", "Inspiration positionnement premium"),
    ("Fournitures scolaires France 2024", "332 M€ à la rentrée", "Recul –7 % en 2024 vs 2023", "Marché cible principal"),
    ("Instruments d'écriture intelligents", "Quasi inexistant", "Fort potentiel", "🎯 Positionnement pionnier — aucun concurrent direct"),
]
for i, (seg, pm, evol, pos) in enumerate(rows):
    y += 0.55
    bg = WHITE if i%2==0 else GRAY_LIGHT
    add_rect(sl, 0.4, y, 12.53, 0.53, bg)
    for x, w, txt in [(0.42,2.8,seg),(3.25,2.4,pm),(5.68,2.5,evol),(8.2,4.7,pos)]:
        add_textbox(sl, txt, x, y+0.05, w, 0.45, font_size=13,
                    color=DARK_BLUE if txt==seg else TEXT_DARK,
                    bold=(txt==seg))

add_rect(sl, 0.4, y+0.6, 12.53, 0.65, MID_BLUE)
add_textbox(sl,
    "📊 Tri à plat : 56,6 % utilisent le correcteur téléphone et 43,4 % le correcteur Word — "
    "mais aucun de ces outils ne fonctionne sur papier. Opportunité unique pour le Stylo Ortho.",
    0.6, y+0.65, 12.1, 0.55, font_size=14, color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 9 — Concurrents directs et indirects
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "3.2 Concurrents directs et indirects", section_label="OFFRE")

y = 1.55
add_rect(sl, 0.4, y, 12.53, 0.42, DARK_BLUE)
for x, w, lbl in [(0.42,1.4,"Type"),(1.85,1.7,"Concurrent"),(3.58,2.2,"PDM estimée"),(5.8,2.8,"Forces"),(8.63,4.1,"Limite face au Stylo Ortho")]:
    add_textbox(sl, lbl, x, y+0.04, w, 0.35, font_size=13, bold=True, color=WHITE)

rows_c = [
    ("Indirecte", "BIC", "≈ 20–25 % stylos France", "Notoriété mondiale, prix très bas, réseau massif", "Aucune technologie de correction intégrée"),
    ("Indirecte", "Pilot / Stabilo / Pentel", "≈ 15–20 % cumulés", "Innovation stylos effaçables (FriXion)", "Pas de capacité de correction orthographique"),
    ("Directe\n(numérique)", "Microsoft Word /\nGoogle Docs", "≈ 43,4 % utilisateurs", "Intégré aux outils bureautiques, gratuit", "Inutilisable sur papier ou en examen manuscrit"),
    ("Directe\n(numérique)", "Correcteur téléphone", "≈ 56,6 % utilisateurs", "Gratuit, accessible partout, instantané", "Ne fonctionne pas lors d'une rédaction manuscrite"),
    ("Directe\n(numérique)", "Grammarly / Antidote", "Niche", "Très performants à l'écrit numérique", "Réservés au numérique, payants, non utilisables sur papier"),
]
for i, (typ, conc, pdm, forces, limite) in enumerate(rows_c):
    y += 0.56
    bg = WHITE if i%2==0 else GRAY_LIGHT
    add_rect(sl, 0.4, y, 12.53, 0.54, bg)
    for x, w, txt, bld in [(0.42,1.4,typ,False),(1.85,1.7,conc,True),(3.58,2.2,pdm,False),(5.8,2.8,forces,False),(8.63,4.1,limite,False)]:
        add_textbox(sl, txt, x, y+0.04, w, 0.48, font_size=12,
                    color=DARK_BLUE if bld else TEXT_DARK, bold=bld)

add_rect(sl, 0.4, y+0.62, 12.53, 0.58, DARK_BLUE)
add_textbox(sl,
    "🎯 Le Stylo Ortho occupe un territoire inexploité. Même les personnes n'étant pas sa cible directe "
    "reconnaissent l'utilité du produit pour les autres — effet de prescription fort.",
    0.6, y+0.67, 12.1, 0.48, font_size=14, color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 10 — Analyse des distributeurs
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "3.3 Analyse des distributeurs", section_label="OFFRE")

y = 1.55
add_rect(sl, 0.4, y, 12.53, 0.42, DARK_BLUE)
for x, w, lbl in [(0.42,2.5,"Distributeur"),(2.95,1.8,"Circuit"),(4.78,1.8,"PDM valeur"),(6.6,1.7,"Évolution"),(8.33,4.6,"Politique Stylo Ortho")]:
    add_textbox(sl, lbl, x, y+0.04, w, 0.35, font_size=13, bold=True, color=WHITE)

rows_d = [
    ("Bureau Vallée, Cultura, Fnac", "Spécialistes\n(circuit court)", "≈ 15–20 %", "+3 % (2020)", "✅ Priorité 1 — Démonstration possible en rayon. Distribution sélective au lancement."),
    ("Amazon, Fnac.com", "E-commerce\n(ultra-court)", "≈ 49 %\ntransactions", "+15 %/an", "✅ Priorité 2 — Distribution directe via stylo-ortho.fr + Amazon. Marge maîtrisée."),
    ("Carrefour, E.Leclerc", "Grandes surfaces\n(circuit long)", "332 M€ rentrée 2024", "–7 % (2024)", "⏳ Phase 2 uniquement — après notoriété établie."),
    ("Établissements scolaires / B2B", "Circuit direct\n(sélectif)", "Non quantifié\n(fort potentiel)", "En développement", "🤝 Partenariat Ministère Éducation nationale — phase de test."),
]
for i, (dist, circ, pdm, evol, pol) in enumerate(rows_d):
    y += 0.68
    bg = WHITE if i%2==0 else GRAY_LIGHT
    add_rect(sl, 0.4, y, 12.53, 0.66, bg)
    for x, w, txt, bld in [(0.42,2.5,dist,True),(2.95,1.8,circ,False),(4.78,1.8,pdm,False),(6.6,1.7,evol,False),(8.33,4.6,pol,False)]:
        add_textbox(sl, txt, x, y+0.05, w, 0.58, font_size=12,
                    color=DARK_BLUE if bld else TEXT_DARK, bold=bld)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 11 — PESTEL
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "3.4 Analyse PESTEL", section_label="OFFRE")

y = 1.55
add_rect(sl, 0.4, y, 12.53, 0.42, DARK_BLUE)
add_textbox(sl, "Facteur", 0.42, y+0.04, 1.5, 0.35, font_size=13, bold=True, color=WHITE)
add_textbox(sl, "Opportunités — influence positive", 1.95, y+0.04, 5.3, 0.35, font_size=13, bold=True, color=RGBColor(0x7F,0xFF,0xB0))
add_textbox(sl, "Menaces / Risques — influence négative", 7.28, y+0.04, 5.6, 0.35, font_size=13, bold=True, color=RGBColor(0xFF,0xB0,0xB0))

pestel = [
    ("🏛 Politique",
     "Politiques publiques favorisant l'éducation et la lutte contre l'illettrisme",
     "Réglementations sur composants électroniques en milieu scolaire. Certifications obligatoires."),
    ("💶 Économique",
     "Marché en croissance sur le segment innovation. Budget scolaire des familles stable.",
     "Pouvoir d'achat sous pression — produit à 99,90 € positionné comme achat réfléchi."),
    ("👥 Socioculturel",
     "Importance de l'orthographe en France. Dyslexie (5–10 % population). 300 M de francophones.",
     "Résistance au changement de certains enseignants. Habitude du stylo classique."),
    ("🔬 Technologique",
     "Développement des outils numériques, innovation stylos intelligents, progrès composants électroniques.",
     "Évolution rapide de l'IA — risque de substitution numérique à moyen terme."),
    ("🌱 Écologique",
     "Produit rechargeable (USB-C) — faible empreinte vs stylos jetables.",
     "Demande de réduction du plastique. Normes de recyclabilité renforcées en Europe."),
    ("⚖ Légal",
     "Normes européennes maîtrisables. Assemblage Nicomatic en France facilite la conformité CE.",
     "Réglementation composants électroniques, conformité produits enfants. Dépôt de brevet nécessaire."),
]
for i, (fac, opp, men) in enumerate(pestel):
    y += 0.49
    bg = WHITE if i%2==0 else GRAY_LIGHT
    add_rect(sl, 0.4, y, 12.53, 0.47, bg)
    add_textbox(sl, fac, 0.42, y+0.04, 1.5, 0.40, font_size=12, bold=True, color=DARK_BLUE)
    add_textbox(sl, opp, 1.95, y+0.04, 5.3, 0.40, font_size=12, color=TEXT_DARK)
    add_textbox(sl, men, 7.28, y+0.04, 5.6, 0.40, font_size=12, color=TEXT_DARK)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 12 — Opportunités et Menaces
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Opportunités et Menaces", section_label="SYNTHÈSE")

add_bullet_box(sl,
    ["Le besoin croissant d'outils pédagogiques innovants et personnalisés",
     "Le segment des stylos intelligents est quasi inexistant en France → positionnement pionnier",
     "82,3 % des répondants intéressés ou potentiellement intéressés (enquête n=214)",
     "Assemblage en France (Nicomatic) — conformité CE et argument marketing fort",
     "Distribution sélective (Bureau Vallée, Fnac) → démonstration produit possible",
     "49 % des transactions papeterie passent par l'e-commerce → canal direct fort"],
    0.4, 1.55, 6.1, 5.55,
    title="✅ Opportunités", title_color=RGBColor(0x00,0x80,0x40), bg=RGBColor(0xE8,0xF8,0xEE),
    font_size=15, title_size=19)

add_bullet_box(sl,
    ["La concurrence des outils numériques (correcteurs, IA, GPT)",
     "Prix de 99,90 € perçu comme élevé vs stylo classique (1–5 €)",
     "Recul du marché papeterie : –7 % en 2024 vs 2023",
     "Résistance culturelle de certains enseignants au changement",
     "Évolution rapide de l'IA — risque de substitution à moyen terme",
     "Certifications et réglementations produits électroniques enfants"],
    6.63, 1.55, 6.1, 5.55,
    title="⚠️ Menaces", title_color=RGBColor(0xCC,0x00,0x00), bg=RGBColor(0xFD,0xED,0xED),
    font_size=15, title_size=19)

add_textbox(sl,
    "Le marché des stylos intelligents représente une opportunité innovante et encore peu exploitée dans la papeterie. "
    "Le Stylo Ortho est positionné pour en être le leader grâce à son USP unique : correction en temps réel sur papier, sans internet.",
    0.4, 6.7, 12.53, 0.65, font_size=15, color=DARK_BLUE, bold=True, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 13 — Concept produit
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Concept produit — Stylo Ortho", section_label="MARKETING")

add_textbox(sl,
    "Un stylo classique enrichi d'une technologie intégrée de correction orthographique "
    "en temps réel — discret, rechargeable, assemblé en France.",
    0.5, 1.55, 12.33, 0.75, font_size=18, italic=True, color=MID_BLUE)

chars = [
    ("Recharge USB-C", "15 min de charge", "Pratique pour les élèves et professionnels du quotidien"),
    ("Encre rechargeable", "Moderne et économique", "Plus pratique et économique que les stylos à jeter"),
    ("Gomme effaçable", "Au bout du stylo", "Pas besoin de correcteur blanc — usage fluide"),
    ("Écran intégré", "Affiche les corrections en rouge", "Correction immédiate en temps réel — USP unique sur le marché"),
    ("Design ergonomique", "Léger, usage longue durée", "Adapté à tous : élèves dès le collège, étudiants, professionnels"),
    ("4 couleurs", "Noir, Bleu, Rouge, Vert", "Personnalisation et adaptation aux préférences"),
]
row_h = 0.72
y = 2.45
for i, (car, desc, avan) in enumerate(chars):
    bg = WHITE if i%2==0 else GRAY_LIGHT
    add_rect(sl, 0.4, y, 12.53, row_h, bg)
    add_textbox(sl, car, 0.5, y+0.1, 2.8, 0.55, font_size=14, bold=True, color=DARK_BLUE)
    add_textbox(sl, desc, 3.35, y+0.1, 3.0, 0.55, font_size=14, color=MID_BLUE)
    add_textbox(sl, avan, 6.38, y+0.1, 6.4, 0.55, font_size=14, color=TEXT_DARK)
    y += row_h

add_kpi(sl, "86,4 %", "jugent le Stylo Ortho utile ou très utile (enquête)", 0.4, 6.55, 12.53, 0.75)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 14 — Segments
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Segmentation du marché", section_label="MARKETING")

segs = [
    ("🎯 Parents d'élèves", "30–50 ans, acheteurs principaux",
     "46,9 % prêts à acheter pour leur enfant (enquête)", "CIBLE PRINCIPALE", DARK_BLUE),
    ("🎯 Élèves et étudiants", "11–25 ans, utilisateurs directs",
     "45 % font des fautes souvent ou très souvent (enquête)", "CIBLE SECONDAIRE", MID_BLUE),
    ("Professionnels / Enseignants", "25–50 ans, prescripteurs",
     "75,2 % voient l'utilité du produit pour leurs élèves ou collègues (enquête)", "CIBLE TERTIAIRE", RGBColor(0x2E,0x7D,0x52)),
]
y = 1.6
for seg, profil, chiffre, badge, color in segs:
    add_rect(sl, 0.4, y, 12.53, 1.5, color)
    add_textbox(sl, badge, 10.5, y+0.1, 2.3, 0.38, font_size=13, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT)
    add_textbox(sl, seg, 0.6, y+0.1, 8.5, 0.55, font_size=20, bold=True, color=WHITE)
    add_textbox(sl, profil, 0.6, y+0.65, 5.5, 0.45, font_size=15, color=LIGHT_BLUE)
    add_textbox(sl, f"📊 {chiffre}", 6.2, y+0.65, 6.5, 0.45, font_size=15, color=ORANGE)
    y += 1.65

add_rect(sl, 0.4, y, 12.53, 1.3, GRAY_LIGHT)
add_textbox(sl, "Stratégie retenue : Indifférenciée",
            0.6, y+0.1, 8, 0.45, font_size=18, bold=True, color=DARK_BLUE)
add_textbox(sl,
    "Le Stylo Ortho s'adresse à toute personne ayant besoin d'écrire sans fautes — élèves, étudiants, "
    "professionnels, apprenants. La campagne de communication cible néanmoins en priorité les parents "
    "et les élèves (rentrée scolaire).",
    0.6, y+0.55, 12.1, 0.65, font_size=14, color=TEXT_DARK)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 15 — Ciblage, Positionnement
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Ciblage et Positionnement", section_label="MARKETING")

add_bullet_box(sl,
    ["Cible principale : Parents d'élèves (30–50 ans) — 46,9 % prêts à acheter (enquête)",
     "Cible secondaire : Élèves et étudiants (11–25 ans) — 45 % font souvent des fautes",
     "Cible tertiaire : Professionnels et enseignants — prescripteurs et ambassadeurs",
     "Stratégie indifférenciée : un seul produit, une seule gamme, pour tous les profils"],
    0.4, 1.55, 12.53, 2.4,
    title="Ciblage", bg=WHITE, font_size=16, title_size=19)

add_rect(sl, 0.4, 4.1, 12.53, 0.48, DARK_BLUE)
add_textbox(sl, "Positionnement", 0.55, 4.15, 12, 0.38, font_size=19, bold=True, color=WHITE)

add_rect(sl, 0.4, 4.58, 12.53, 1.3, LIGHT_BLUE)
add_textbox(sl,
    "Dans la tête du consommateur, le Stylo Ortho doit être perçu comme :",
    0.6, 4.63, 12.1, 0.42, font_size=16, color=DARK_BLUE)
add_textbox(sl,
    "« Le stylo intelligent qui corrige mes fautes quand j'écris à la main »",
    0.6, 5.05, 12.1, 0.72, font_size=22, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

add_bullet_box(sl,
    ["Axe fonctionnel : correction orthographique en temps réel sur papier — USP unique",
     "Axe émotionnel : confiance en soi, réussite scolaire et professionnelle",
     "Axe symbolique : produit made in France, innovant, discret et non intrusif"],
    0.4, 6.0, 12.53, 1.35,
    bg=GRAY_LIGHT, font_size=15)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 16 — Bénéfice consommateur
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "Bénéfice consommateur", section_label="MARKETING")

benefits = [
    ("Bénéfice fonctionnel\n(USP — utilité)",
     "Ce que le produit fait concrètement",
     "Corriger les fautes d'orthographe en temps réel sur papier, sans connexion, "
     "en situation de manuscrit (école, bureau, examen)."),
    ("Bénéfice émotionnel\n(ressenti)",
     "Ce que l'utilisateur ressent",
     "Confiance en soi, réduction de la honte liée aux fautes, "
     "sérénité lors des examens et prises de notes importantes."),
    ("Bénéfice symbolique\n(image et statut)",
     "Ce que le produit dit de soi",
     "Utiliser un produit made in France innovant. Montrer qu'on valorise "
     "l'écrit soigné tout en adoptant les nouvelles technologies de manière discrète."),
]
y = 1.6
for typ, defn, appli in benefits:
    add_rect(sl, 0.4, y, 12.53, 1.45, WHITE if benefits.index((typ,defn,appli))%2==0 else GRAY_LIGHT)
    add_rect(sl, 0.4, y, 3.5, 1.45, DARK_BLUE)
    add_textbox(sl, typ, 0.55, y+0.2, 3.2, 1.0, font_size=15, bold=True, color=WHITE)
    add_textbox(sl, defn, 4.0, y+0.08, 3.0, 0.42, font_size=13, italic=True, color=MID_BLUE)
    add_textbox(sl, appli, 4.0, y+0.52, 8.7, 0.85, font_size=15, color=TEXT_DARK)
    y += 1.55

add_rect(sl, 0.4, y, 12.53, 1.0, ORANGE)
add_textbox(sl, "Bénéfice consommateur synthétique :", 0.6, y+0.08, 12, 0.38, font_size=15, bold=True, color=WHITE)
add_textbox(sl,
    "« Avec le Stylo Ortho, j'écris à la main sans stresser pour mes fautes — "
    "en cours, au bureau ou en examen, c'est le seul outil qui fait ça sur papier. »",
    0.6, y+0.48, 12, 0.42, font_size=16, bold=True, color=WHITE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 17 — Marketing Mix : Produit
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "P1 — Produit", section_label="MARKETING MIX")

add_textbox(sl,
    "Le Stylo Ortho est un stylo intelligent qui aide à écrire sans fautes. "
    "Il ressemble à un stylo classique, mais il contient une petite technologie intégrée.",
    0.5, 1.55, 12.33, 0.65, font_size=17, italic=True, color=MID_BLUE)

# Tableau caractéristiques
y = 2.3
add_rect(sl, 0.4, y, 12.53, 0.4, DARK_BLUE)
for x, w, lbl in [(0.42,2.5,"Caractéristique"),(2.95,3.3,"Description"),(6.28,6.6,"Avantage consommateur (USP)")]:
    add_textbox(sl, lbl, x, y+0.04, w, 0.33, font_size=13, bold=True, color=WHITE)

chars2 = [
    ("Recharge USB-C", "15 minutes de charge", "Pratique pour les élèves et professionnels du quotidien"),
    ("Encre rechargeable", "Encre moderne rechargeable, économique sur le long terme", "Plus pratique et économique que les stylos à jeter"),
    ("Gomme effaçable", "Petite gomme au bout du stylo pour effacer et réécrire directement", "Pas besoin de correcteur blanc, usage fluide"),
    ("Écran intégré", "Affiche les corrections en rouge et propose la bonne orthographe", "Correction immédiate en temps réel — USP unique sur le marché"),
    ("Design ergonomique", "Léger, conçu pour une utilisation longue sans fatigue", "Adapté à tous : élèves dès le collège, étudiants, professionnels"),
    ("4 couleurs disponibles", "Noir, Bleu, Rouge, Vert", "Personnalisation selon les préférences"),
]
for i, (c, d, a) in enumerate(chars2):
    y += 0.55
    bg = WHITE if i%2==0 else GRAY_LIGHT
    add_rect(sl, 0.4, y, 12.53, 0.53, bg)
    add_textbox(sl, c, 0.42, y+0.06, 2.5, 0.43, font_size=13, bold=True, color=DARK_BLUE)
    add_textbox(sl, d, 2.95, y+0.06, 3.3, 0.43, font_size=12, color=TEXT_DARK)
    add_textbox(sl, a, 6.28, y+0.06, 6.5, 0.43, font_size=12, color=TEXT_DARK)

# Partenaires
add_rect(sl, 0.4, y+0.6, 12.53, 0.72, LIGHT_BLUE)
add_textbox(sl, "Partenaires industriels : Shenzhen Apec (électronique 🇨🇳)  •  Pentel / Aihao (encre 🇯🇵🇨🇳)  •  Nicomatic assemblage final 🇫🇷",
            0.6, y+0.65, 12, 0.6, font_size=14, color=DARK_BLUE)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 18 — Marketing Mix : Prix
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "P2 — Prix", section_label="MARKETING MIX")

# KPI prix
add_kpi(sl, "99,90 €", "Prix de vente public\n(Standard)", 0.5, 1.6, 3.9, 1.6)
add_kpi(sl, "44,90 €", "Prix psychologique\n(promotionnel rentrée)", 4.7, 1.6, 3.9, 1.6)
add_kpi(sl, "82,3 %", "des répondants\nintéressés (enquête)", 8.9, 1.6, 3.9, 1.6)

add_bullet_box(sl,
    ["Version unique : Stylo Ortho Standard à 99,90 € — cible principale : élèves dès le collège, étudiants, parents",
     "82,3 % des répondants sont intéressés ou potentiellement intéressés (enquête n=214)",
     "Le prix de vente de 99,90 € est justifié par la valeur fonctionnelle unique (USP) : "
     "aucun concurrent sur papier",
     "Le prix psychologique de 44,90 € sera utilisé comme prix promotionnel à la rentrée "
     "pour déclencher les premiers achats (offre de lancement)"],
    0.4, 3.5, 12.53, 3.55,
    title="Stratégie de prix", bg=WHITE, font_size=16, title_size=19)

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 19 — Marketing Mix : Distribution
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
slide_header(sl, "P3 — Distribution", section_label="MARKETING MIX")

dist_data = [
    ("✅ Priorité 1",
     "Magasins spécialisés (circuit court)",
     "Bureau Vallée, Fnac, Cultura, librairies scolaires",
     "Le vendeur peut présenter et expliquer le produit en rayon. "
     "Essentiel pour un produit innovant nécessitant une démonstration."),
    ("✅ Priorité 2",
     "Vente en ligne (circuit ultra-court)",
     "Site e-commerce propre + Amazon + Fnac.com",
     "Indispensable pour toucher les parents qui achètent des fournitures scolaires en ligne. "
     "Marge maîtrisée. 49 % des transactions papeterie passent par l'e-commerce."),
    ("⏳ Phase 2",
     "Grandes surfaces (circuit long)",
     "Carrefour, E.Leclerc",
     "En phase 2 uniquement, une fois la notoriété du produit établie. "
     "Recul de –7 % en 2024 — approche prudente justifiée."),
    ("🤝 B2B",
     "Établissements scolaires / Circuit direct",
     "Lycées, collèges, centres de formation",
     "Partenariat Ministère Éducation nationale (phase de test). "
     "Objectif : 500+ stylos vendus en Année 1, 2 000+ en Année 2."),
]
y = 1.6
for badge, canal, points, why in dist_data:
    h = 1.18
    add_rect(sl, 0.4, y, 12.53, h, WHITE if dist_data.index((badge,canal,points,why))%2==0 else GRAY_LIGHT)
    add_rect(sl, 0.4, y, 1.8, h, DARK_BLUE)
    add_textbox(sl, badge, 0.5, y+0.35, 1.6, 0.48, font_size=14, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    add_textbox(sl, canal, 2.3, y+0.08, 4.0, 0.42, font_size=15, bold=True, color=DARK_BLUE)
    add_textbox(sl, points, 2.3, y+0.52, 4.0, 0.55, font_size=13, color=MID_BLUE, italic=True)
    add_textbox(sl, why, 6.4, y+0.1, 6.4, 0.95, font_size=13, color=TEXT_DARK)
    y += h + 0.07

# ─────────────────────────────────────────────────────────────────────────────
# DIAPO 20 — Conclusion / Slogan
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_rect(sl, 0, 0, 13.33, 7.5, DARK_BLUE)
add_rect(sl, 0, 4.8, 13.33, 2.7, MID_BLUE)

add_textbox(sl, "CONCLUSION", 1, 0.5, 11.33, 0.7,
            font_size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)

add_textbox(sl,
    "Le Stylo Ortho est le premier stylo correcteur d'orthographe en temps réel sur papier. "
    "Il répond à un besoin réel, validé par 214 répondants, sur un marché pionnier.",
    1, 1.25, 11.33, 1.0, font_size=19, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

add_rect(sl, 2.5, 2.45, 8.33, 0.06, ORANGE)

add_textbox(sl, "Notre slogan :", 1, 2.65, 11.33, 0.5,
            font_size=18, color=RGBColor(0xB0,0xD0,0xE8), align=PP_ALIGN.CENTER)
add_textbox(sl,
    "« Écrire juste, à la main. »",
    1, 3.15, 11.33, 1.1,
    font_size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_textbox(sl, "Stylo Ortho  •  Assemblé en France  •  stylo-ortho.fr",
            1, 5.2, 11.33, 0.6,
            font_size=18, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(sl,
    "82,3 % intéressés  •  86,4 % trouvent le produit utile  •  99,90 € prix de lancement",
    1, 5.9, 11.33, 0.5,
    font_size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(sl, "Mathéo Guzzi  •  Pierre Lachat  •  Valentin Imbault-Casset",
            1, 6.55, 11.33, 0.5,
            font_size=13, color=RGBColor(0x80,0xA0,0xC0), align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
output_path = "/home/user/Stylo/Analyse_du_marche_StyloOrtho.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
print(f"Slides: {len(prs.slides)}")
