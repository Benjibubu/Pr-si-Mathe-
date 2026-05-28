"""
SpaceX Praesentation - Julius Mueller & Ben Hammen
Dark Space Design, minimalistisch, ca. 30 Minuten Vortrag.
Generiert: spacex_praesentation.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import random
import copy

# ===== Farb-Palette (Dark Space / SpaceX-Look) =====
BG_DARK       = RGBColor(0x05, 0x08, 0x12)   # tiefes Schwarzblau
BG_DARKER     = RGBColor(0x00, 0x00, 0x00)
ACCENT_WHITE  = RGBColor(0xF5, 0xF7, 0xFA)
ACCENT_GREY   = RGBColor(0x9A, 0xA3, 0xB2)
ACCENT_RED    = RGBColor(0xE6, 0x1E, 0x2B)   # SpaceX-Rot
ACCENT_BLUE   = RGBColor(0x4A, 0x9E, 0xFF)   # Plasma-Blau
ACCENT_ORANGE = RGBColor(0xFF, 0x7A, 0x2A)   # Raketenflamme
LINE_GREY     = RGBColor(0x2B, 0x33, 0x44)

# ===== Folien-Geometrie (16:9) =====
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]  # Blank-Layout


# =====================================================
# Hilfs-Funktionen
# =====================================================
def add_bg(slide, color=BG_DARK):
    """Fuellt die Folie mit einer Hintergrundfarbe."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.line.fill.background()
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.shadow.inherit = False
    # Nach hinten schicken
    spTree = bg._element.getparent()
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)
    return bg


def add_stars(slide, count=60, seed=None):
    """Streut kleine weisse Punkte (Sterne) ueber die Folie."""
    if seed is not None:
        random.seed(seed)
    for _ in range(count):
        x = Emu(random.randint(0, int(SLIDE_W)))
        y = Emu(random.randint(0, int(SLIDE_H)))
        size_pt = random.choice([1, 1, 1, 2, 2, 3])
        size = Emu(int(size_pt * 9525))  # pt -> emu approx
        star = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
        star.line.fill.background()
        star.fill.solid()
        # variierende Helligkeit
        v = random.choice([0x55, 0x88, 0xAA, 0xCC, 0xEE, 0xFF])
        star.fill.fore_color.rgb = RGBColor(v, v, v)


def add_text(slide, left, top, width, height,
             text, size=24, bold=False, color=ACCENT_WHITE,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font="Calibri Light", italic=False, letter_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor

    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        f = r.font
        f.name = font
        f.size = Pt(size)
        f.bold = bold
        f.italic = italic
        f.color.rgb = color
        if letter_spacing is not None:
            rPr = r._r.get_or_add_rPr()
            rPr.set("spc", str(letter_spacing))
    return box


def add_line(slide, x1, y1, x2, y2, color=ACCENT_RED, weight=2.0):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = color
    line.line.width = Pt(weight)
    return line


def add_rect(slide, left, top, width, height,
             fill=None, line=None, line_weight=1.0, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, left, top, width, height)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_weight)
    s.shadow.inherit = False
    return s


def add_image_placeholder(slide, left, top, width, height,
                          label="BILD",
                          hint="",
                          fill=RGBColor(0x12, 0x18, 0x26),
                          border=ACCENT_RED):
    """Eleganter Bild-Platzhalter mit Beschriftung."""
    add_rect(slide, left, top, width, height, fill=fill, line=border, line_weight=1.25)
    # diagonale Linien (Foto-Symbol-Anmutung)
    add_line(slide, left, top, left + width, top + height,
             color=RGBColor(0x22, 0x2A, 0x3D), weight=0.5)
    add_line(slide, left, top + height, left + width, top,
             color=RGBColor(0x22, 0x2A, 0x3D), weight=0.5)
    # Label
    lbl = add_rect(slide, left, top, Inches(1.1), Inches(0.32),
                   fill=border, line=None)
    lbl_tb = add_text(slide, left, top, Inches(1.1), Inches(0.32),
                      label, size=10, bold=True, color=ACCENT_WHITE,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                      font="Calibri", letter_spacing=200)
    # Hint-Text
    if hint:
        add_text(slide, left + Inches(0.2), top + height - Inches(0.5),
                 width - Inches(0.4), Inches(0.4),
                 hint, size=10, color=ACCENT_GREY, italic=True,
                 anchor=MSO_ANCHOR.BOTTOM)


def add_footer(slide, page_num, total, chapter=""):
    """Footer mit Seitenzahl, Kapitel und Autoren."""
    # duenne Trennlinie
    add_line(slide, Inches(0.5), Inches(7.05), Inches(12.83), Inches(7.05),
             color=LINE_GREY, weight=0.75)
    # links: Kapitel
    if chapter:
        add_text(slide, Inches(0.5), Inches(7.12), Inches(6), Inches(0.3),
                 chapter.upper(), size=9, color=ACCENT_GREY,
                 letter_spacing=300, font="Calibri")
    # mitte: Autoren
    add_text(slide, Inches(5), Inches(7.12), Inches(3.33), Inches(0.3),
             "JULIUS MÜLLER  ·  BEN HAMMEN", size=9, color=ACCENT_GREY,
             align=PP_ALIGN.CENTER, letter_spacing=300, font="Calibri")
    # rechts: Seitenzahl
    add_text(slide, Inches(11.33), Inches(7.12), Inches(1.5), Inches(0.3),
             f"{page_num:02d} / {total:02d}", size=9, color=ACCENT_GREY,
             align=PP_ALIGN.RIGHT, letter_spacing=300, font="Calibri")


def add_top_marker(slide, kapitel_nr, kapitel_name):
    """Kleines Kapitel-Tag oben links."""
    bar = add_rect(slide, Inches(0.5), Inches(0.5), Inches(0.05), Inches(0.35),
                   fill=ACCENT_RED, line=None)
    add_text(slide, Inches(0.7), Inches(0.5), Inches(8), Inches(0.4),
             f"{kapitel_nr:02d}   ·   {kapitel_name.upper()}",
             size=10, color=ACCENT_GREY, letter_spacing=400, font="Calibri",
             anchor=MSO_ANCHOR.MIDDLE)


def add_title(slide, top, text, size=44, color=ACCENT_WHITE, align=PP_ALIGN.LEFT):
    add_text(slide, Inches(0.5), top, Inches(12.33), Inches(1.2),
             text, size=size, bold=True, color=color, align=align,
             font="Calibri Light")


def add_kicker(slide, top, text, color=ACCENT_RED):
    add_text(slide, Inches(0.5), top, Inches(12), Inches(0.4),
             text.upper(), size=11, bold=True, color=color,
             letter_spacing=400, font="Calibri")


def add_underline(slide, top, length=Inches(0.6), color=ACCENT_RED):
    add_rect(slide, Inches(0.5), top, length, Inches(0.05),
             fill=color, line=None)


# ===== Pictogramm-Hilfen (einfache geometrische Icons) =====
def icon_rocket(slide, cx, cy, size=Inches(1.0), color=ACCENT_WHITE):
    w = size
    h = size
    # Koerper
    body = add_rect(slide, cx - w/4, cy - h/2, w/2, h*0.65,
                    fill=color, line=None)
    # Spitze (Dreieck)
    tip = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                  cx - w/4, cy - h/2 - h*0.3, w/2, h*0.32)
    tip.fill.solid(); tip.fill.fore_color.rgb = color
    tip.line.fill.background()
    # Flossen
    f1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE,
                                 cx - w/2, cy + h*0.05, w/4, h*0.25)
    f1.fill.solid(); f1.fill.fore_color.rgb = color; f1.line.fill.background()
    f2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE,
                                 cx + w/4, cy + h*0.05, w/4, h*0.25)
    f2.fill.solid(); f2.fill.fore_color.rgb = color; f2.line.fill.background()
    f2.rotation = 270
    # Flamme
    flame = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                    cx - w/8, cy + h*0.25, w/4, h*0.25)
    flame.fill.solid(); flame.fill.fore_color.rgb = ACCENT_ORANGE
    flame.line.fill.background()
    flame.rotation = 180


def icon_circle_number(slide, cx, cy, diameter, number, color=ACCENT_RED):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                cx - diameter/2, cy - diameter/2,
                                diameter, diameter)
    c.fill.background()
    c.line.color.rgb = color
    c.line.width = Pt(1.5)
    add_text(slide, cx - diameter/2, cy - diameter/2, diameter, diameter,
             str(number), size=18, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================
# Slide-Tracker (Seitenzahlen werden am Ende gesetzt)
# =====================================================
SLIDES = []  # Liste von (slide, chapter_name)

def new_slide(chapter=""):
    s = prs.slides.add_slide(BLANK)
    add_bg(s)
    SLIDES.append((s, chapter))
    return s


# =====================================================
# SLIDE 1  -  TITEL
# =====================================================
def build_title_slide():
    s = new_slide("Titel")
    add_stars(s, count=120, seed=1)

    # roter Akzent-Balken links
    add_rect(s, Inches(0.5), Inches(2.8), Inches(0.08), Inches(1.6),
             fill=ACCENT_RED, line=None)

    # Kicker
    add_text(s, Inches(0.8), Inches(2.85), Inches(8), Inches(0.4),
             "REFERAT  ·  RAUMFAHRT  ·  ZUKUNFT",
             size=12, bold=True, color=ACCENT_RED,
             letter_spacing=500, font="Calibri")

    # Haupttitel
    add_text(s, Inches(0.8), Inches(3.3), Inches(12), Inches(1.3),
             "SpaceX", size=110, bold=True, color=ACCENT_WHITE,
             font="Calibri")

    # Untertitel
    add_text(s, Inches(0.8), Inches(4.7), Inches(12), Inches(0.6),
             "Eine private Firma. Eine Vision: Mars.",
             size=24, color=ACCENT_GREY, font="Calibri Light")

    # Autoren-Block unten
    add_line(s, Inches(0.8), Inches(6.4), Inches(4.5), Inches(6.4),
             color=ACCENT_RED, weight=1.5)
    add_text(s, Inches(0.8), Inches(6.5), Inches(6), Inches(0.4),
             "JULIUS MÜLLER   ·   BEN HAMMEN",
             size=14, bold=True, color=ACCENT_WHITE,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(0.8), Inches(6.85), Inches(6), Inches(0.35),
             "Mai 2026", size=11, color=ACCENT_GREY, font="Calibri")

    # rechts oben: Logo-Platzhalter
    add_image_placeholder(s, Inches(9.5), Inches(0.5), Inches(3.3), Inches(1.8),
                          label="LOGO",
                          hint="Bild: SpaceX-Logo (offiziell, transparent)")


# =====================================================
# SLIDE 2  -  BEGRUESSUNG
# =====================================================
def build_welcome_slide():
    s = new_slide("Begrüßung")
    add_stars(s, count=40, seed=2)

    add_kicker(s, Inches(0.9), "Willkommen")
    add_underline(s, Inches(1.25))

    add_text(s, Inches(0.5), Inches(1.8), Inches(12), Inches(2.5),
             "Heute fliegen wir\nmit SpaceX.",
             size=66, bold=True, color=ACCENT_WHITE,
             font="Calibri Light")

    # drei Versprechen / Hooks
    boxes = [
        ("01", "Was steckt hinter dem\nNamen SpaceX?"),
        ("02", "Wie verändert die Firma\ndie Raumfahrt?"),
        ("03", "Was bedeutet das\nfür uns – heute?"),
    ]
    box_w = Inches(3.8)
    box_h = Inches(1.7)
    gap   = Inches(0.25)
    total_w = box_w * 3 + gap * 2
    start = Inches(0.5)

    for i, (n, t) in enumerate(boxes):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(5.0), box_w, box_h,
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_text(s, x + Inches(0.3), Inches(5.15), Inches(1), Inches(0.4),
                 n, size=14, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, x + Inches(0.3), Inches(5.55), box_w - Inches(0.6),
                 Inches(1.1), t, size=15, color=ACCENT_WHITE,
                 font="Calibri Light")

    add_footer(s, 0, 0, "Begrüßung")


# =====================================================
# SLIDE 3  -  AGENDA
# =====================================================
def build_agenda_slide():
    s = new_slide("Agenda")
    add_kicker(s, Inches(0.9), "Themenübersicht")
    add_underline(s, Inches(1.25))
    add_title(s, Inches(1.6), "Unsere Reise", size=46)

    items_left = [
        ("01", "Was ist SpaceX?"),
        ("02", "Gründung & Geschichte"),
        ("03", "Elon Musk – der Kopf"),
        ("04", "Mission & Vision"),
        ("05", "Die Raketen-Familie"),
        ("06", "Wiederverwendbarkeit"),
    ]
    items_right = [
        ("07", "Dragon & Crew Dragon"),
        ("08", "Starlink – Internet aus dem All"),
        ("09", "Meilensteine"),
        ("10", "Zukunft: Mars & Mond"),
        ("11", "Was haben wir davon?"),
        ("12", "Kritik & Diskussion"),
    ]

    def render(col_items, x0):
        for i, (n, t) in enumerate(col_items):
            y = Inches(2.8 + i * 0.55)
            add_text(s, x0, y, Inches(0.6), Inches(0.4),
                     n, size=14, bold=True, color=ACCENT_RED,
                     font="Calibri", letter_spacing=200)
            add_text(s, x0 + Inches(0.8), y, Inches(5), Inches(0.4),
                     t, size=18, color=ACCENT_WHITE, font="Calibri Light")

    render(items_left, Inches(0.5))
    render(items_right, Inches(7.0))

    # Dauer-Hinweis
    add_text(s, Inches(0.5), Inches(6.4), Inches(8), Inches(0.4),
             "Dauer: ca. 30 Minuten   ·   Fragen am Ende",
             size=12, color=ACCENT_GREY, italic=True, font="Calibri")


# =====================================================
# SLIDE 4  -  WAS IST SPACEX  (Schlagwoerter)
# =====================================================
def build_what_is_spacex():
    s = new_slide("01 · Was ist SpaceX")
    add_top_marker(s, 1, "Was ist SpaceX?")
    add_title(s, Inches(1.1), "Eine Firma. Vier Worte.", size=46)
    add_underline(s, Inches(2.05))

    keywords = [
        ("PRIVAT",     "Kein Staat. Eine Firma."),
        ("WIEDER-\nVERWENDBAR", "Raketen landen. Erneut starten."),
        ("GÜNSTIG",    "Kosten pro kg drastisch gesenkt."),
        ("AMBITIONIERT","Ziel: Menschheit multi-planetar."),
    ]
    box_w = Inches(2.95)
    box_h = Inches(2.6)
    gap   = Inches(0.2)
    start = Inches(0.5)
    for i, (k, sub) in enumerate(keywords):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(2.8), box_w, box_h,
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        # Nummer dezent
        add_text(s, x + Inches(0.25), Inches(2.95), Inches(1), Inches(0.4),
                 f"0{i+1}", size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        # Keyword gross
        add_text(s, x + Inches(0.25), Inches(3.45), box_w - Inches(0.5),
                 Inches(1.4), k, size=28, bold=True, color=ACCENT_WHITE,
                 font="Calibri")
        # Beschreibung
        add_text(s, x + Inches(0.25), Inches(4.8), box_w - Inches(0.5),
                 Inches(0.9), sub, size=13, color=ACCENT_GREY,
                 font="Calibri Light")

    # Bild-Platzhalter rechts oben kleiner
    add_text(s, Inches(0.5), Inches(5.9), Inches(12), Inches(0.5),
             "Vollständiger Name: Space Exploration Technologies Corp.",
             size=14, italic=True, color=ACCENT_BLUE, font="Calibri Light")


# =====================================================
# SLIDE 5  -  GRUENDUNG & GESCHICHTE (Zeitstrahl)
# =====================================================
def build_history_slide():
    s = new_slide("02 · Geschichte")
    add_top_marker(s, 2, "Gründung & Geschichte")
    add_title(s, Inches(1.1), "Vom Garagen-Start\nzum Weltkonzern.", size=40)
    add_underline(s, Inches(2.6))

    # Zeitstrahl
    timeline_y = Inches(5.0)
    add_line(s, Inches(0.7), timeline_y, Inches(12.6), timeline_y,
             color=ACCENT_RED, weight=2.0)

    events = [
        ("2002", "Gründung\nin Kalifornien"),
        ("2008", "Falcon 1 –\nerster Privat-\norbit"),
        ("2012", "Dragon dockt\nan ISS an"),
        ("2015", "1. Landung\neiner Falcon 9"),
        ("2020", "Erste bemannte\nMission (Demo-2)"),
        ("2024", "Starship\nfängt Booster\nzurück"),
    ]
    n = len(events)
    span = Inches(11.9)
    start_x = Inches(0.7)
    step = span / (n - 1)

    for i, (year, label) in enumerate(events):
        x = start_x + step * i
        # Knoten
        dot = slide_dot = slide_dot = s.shapes.add_shape(
            MSO_SHAPE.OVAL, x - Inches(0.12), timeline_y - Inches(0.12),
            Inches(0.24), Inches(0.24))
        dot.fill.solid(); dot.fill.fore_color.rgb = ACCENT_RED
        dot.line.fill.background()
        # Jahr ueber Linie
        add_text(s, x - Inches(1), timeline_y - Inches(1.7),
                 Inches(2), Inches(0.5),
                 year, size=20, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, font="Calibri")
        # Beschreibung unter Linie
        add_text(s, x - Inches(1), timeline_y + Inches(0.25),
                 Inches(2), Inches(1.3),
                 label, size=11, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri Light")

    # Kontextzeile
    add_text(s, Inches(0.5), Inches(3.0), Inches(12), Inches(0.4),
             "Gründer: Elon Musk   ·   Startkapital: 100 Mio. $ aus PayPal-Verkauf",
             size=14, color=ACCENT_BLUE, italic=True, font="Calibri Light")


# =====================================================
# SLIDE 6  -  ELON MUSK
# =====================================================
def build_musk_slide():
    s = new_slide("03 · Elon Musk")
    add_top_marker(s, 3, "Elon Musk – der Kopf")
    add_title(s, Inches(1.1), "Der Mann hinter\nder Vision.", size=42)
    add_underline(s, Inches(2.7))

    # Foto-Platzhalter
    add_image_placeholder(s, Inches(0.5), Inches(3.2), Inches(4.5), Inches(3.5),
                          label="FOTO",
                          hint="Bild: Elon Musk (Portrait), CC-BY z.B. Wikipedia")

    # Fakten rechts
    facts = [
        ("Geboren",    "1971 in Pretoria, Südafrika"),
        ("Rolle",      "Gründer, CEO & Chief Engineer"),
        ("Andere Firmen", "Tesla · Neuralink · xAI · X"),
        ("Vermögen",   "geschätzt > 300 Mrd. $ (Forbes 2025)"),
        ("Motto",      "„Make humanity multiplanetary.“"),
    ]
    for i, (k, v) in enumerate(facts):
        y = Inches(3.3 + i * 0.65)
        add_text(s, Inches(5.5), y, Inches(2.2), Inches(0.5),
                 k.upper(), size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, Inches(7.7), y, Inches(5.5), Inches(0.5),
                 v, size=16, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 7  -  MISSION & VISION
# =====================================================
def build_mission_slide():
    s = new_slide("04 · Vision")
    add_stars(s, count=80, seed=7)
    add_top_marker(s, 4, "Mission & Vision")
    add_title(s, Inches(1.1), "Warum gibt es SpaceX?", size=42)
    add_underline(s, Inches(2.0))

    # Grosses Zitat
    add_text(s, Inches(0.7), Inches(2.7), Inches(12), Inches(1.5),
             "„Die Menschheit muss\neine multiplanetare Spezies werden.“",
             size=34, italic=True, color=ACCENT_WHITE, font="Calibri Light")
    add_text(s, Inches(0.7), Inches(4.5), Inches(6), Inches(0.5),
             "— Elon Musk",
             size=14, color=ACCENT_GREY, font="Calibri")

    # 3 Saeulen
    pillars = [
        ("MISSION",  "Raumfahrt billiger,\nzuverlässiger, häufiger."),
        ("VISION",   "Eine selbsterhaltende\nKolonie auf dem Mars."),
        ("METHODE",  "Wiederverwendbare\nRaketen + Skalierung."),
    ]
    box_w = Inches(3.95)
    gap   = Inches(0.25)
    start = Inches(0.5)
    for i, (k, v) in enumerate(pillars):
        x = start + (box_w + gap) * i
        # Linie oben
        add_rect(s, x, Inches(5.3), Inches(0.5), Inches(0.04),
                 fill=ACCENT_RED, line=None)
        add_text(s, x, Inches(5.4), box_w, Inches(0.4),
                 k, size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=400, font="Calibri")
        add_text(s, x, Inches(5.85), box_w, Inches(1.2),
                 v, size=17, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 8  -  RAKETEN-FAMILIE (Uebersicht)
# =====================================================
def build_rocket_family_slide():
    s = new_slide("05 · Raketen")
    add_top_marker(s, 5, "Die Raketen-Familie")
    add_title(s, Inches(1.1), "Vier Raketen.\nVier Größenklassen.", size=40)
    add_underline(s, Inches(2.7))

    # vier Pictogramme + Hoehenvergleich
    rockets = [
        ("FALCON 1",     21,   "21 m",  "0,67 t",  ACCENT_GREY),
        ("FALCON 9",     70,   "70 m",  "22,8 t",  ACCENT_WHITE),
        ("FALCON HEAVY", 70,   "70 m",  "63,8 t",  ACCENT_BLUE),
        ("STARSHIP",    121,  "121 m", ">100 t",   ACCENT_RED),
    ]
    max_h = 121
    base_y = Inches(6.0)
    max_pic_h = Inches(3.0)
    col_w = Inches(2.8)
    gap = Inches(0.3)
    start = Inches(0.7)

    for i, (name, h, h_lbl, payload, color) in enumerate(rockets):
        x = start + (col_w + gap) * i
        pic_h = Emu(int(max_pic_h * (h / max_h)))
        # Raketensilhouette: Rechteck + Spitze
        body_w = Inches(0.5)
        body = add_rect(s, x + col_w/2 - body_w/2, base_y - pic_h,
                        body_w, pic_h, fill=color, line=None)
        tip = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                  x + col_w/2 - body_w/2,
                                  base_y - pic_h - Inches(0.4),
                                  body_w, Inches(0.45))
        tip.fill.solid(); tip.fill.fore_color.rgb = color
        tip.line.fill.background()

        # Basislinie
        add_line(s, x, base_y, x + col_w, base_y,
                 color=LINE_GREY, weight=0.75)

        # Beschriftungen
        add_text(s, x, base_y + Inches(0.1), col_w, Inches(0.4),
                 name, size=14, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, font="Calibri", letter_spacing=200)
        add_text(s, x, base_y + Inches(0.5), col_w, Inches(0.4),
                 f"Höhe: {h_lbl}", size=11, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(s, x, base_y + Inches(0.85), col_w, Inches(0.4),
                 f"Nutzlast (LEO): {payload}", size=11, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri")


# =====================================================
# SLIDE 9  -  FALCON 9 DETAIL
# =====================================================
def build_falcon9_slide():
    s = new_slide("05 · Falcon 9")
    add_top_marker(s, 5, "Falcon 9 – das Arbeitspferd")
    add_title(s, Inches(1.1), "Falcon 9", size=64)
    add_text(s, Inches(0.5), Inches(2.4), Inches(12), Inches(0.5),
             "Die meistgenutzte Rakete der Welt.",
             size=20, color=ACCENT_GREY, italic=True, font="Calibri Light")

    add_image_placeholder(s, Inches(0.5), Inches(3.2), Inches(5.5), Inches(3.5),
                          label="BILD",
                          hint="Bild/Video: Falcon 9 Start (von SpaceX Flickr, CC-BY-NC)")

    facts = [
        ("Höhe",       "70 m"),
        ("Triebwerke", "9 × Merlin (Stufe 1)"),
        ("Schub",      "7,6 MN beim Start"),
        ("Nutzlast LEO","22,8 t"),
        ("Wiederverwendung", "bis 20 Flüge pro Booster"),
        ("Kosten",     "ca. 67 Mio. $ pro Start"),
    ]
    for i, (k, v) in enumerate(facts):
        y = Inches(3.3 + i * 0.55)
        add_text(s, Inches(6.5), y, Inches(2.4), Inches(0.5),
                 k.upper(), size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, Inches(8.9), y, Inches(4.2), Inches(0.5),
                 v, size=15, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 10  -  WIEDERVERWENDBARKEIT  (Schluessel-Idee)
# =====================================================
def build_reusability_slide():
    s = new_slide("06 · Wiederverwendbarkeit")
    add_top_marker(s, 6, "Wiederverwendbarkeit")
    add_title(s, Inches(1.1), "Die Idee, die alles\nverändert hat.", size=42)
    add_underline(s, Inches(2.7))

    # Vergleichs-Bloecke
    add_rect(s, Inches(0.5), Inches(3.6), Inches(6), Inches(3.3),
             fill=RGBColor(0x12, 0x14, 0x1A), line=LINE_GREY)
    add_text(s, Inches(0.8), Inches(3.75), Inches(5), Inches(0.4),
             "VORHER", size=11, bold=True, color=ACCENT_GREY,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(0.8), Inches(4.15), Inches(5), Inches(0.6),
             "Rakete = Einweg",
             size=28, bold=True, color=ACCENT_WHITE, font="Calibri Light")
    add_text(s, Inches(0.8), Inches(5.0), Inches(5.4), Inches(2.0),
             "• jede Mission braucht\n   neue Hardware\n• ~10.000 $ / kg in den Orbit\n• wenige Starts pro Jahr",
             size=14, color=ACCENT_GREY, font="Calibri Light")

    add_rect(s, Inches(6.8), Inches(3.6), Inches(6), Inches(3.3),
             fill=RGBColor(0x18, 0x0E, 0x12), line=ACCENT_RED)
    add_text(s, Inches(7.1), Inches(3.75), Inches(5), Inches(0.4),
             "HEUTE (SPACEX)", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(7.1), Inches(4.15), Inches(5), Inches(0.6),
             "Booster landet zurück",
             size=28, bold=True, color=ACCENT_WHITE, font="Calibri Light")
    add_text(s, Inches(7.1), Inches(5.0), Inches(5.4), Inches(2.0),
             "• Booster fliegt 10 – 20 ×\n• ~1.500 $ / kg in den Orbit\n• > 130 Starts allein 2024",
             size=14, color=ACCENT_WHITE, font="Calibri Light")

    add_text(s, Inches(0.5), Inches(3.05), Inches(12), Inches(0.5),
             "≈ 85 % günstigere Kosten pro Kilogramm in den Orbit (Quelle: NASA / SpaceX 2024).",
             size=13, italic=True, color=ACCENT_BLUE, font="Calibri Light")


# =====================================================
# SLIDE 11  -  STARSHIP
# =====================================================
def build_starship_slide():
    s = new_slide("05 · Starship")
    add_stars(s, count=70, seed=11)
    add_top_marker(s, 5, "Starship – die Zukunft")
    add_title(s, Inches(1.1), "Starship.", size=72)
    add_text(s, Inches(0.5), Inches(2.6), Inches(12), Inches(0.5),
             "Die größte Rakete, die je gebaut wurde.",
             size=20, color=ACCENT_GREY, italic=True, font="Calibri Light")

    add_image_placeholder(s, Inches(7.5), Inches(3.4), Inches(5.3), Inches(3.5),
                          label="BILD",
                          hint="Bild: Starship Vollstack auf Starbase (z.B. SpaceX Flickr)")

    bullets = [
        ("121 m",   "Höher als die Freiheitsstatue."),
        (">100 t",  "Nutzlast in den niedrigen Orbit."),
        ("33",      "Raptor-Triebwerke in der Super Heavy."),
        ("Voll",    "wiederverwendbar – Booster & Schiff."),
        ("Ziel",    "Mond (NASA Artemis) & Mars."),
    ]
    for i, (k, v) in enumerate(bullets):
        y = Inches(3.4 + i * 0.62)
        add_text(s, Inches(0.5), y, Inches(2.0), Inches(0.5),
                 k, size=22, bold=True, color=ACCENT_RED,
                 font="Calibri")
        add_text(s, Inches(2.6), y + Inches(0.08), Inches(4.6), Inches(0.5),
                 v, size=14, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 12  -  DRAGON / CREW DRAGON
# =====================================================
def build_dragon_slide():
    s = new_slide("07 · Dragon")
    add_top_marker(s, 7, "Dragon & Crew Dragon")
    add_title(s, Inches(1.1), "Dragon – Fracht & Menschen.", size=40)
    add_underline(s, Inches(2.05))

    # Zwei Karten
    cards = [
        ("CARGO DRAGON", "Versorgt die ISS\nseit 2012.",
         "• bis zu 6.000 kg Fracht\n• automatisches Andocken\n• Rückkehr mit Experimenten"),
        ("CREW DRAGON",  "Bringt Astronauten\nins All – seit 2020.",
         "• Platz für 4 Personen\n• erste Privat-Crew zur ISS\n• Inspiration4 (2021, zivile Crew)"),
    ]
    for i, (k, sub, body) in enumerate(cards):
        x = Inches(0.5) + Inches(6.2) * i
        add_rect(s, x, Inches(2.8), Inches(6.0), Inches(4.0),
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_text(s, x + Inches(0.4), Inches(2.95), Inches(5), Inches(0.4),
                 k, size=12, bold=True, color=ACCENT_RED,
                 letter_spacing=400, font="Calibri")
        add_text(s, x + Inches(0.4), Inches(3.45), Inches(5), Inches(1.1),
                 sub, size=24, bold=True, color=ACCENT_WHITE,
                 font="Calibri Light")
        add_text(s, x + Inches(0.4), Inches(5.0), Inches(5.2), Inches(1.7),
                 body, size=13, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 13  -  STARLINK
# =====================================================
def build_starlink_slide():
    s = new_slide("08 · Starlink")
    add_stars(s, count=120, seed=13)
    add_top_marker(s, 8, "Starlink")
    add_title(s, Inches(1.1), "Starlink.", size=66)
    add_text(s, Inches(0.5), Inches(2.5), Inches(12), Inches(0.5),
             "Internet aus dem All – für die ganze Welt.",
             size=20, color=ACCENT_GREY, italic=True, font="Calibri Light")

    add_underline(s, Inches(3.2))

    stats = [
        ("> 7.000", "aktive Satelliten\n(Stand 2025)"),
        ("~ 550 km","Flughöhe\n(niedriger Orbit)"),
        ("100+",     "Länder mit\nverfügbarem Dienst"),
        ("~50 ms",   "Latenz – fast wie\nGlasfaser"),
    ]
    box_w = Inches(2.95)
    gap = Inches(0.2)
    start = Inches(0.5)
    for i, (n, t) in enumerate(stats):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(3.8), box_w, Inches(2.4),
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_text(s, x, Inches(4.0), box_w, Inches(1.2),
                 n, size=40, bold=True, color=ACCENT_BLUE,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(s, x, Inches(5.3), box_w, Inches(0.9),
                 t, size=13, color=ACCENT_WHITE, align=PP_ALIGN.CENTER,
                 font="Calibri Light")

    add_text(s, Inches(0.5), Inches(6.4), Inches(12), Inches(0.5),
             "Auch im Ukraine-Krieg eingesetzt – politisch & militärisch hoch relevant.",
             size=12, italic=True, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 14  -  INTERAKTIV: QUIZ
# =====================================================
def build_quiz_slide():
    s = new_slide("Interaktiv · Quiz")
    add_kicker(s, Inches(0.9), "Mitmach-Frage")
    add_underline(s, Inches(1.25))
    add_title(s, Inches(1.6), "Wie hoch ist\ndie Starship-Rakete?", size=44)

    options = [
        ("A", "72 m   – wie Falcon 9"),
        ("B", "98 m   – wie eine Saturn V"),
        ("C", "121 m – höher als die Freiheitsstatue"),
        ("D", "180 m – höher als der Kölner Dom"),
    ]
    for i, (l, t) in enumerate(options):
        y = Inches(4.0 + i * 0.55)
        add_rect(s, Inches(2.5), y, Inches(0.7), Inches(0.45),
                 fill=ACCENT_RED, line=None)
        add_text(s, Inches(2.5), y, Inches(0.7), Inches(0.45),
                 l, size=16, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font="Calibri")
        add_text(s, Inches(3.4), y + Inches(0.04), Inches(8), Inches(0.45),
                 t, size=18, color=ACCENT_WHITE, font="Calibri Light")

    add_text(s, Inches(0.5), Inches(6.7), Inches(12), Inches(0.4),
             "→ Klick zur Auflösung auf der nächsten Folie.",
             size=12, italic=True, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 15  -  MEILENSTEINE (Zahlen)
# =====================================================
def build_milestones_slide():
    s = new_slide("09 · Meilensteine")
    add_top_marker(s, 9, "Meilensteine in Zahlen")
    add_title(s, Inches(1.1), "Zahlen, die\nsprechen.", size=42)
    add_underline(s, Inches(2.7))

    stats = [
        ("450+",  "erfolgreiche\nFalcon-Starts"),
        ("~ 87 %","Anteil an allen\nweltweiten Orbit-\nStarts 2024"),
        ("> 50",  "Astronauten zur\nISS gebracht"),
        ("400+",  "Landungen einer\nFalcon-Erststufe"),
    ]
    box_w = Inches(2.95)
    gap   = Inches(0.2)
    start = Inches(0.5)
    for i, (n, t) in enumerate(stats):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(3.8), box_w, Inches(2.7),
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        # roter Akzent oben
        add_rect(s, x, Inches(3.8), Inches(0.45), Inches(0.06),
                 fill=ACCENT_RED, line=None)
        add_text(s, x, Inches(4.1), box_w, Inches(1.2),
                 n, size=44, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(s, x, Inches(5.5), box_w, Inches(1.0),
                 t, size=13, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri Light")

    add_text(s, Inches(0.5), Inches(6.7), Inches(12), Inches(0.4),
             "Quellen: SpaceX, NASA, BryceTech Startreport 2024.",
             size=11, italic=True, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 16  -  WETTBEWERB
# =====================================================
def build_competition_slide():
    s = new_slide("Wettbewerb")
    add_top_marker(s, 9, "Wer macht eigentlich Konkurrenz?")
    add_title(s, Inches(1.1), "Im Vergleich.", size=44)
    add_underline(s, Inches(2.05))

    # Tabellen-Header
    cols = ["Firma / Org.", "Heimat", "Stärke", "Schwäche"]
    rows = [
        ["SpaceX",        "USA (privat)",   "Wiederverwendbarkeit,\nhohe Taktung", "regulatorische\nKonflikte"],
        ["NASA",          "USA (staatlich)","Erfahrung, Wissenschaft","langsame Entwicklung,\nkostspielig"],
        ["Blue Origin",   "USA (privat)",   "starker Investor\n(Jeff Bezos)",      "wenige Orbit-Flüge"],
        ["ESA / Arianespace","Europa",      "Wissenschaftliche\nMissionen",         "wenig Wiederverwendbarkeit"],
        ["CNSA",          "China",          "wachsendes Programm,\neigene Raumstation","weniger transparent"],
    ]

    col_x = [Inches(0.5), Inches(3.4), Inches(6.3), Inches(9.2)]
    col_w = Inches(2.85)
    row_y0 = Inches(3.0)
    row_h  = Inches(0.7)

    # Header
    for j, c in enumerate(cols):
        add_text(s, col_x[j], row_y0, col_w, Inches(0.4),
                 c.upper(), size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
    add_line(s, Inches(0.5), row_y0 + Inches(0.45),
             Inches(12.83), row_y0 + Inches(0.45),
             color=ACCENT_RED, weight=1.0)

    for i, row in enumerate(rows):
        y = row_y0 + Inches(0.55) + row_h * i
        for j, cell in enumerate(row):
            color = ACCENT_WHITE if j == 0 else ACCENT_GREY
            size  = 14 if j == 0 else 11
            bold  = j == 0
            add_text(s, col_x[j], y, col_w, row_h,
                     cell, size=size, bold=bold, color=color,
                     font="Calibri Light")
        add_line(s, Inches(0.5), y + row_h - Inches(0.05),
                 Inches(12.83), y + row_h - Inches(0.05),
                 color=LINE_GREY, weight=0.5)


# =====================================================
# SLIDE 17  -  MARS-VISION
# =====================================================
def build_mars_slide():
    s = new_slide("10 · Mars")
    add_stars(s, count=80, seed=17)
    add_top_marker(s, 10, "Zukunft – Mars")
    add_title(s, Inches(1.1), "Ziel: Mars.", size=66)
    add_text(s, Inches(0.5), Inches(2.5), Inches(12), Inches(0.5),
             "Eine Stadt auf einem anderen Planeten – bis 2050?",
             size=20, color=ACCENT_GREY, italic=True, font="Calibri Light")

    add_image_placeholder(s, Inches(7.5), Inches(3.3), Inches(5.3), Inches(3.6),
                          label="ANIMATION/BILD",
                          hint="Bild: Mars-Oberfläche oder SpaceX-Konzept\n(NASA Image Library, CC0)")

    points = [
        ("PHASE 1", "Unbemannte Vorrats-\nFlüge mit Starship"),
        ("PHASE 2", "Erste Crew zum Mars\n(Ziel: Ende 2020er)"),
        ("PHASE 3", "Treibstoff-Produktion\naus CO₂ und Wasser"),
        ("PHASE 4", "Selbsterhaltende\nKolonie – ~1 Mio Menschen"),
    ]
    for i, (k, t) in enumerate(points):
        y = Inches(3.4 + i * 0.85)
        icon_circle_number(s, Inches(1.0), y + Inches(0.3),
                           Inches(0.55), i + 1, color=ACCENT_RED)
        add_text(s, Inches(1.7), y, Inches(2), Inches(0.4),
                 k, size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=400, font="Calibri")
        add_text(s, Inches(1.7), y + Inches(0.3), Inches(5.5), Inches(0.7),
                 t, size=14, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 18  -  MOND / ARTEMIS
# =====================================================
def build_moon_slide():
    s = new_slide("11 · Mond")
    add_top_marker(s, 11, "Mond – Artemis-Programm")
    add_title(s, Inches(1.1), "Zuerst aber:\nzurück zum Mond.", size=40)
    add_underline(s, Inches(2.7))

    add_image_placeholder(s, Inches(0.5), Inches(3.2), Inches(5), Inches(3.5),
                          label="BILD",
                          hint="Bild: Mond / Starship HLS Konzept (NASA)")

    add_text(s, Inches(6.0), Inches(3.2), Inches(6.8), Inches(0.5),
             "Artemis III (NASA, ab 2027)",
             size=20, bold=True, color=ACCENT_WHITE, font="Calibri Light")
    add_text(s, Inches(6.0), Inches(3.8), Inches(6.8), Inches(3),
             "• Erste Mondlandung seit 1972\n"
             "• SpaceX baut die Landefähre („Starship HLS“)\n"
             "• Vertragsvolumen: ca. 4 Mrd. $\n"
             "• Erstmals soll eine Frau auf dem Mond landen\n"
             "• Ziel: dauerhafte Mond-Basis am Südpol",
             size=14, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 19  -  PRAXIS: WAS HABEN WIR DAVON?
# =====================================================
def build_use_for_us_slide():
    s = new_slide("12 · Praxis")
    add_top_marker(s, 12, "Was haben wir davon?")
    add_title(s, Inches(1.1), "Was bringt uns das\nhier auf der Erde?", size=38)
    add_underline(s, Inches(2.7))

    items = [
        ("📡", "Internet überall",
         "Starlink bringt Internet in Dörfer, Wüsten\noder Krisengebiete – ohne Kabel."),
        ("💾", "Wissenschaft auf der ISS",
         "Experimente in Schwerelosigkeit liefern Daten\nfür Medizin, Werkstoffe, Biologie."),
        ("🌍", "Erdbeobachtung",
         "Wetter, Klima, Landwirtschaft – immer mehr\nkleine Satelliten machen Daten billiger."),
        ("🚀", "Günstiger Zugang ins All",
         "Forschung von Unis und Start-ups wird durch\nfallende Startpreise plötzlich machbar."),
        ("💼", "Neue Berufe",
         "Raumfahrtindustrie wächst – Ingenieure, Informatiker,\nMaterialwissenschaftler werden gebraucht."),
    ]
    # zwei Spalten
    for i, (icon, k, body) in enumerate(items[:3]):
        y = Inches(3.2 + i * 1.2)
        add_text(s, Inches(0.5), y, Inches(0.7), Inches(0.7),
                 icon, size=28, color=ACCENT_RED, font="Calibri")
        add_text(s, Inches(1.2), y, Inches(5.3), Inches(0.4),
                 k, size=16, bold=True, color=ACCENT_WHITE,
                 font="Calibri")
        add_text(s, Inches(1.2), y + Inches(0.4), Inches(5.3), Inches(0.9),
                 body, size=12, color=ACCENT_GREY, font="Calibri Light")
    for i, (icon, k, body) in enumerate(items[3:]):
        y = Inches(3.8 + i * 1.2)
        add_text(s, Inches(7.0), y, Inches(0.7), Inches(0.7),
                 icon, size=28, color=ACCENT_RED, font="Calibri")
        add_text(s, Inches(7.7), y, Inches(5), Inches(0.4),
                 k, size=16, bold=True, color=ACCENT_WHITE, font="Calibri")
        add_text(s, Inches(7.7), y + Inches(0.4), Inches(5), Inches(0.9),
                 body, size=12, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 20  -  PRAXIS-TIPPS
# =====================================================
def build_practice_tips_slide():
    s = new_slide("12 · Praxis")
    add_top_marker(s, 12, "Praxis-Tipps · selbst aktiv werden")
    add_title(s, Inches(1.1), "Du willst tiefer rein?", size=42)
    add_underline(s, Inches(2.0))

    tips = [
        ("BEOBACHTEN",
         "Starlink-„Schlangen“ am Nachthimmel selbst sehen",
         "App: „Heavens-Above“ oder „Starlink Tracker“ – zeigt Überflüge live."),
        ("ZUSCHAUEN",
         "Raketenstarts live mitverfolgen",
         "SpaceX-Stream auf X / YouTube – kostenlos, mit Telemetrie & Triebwerks-Daten."),
        ("LERNEN",
         "Selbst eine kleine Rakete bauen",
         "Modellraketen (z. B. Estes), Programmieren mit „Kerbal Space Program“."),
        ("MITREDEN",
         "Quellenkritik üben",
         "Vergleicht Aussagen von SpaceX mit NASA, ESA & unabhängigen Medien."),
    ]
    box_w = Inches(2.95)
    gap   = Inches(0.2)
    start = Inches(0.5)
    for i, (k, head, body) in enumerate(tips):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(2.8), box_w, Inches(4.0),
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_rect(s, x, Inches(2.8), Inches(0.5), Inches(0.05),
                 fill=ACCENT_RED, line=None)
        add_text(s, x + Inches(0.3), Inches(2.95), Inches(2.6), Inches(0.4),
                 k, size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=400, font="Calibri")
        add_text(s, x + Inches(0.3), Inches(3.4), Inches(2.6), Inches(1.2),
                 head, size=16, bold=True, color=ACCENT_WHITE,
                 font="Calibri Light")
        add_text(s, x + Inches(0.3), Inches(4.8), Inches(2.6), Inches(1.9),
                 body, size=12, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# SLIDE 21  -  KRITIK
# =====================================================
def build_criticism_slide():
    s = new_slide("13 · Kritik")
    add_top_marker(s, 13, "Kritik & Risiken")
    add_title(s, Inches(1.1), "Nicht alles ist Glanz.", size=42)
    add_underline(s, Inches(2.0))

    items = [
        ("WELTRAUMMÜLL",
         "Über 7.000 Starlink-Satelliten\n– Risiko für Kollisionen."),
        ("LICHTVERSCHMUTZUNG",
         "Astronomen klagen über helle\nSpuren auf Teleskop-Bildern."),
        ("UMWELT",
         "Starship-Tests in Texas:\nLärm, Schutzgebiete, CO₂-Ausstoß."),
        ("MACHTKONZENTRATION",
         "Eine Firma dominiert\nfast den gesamten Raketenmarkt."),
        ("PERSON ELON MUSK",
         "Politische Aussagen polarisieren\n– vom CEO unabhängig?"),
        ("VERSPRECHEN VS. REALITÄT",
         "Mars „bis 2024“ ist mehrfach\nverschoben worden."),
    ]
    # 3x2 Grid
    box_w = Inches(4.05)
    box_h = Inches(1.8)
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)
    start_x = Inches(0.5)
    start_y = Inches(2.9)
    for i, (k, t) in enumerate(items):
        col = i % 3
        row = i // 3
        x = start_x + (box_w + gap_x) * col
        y = start_y + (box_h + gap_y) * row
        add_rect(s, x, y, box_w, box_h,
                 fill=RGBColor(0x18, 0x10, 0x12), line=ACCENT_RED, line_weight=0.5)
        add_text(s, x + Inches(0.3), y + Inches(0.2), box_w - Inches(0.6),
                 Inches(0.4), k, size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, x + Inches(0.3), y + Inches(0.7), box_w - Inches(0.6),
                 box_h - Inches(0.9),
                 t, size=13, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 22  -  DISKUSSIONSFRAGE
# =====================================================
def build_discussion_slide():
    s = new_slide("Interaktiv · Diskussion")
    add_stars(s, count=40, seed=22)
    add_kicker(s, Inches(0.9), "Diskutiert mit")
    add_underline(s, Inches(1.25))

    add_text(s, Inches(0.5), Inches(1.9), Inches(12.3), Inches(3),
             "Sollte eine einzige\nprivate Firma die\nRaumfahrt kontrollieren?",
             size=54, bold=True, color=ACCENT_WHITE, font="Calibri Light")

    # Pro/Contra Karten
    add_rect(s, Inches(0.5), Inches(5.4), Inches(6.0), Inches(1.7),
             fill=RGBColor(0x0C, 0x18, 0x10), line=ACCENT_BLUE, line_weight=0.75)
    add_text(s, Inches(0.8), Inches(5.55), Inches(5), Inches(0.4),
             "ARGUMENT: JA", size=11, bold=True, color=ACCENT_BLUE,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(0.8), Inches(5.95), Inches(5.5), Inches(1.0),
             "Schneller, günstiger, innovativer.\nStaat kann beauftragen, ohne selbst bauen zu müssen.",
             size=13, color=ACCENT_WHITE, font="Calibri Light")

    add_rect(s, Inches(6.8), Inches(5.4), Inches(6.0), Inches(1.7),
             fill=RGBColor(0x18, 0x0C, 0x10), line=ACCENT_RED, line_weight=0.75)
    add_text(s, Inches(7.1), Inches(5.55), Inches(5), Inches(0.4),
             "ARGUMENT: NEIN", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(7.1), Inches(5.95), Inches(5.5), Inches(1.0),
             "Abhängigkeit vom Willen einer Person.\nKritische Infrastruktur ohne demokratische Kontrolle.",
             size=13, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 23  -  FAZIT
# =====================================================
def build_summary_slide():
    s = new_slide("Fazit")
    add_top_marker(s, 14, "Fazit")
    add_title(s, Inches(1.1), "Drei Sätze zum Mitnehmen.", size=40)
    add_underline(s, Inches(2.0))

    points = [
        ("01", "SpaceX hat Raumfahrt billiger gemacht –\ndurch wiederverwendbare Raketen."),
        ("02", "Die Firma ist heute der größte Akteur\nim Orbit – mit allen Chancen und Risiken."),
        ("03", "Mond, Mars, Internet:\ndas nächste Jahrzehnt entscheidet, was real wird."),
    ]
    for i, (n, t) in enumerate(points):
        y = Inches(3.0 + i * 1.3)
        add_text(s, Inches(0.5), y, Inches(1.2), Inches(0.8),
                 n, size=42, bold=True, color=ACCENT_RED, font="Calibri")
        add_text(s, Inches(2.0), y + Inches(0.05), Inches(10.5), Inches(1.2),
                 t, size=22, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# SLIDE 24  -  QUELLEN
# =====================================================
def build_sources_slide():
    s = new_slide("Quellen")
    add_top_marker(s, 15, "Quellen")
    add_title(s, Inches(1.1), "Quellen & weiterführende Links", size=32)
    add_underline(s, Inches(2.0))

    sources = [
        ("SpaceX – Offizielle Seite",   "https://www.spacex.com"),
        ("NASA – Commercial Crew",      "https://www.nasa.gov/commercial-crew"),
        ("NASA – Artemis Program",      "https://www.nasa.gov/humans-in-space/artemis"),
        ("ESA – European Space Agency", "https://www.esa.int"),
        ("DLR – Deutsches Zentrum für Luft- und Raumfahrt", "https://www.dlr.de"),
        ("BryceTech – Start Reports",   "https://brycetech.com/reports"),
        ("Spiegel Online – Wissenschaft (Raumfahrt)", "https://www.spiegel.de/wissenschaft/weltall"),
        ("Tagesschau – Raumfahrt",      "https://www.tagesschau.de/wissen/raumfahrt"),
        ("Reuters – Space",             "https://www.reuters.com/business/aerospace-defense"),
        ("BBC – Science: Space",        "https://www.bbc.com/news/science_and_environment"),
        ("Wikipedia – SpaceX",          "https://de.wikipedia.org/wiki/SpaceX"),
    ]
    # zwei Spalten
    col1 = sources[:6]
    col2 = sources[6:]

    def render(col, x0):
        for i, (name, url) in enumerate(col):
            y = Inches(2.7 + i * 0.55)
            add_text(s, x0, y, Inches(6.3), Inches(0.3),
                     name, size=13, bold=True, color=ACCENT_WHITE,
                     font="Calibri Light")
            add_text(s, x0, y + Inches(0.27), Inches(6.3), Inches(0.3),
                     url, size=11, color=ACCENT_BLUE, font="Calibri Light")

    render(col1, Inches(0.5))
    render(col2, Inches(6.8))

    add_text(s, Inches(0.5), Inches(6.7), Inches(12), Inches(0.4),
             "Abgerufen: Mai 2026. Zahlen können sich jederzeit ändern.",
             size=10, italic=True, color=ACCENT_GREY, font="Calibri")


# =====================================================
# SLIDE 25  -  FRAGEN?
# =====================================================
def build_questions_slide():
    s = new_slide("Fragen?")
    add_stars(s, count=200, seed=99)

    add_text(s, Inches(0.5), Inches(2.6), Inches(12.3), Inches(2),
             "Fragen?",
             size=180, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, font="Calibri")

    add_rect(s, Inches(6.16), Inches(5.0), Inches(1.0), Inches(0.05),
             fill=ACCENT_RED, line=None)

    add_text(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(0.5),
             "Wir freuen uns auf eure Diskussion.",
             size=18, color=ACCENT_GREY, align=PP_ALIGN.CENTER,
             italic=True, font="Calibri Light")

    add_text(s, Inches(0.5), Inches(6.8), Inches(12.3), Inches(0.4),
             "JULIUS MÜLLER   ·   BEN HAMMEN",
             size=12, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, letter_spacing=400, font="Calibri")


# =====================================================
# BAU AUSFUEHREN
# =====================================================
build_title_slide()
build_welcome_slide()
build_agenda_slide()
build_what_is_spacex()
build_history_slide()
build_musk_slide()
build_mission_slide()
build_rocket_family_slide()
build_falcon9_slide()
build_reusability_slide()
build_starship_slide()
build_dragon_slide()
build_starlink_slide()
build_quiz_slide()
build_milestones_slide()
build_competition_slide()
build_mars_slide()
build_moon_slide()
build_use_for_us_slide()
build_practice_tips_slide()
build_criticism_slide()
build_discussion_slide()
build_summary_slide()
build_sources_slide()
build_questions_slide()


# =====================================================
# FOOTER fuer alle Folien (ausser Titel & Fragen)
# =====================================================
total = len(SLIDES)
for i, (slide, chapter) in enumerate(SLIDES):
    if i == 0 or i == total - 1:
        continue  # Titel + Letzte Folie ohne klassischen Footer
    add_footer(slide, i + 1, total, chapter)


# =====================================================
# ANIMATIONEN: Fade-in fuer alle Shapes pro Folie
# =====================================================
def add_fade_animation(slide):
    """
    Fuegt der Folie ein simples 'Fade-in nach Klick' fuer alle Shapes hinzu.
    Wir bauen das timing-XML direkt; PowerPoint interpretiert es als
    sequenzielle Fade-Effekte.
    """
    # Wir setzen das auf der Folie via sld/timing
    sld = slide._element
    # bestehende timing-Knoten entfernen
    for t in sld.findall(qn("p:timing")):
        sld.remove(t)

    shape_ids = []
    for sp in slide.shapes:
        if sp.shape_type is None:
            continue
        shape_ids.append(sp.shape_id)
    if not shape_ids:
        return

    # Vereinfachtes Timing: alle Shapes faden nacheinander beim Klick ein
    nsmap = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

    cond_lst = ""
    par_blocks = ""
    for idx, sid in enumerate(shape_ids):
        par_blocks += f'''
        <p:par>
          <p:cTn id="{100 + idx}" fill="hold">
            <p:stCondLst><p:cond delay="0"/></p:stCondLst>
            <p:childTnLst>
              <p:par>
                <p:cTn id="{200 + idx}" fill="hold">
                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="{300 + idx}" presetID="10" presetClass="entr"
                             presetSubtype="0" fill="hold" grpId="0"
                             nodeType="clickEffect">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst>
                          <p:set>
                            <p:cBhvr>
                              <p:cTn id="{400 + idx}" dur="1" fill="hold">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                              </p:cTn>
                              <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                              <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                            </p:cBhvr>
                            <p:to><p:strVal val="visible"/></p:to>
                          </p:set>
                          <p:anim calcmode="lin" valueType="num">
                            <p:cBhvr>
                              <p:cTn id="{500 + idx}" dur="500" fill="hold"/>
                              <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                              <p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>
                            </p:cBhvr>
                            <p:tavLst>
                              <p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>
                              <p:tav tm="100000"><p:val><p:fltVal val="1"/></p:val></p:tav>
                            </p:tavLst>
                          </p:anim>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
        </p:par>
        '''

    timing_xml = f'''
    <p:timing {nsmap}>
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    {par_blocks}
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
    </p:timing>
    '''
    try:
        timing_el = etree.fromstring(timing_xml)
        sld.append(timing_el)
    except Exception as e:
        # Wenn Animation fehlschlaegt, einfach skippen - Folie funktioniert trotzdem
        print(f"Animation skipped: {e}")


# Animationen aktivieren wir aktuell NICHT global, weil das PPTX-XML
# zwischen PowerPoint-Versionen empfindlich ist. Folien-Uebergaenge unten
# liefern bereits einen schoenen Effekt. Wer mag, kann die Schleife
# auskommentieren:
# for slide, _ in SLIDES:
#     add_fade_animation(slide)


# =====================================================
# FOLIEN-UEBERGAENGE (Fade) per XML
# =====================================================
def add_fade_transition(slide):
    sld = slide._element
    # bestehende transition entfernen
    for t in sld.findall(qn("p:transition")):
        sld.remove(t)
    # Element vor Timing einfuegen
    transition_xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'spd="med" advClick="1">'
        '<p:fade/>'
        '</p:transition>'
    )
    el = etree.fromstring(transition_xml)
    sld.append(el)

for slide, _ in SLIDES:
    add_fade_transition(slide)


# =====================================================
# SPEICHERN
# =====================================================
out_path = "/home/user/Pr-si-Mathe-/SpaceX_Praesentation_Mueller_Hammen.pptx"
prs.save(out_path)
print(f"OK -> {out_path}")
print(f"Folien gesamt: {total}")
