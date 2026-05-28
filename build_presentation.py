"""
SpaceX Praesentation v2 - Julius Mueller & Ben Hammen
- minimalistisches Dark-Space-Design
- viele grosse Bild-Platzhalter (leicht ersetzbar)
- 3 Histogramme (Starts/Jahr, Kosten/kg, Starlink)
- viele Zeitangaben
- Stichpunkte statt Fliesstext
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import random

# ===== Farben =====
BG_DARK       = RGBColor(0x05, 0x08, 0x12)
ACCENT_WHITE  = RGBColor(0xF5, 0xF7, 0xFA)
ACCENT_GREY   = RGBColor(0x9A, 0xA3, 0xB2)
ACCENT_DIM    = RGBColor(0x60, 0x68, 0x78)
ACCENT_RED    = RGBColor(0xE6, 0x1E, 0x2B)
ACCENT_BLUE   = RGBColor(0x4A, 0x9E, 0xFF)
ACCENT_ORANGE = RGBColor(0xFF, 0x7A, 0x2A)
ACCENT_GREEN  = RGBColor(0x3F, 0xD0, 0x8A)
LINE_GREY     = RGBColor(0x2B, 0x33, 0x44)
PLACEHOLDER_FILL = RGBColor(0x0B, 0x10, 0x1C)
PLACEHOLDER_BORDER = RGBColor(0xE6, 0x1E, 0x2B)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

SLIDES = []
IMG_COUNTER = [0]


# =====================================================
# HELPERS
# =====================================================
def add_bg(slide, color=BG_DARK):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.line.fill.background()
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.shadow.inherit = False
    spTree = bg._element.getparent()
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)
    return bg


def add_stars(slide, count=60, seed=None):
    if seed is not None:
        random.seed(seed)
    for _ in range(count):
        x = Emu(random.randint(0, int(SLIDE_W)))
        y = Emu(random.randint(0, int(SLIDE_H)))
        size_pt = random.choice([1, 1, 1, 2, 2, 3])
        size = Emu(int(size_pt * 9525))
        star = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
        star.line.fill.background()
        star.fill.solid()
        v = random.choice([0x55, 0x88, 0xAA, 0xCC, 0xEE, 0xFF])
        star.fill.fore_color.rgb = RGBColor(v, v, v)


def add_text(slide, left, top, width, height, text,
             size=24, bold=False, color=ACCENT_WHITE,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font="Calibri Light", italic=False, letter_spacing=None,
             line_spacing=None):
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
        if line_spacing is not None:
            p.line_spacing = line_spacing
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
             fill=None, line=None, line_weight=1.0,
             shape=MSO_SHAPE.RECTANGLE, line_dash=None):
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
        if line_dash is not None:
            ln = s.line._get_or_add_ln()
            prstDash = etree.SubElement(ln, qn("a:prstDash"))
            prstDash.set("val", line_dash)
    s.shadow.inherit = False
    return s


def add_image_placeholder(slide, left, top, width, height, description,
                          ratio_hint=None):
    """Grosser, gut sichtbarer Bild-Platzhalter mit Nummer und Beschreibung."""
    IMG_COUNTER[0] += 1
    n = IMG_COUNTER[0]

    add_rect(slide, left, top, width, height,
             fill=PLACEHOLDER_FILL, line=PLACEHOLDER_BORDER,
             line_weight=1.5, line_dash="dash")
    add_line(slide, left, top, left + width, top + height,
             color=RGBColor(0x1A, 0x22, 0x36), weight=0.5)
    add_line(slide, left, top + height, left + width, top,
             color=RGBColor(0x1A, 0x22, 0x36), weight=0.5)

    # Kamera-Symbol
    cam_w = Inches(0.7)
    cam_h = Inches(0.5)
    cam_x = left + width/2 - cam_w/2
    cam_y = top + height/2 - Inches(0.9)
    add_rect(slide, cam_x, cam_y, cam_w, cam_h,
             fill=None, line=PLACEHOLDER_BORDER, line_weight=1.5,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, cam_x + cam_w/2 - Inches(0.13), cam_y + Inches(0.13),
             Inches(0.26), Inches(0.24),
             fill=None, line=PLACEHOLDER_BORDER, line_weight=1.2,
             shape=MSO_SHAPE.OVAL)

    add_text(slide, left, top + height/2 - Inches(0.25),
             width, Inches(0.6),
             f"BILD {n:02d}", size=22, bold=True,
             color=ACCENT_WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Calibri",
             letter_spacing=400)

    add_text(slide, left + Inches(0.3), top + height/2 + Inches(0.15),
             width - Inches(0.6), Inches(0.8),
             description, size=11, color=ACCENT_GREY,
             align=PP_ALIGN.CENTER, italic=True, font="Calibri Light")

    if ratio_hint:
        add_text(slide, left + width - Inches(2.2), top + height - Inches(0.35),
                 Inches(2.0), Inches(0.3),
                 ratio_hint, size=9, color=ACCENT_DIM,
                 align=PP_ALIGN.RIGHT, font="Calibri")

    tag_w = Inches(1.4)
    tag_h = Inches(0.35)
    add_rect(slide, left, top, tag_w, tag_h, fill=PLACEHOLDER_BORDER, line=None)
    add_text(slide, left, top, tag_w, tag_h,
             f"BILD {n:02d}", size=10, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             font="Calibri", letter_spacing=300)


def add_footer(slide, page_num, total, chapter="", speaker=""):
    add_line(slide, Inches(0.5), Inches(7.05), Inches(12.83), Inches(7.05),
             color=LINE_GREY, weight=0.75)
    if chapter:
        add_text(slide, Inches(0.5), Inches(7.12), Inches(6), Inches(0.3),
                 chapter.upper(), size=9, color=ACCENT_GREY,
                 letter_spacing=300, font="Calibri")
    if speaker:
        speaker_color = ACCENT_RED if "BEN" in speaker.upper() else (
                        ACCENT_BLUE if "JULIUS" in speaker.upper() else ACCENT_ORANGE)
        add_text(slide, Inches(5), Inches(7.12), Inches(3.33), Inches(0.3),
                 f"SPRECHER: {speaker}", size=9, bold=True,
                 color=speaker_color,
                 align=PP_ALIGN.CENTER, letter_spacing=300, font="Calibri")
    add_text(slide, Inches(11.33), Inches(7.12), Inches(1.5), Inches(0.3),
             f"{page_num:02d} / {total:02d}", size=9, color=ACCENT_GREY,
             align=PP_ALIGN.RIGHT, letter_spacing=300, font="Calibri")


def add_top_marker(slide, kapitel_nr, kapitel_name):
    add_rect(slide, Inches(0.5), Inches(0.5), Inches(0.05), Inches(0.35),
             fill=ACCENT_RED, line=None)
    add_text(slide, Inches(0.7), Inches(0.5), Inches(10), Inches(0.4),
             f"{kapitel_nr:02d}   ·   {kapitel_name.upper()}",
             size=10, color=ACCENT_GREY, letter_spacing=400, font="Calibri",
             anchor=MSO_ANCHOR.MIDDLE)


def add_date_chip(slide, left, top, date_text):
    w = Inches(2.0)
    h = Inches(0.4)
    add_rect(slide, left, top, w, h,
             fill=None, line=ACCENT_GREY, line_weight=0.5,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(slide, left, top, w, h,
             date_text, size=10, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             font="Calibri", letter_spacing=200)


def add_title(slide, top, text, size=44, color=ACCENT_WHITE,
              align=PP_ALIGN.LEFT, width=Inches(12.33)):
    add_text(slide, Inches(0.5), top, width, Inches(1.5),
             text, size=size, bold=True, color=color, align=align,
             font="Calibri Light", line_spacing=1.05)


def add_kicker(slide, top, text, color=ACCENT_RED):
    add_text(slide, Inches(0.5), top, Inches(12), Inches(0.4),
             text.upper(), size=11, bold=True, color=color,
             letter_spacing=400, font="Calibri")


def add_underline(slide, top, length=Inches(0.6), color=ACCENT_RED):
    add_rect(slide, Inches(0.5), top, length, Inches(0.05),
             fill=color, line=None)


def add_bullet(slide, x, y, w, text, size=15):
    bullet_size = Inches(0.12)
    add_rect(slide, x, y + Inches(0.12), bullet_size, bullet_size,
             fill=ACCENT_RED, line=None, shape=MSO_SHAPE.OVAL)
    add_text(slide, x + Inches(0.3), y, w, Inches(0.5),
             text, size=size, color=ACCENT_WHITE, font="Calibri Light")


def add_bar_chart(slide, x, y, width, height, data, max_val=None,
                  color=ACCENT_RED, label_color=ACCENT_WHITE,
                  value_color=ACCENT_GREY, highlight_last=False):
    if max_val is None:
        max_val = max(v for _, v in data)
    n = len(data)
    add_line(slide, x, y + height, x + width, y + height,
             color=LINE_GREY, weight=0.75)
    bar_area_w = width
    gap = bar_area_w / (n * 6)
    bar_w = (bar_area_w - gap * (n + 1)) / n
    for i, (label, val) in enumerate(data):
        bar_h = Emu(int(height * (val / max_val) * 0.85))
        bx = x + gap + (bar_w + gap) * i
        by = y + height - bar_h
        bar_color = color
        if highlight_last and i == n - 1:
            bar_color = ACCENT_WHITE
        add_rect(slide, bx, by, bar_w, bar_h, fill=bar_color, line=None)
        add_text(slide, bx - Inches(0.2), by - Inches(0.4),
                 bar_w + Inches(0.4), Inches(0.35),
                 str(val), size=9, bold=True, color=value_color,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(slide, bx - Inches(0.2), y + height + Inches(0.05),
                 bar_w + Inches(0.4), Inches(0.3),
                 label, size=9, color=label_color,
                 align=PP_ALIGN.CENTER, font="Calibri")


def new_slide(chapter="", speaker=""):
    s = prs.slides.add_slide(BLANK)
    add_bg(s)
    SLIDES.append((s, chapter, speaker))
    return s


# =====================================================
# 1. TITEL (BEIDE)
# =====================================================
def s01_title():
    s = new_slide("Titel", "BEIDE")
    add_stars(s, count=140, seed=1)
    add_rect(s, Inches(0.5), Inches(2.8), Inches(0.08), Inches(1.6),
             fill=ACCENT_RED, line=None)
    add_text(s, Inches(0.8), Inches(2.85), Inches(8), Inches(0.4),
             "REFERAT  ·  RAUMFAHRT  ·  ZUKUNFT",
             size=12, bold=True, color=ACCENT_RED,
             letter_spacing=500, font="Calibri")
    add_text(s, Inches(0.8), Inches(3.3), Inches(12), Inches(1.5),
             "SpaceX", size=130, bold=True, color=ACCENT_WHITE,
             font="Calibri")
    add_text(s, Inches(0.8), Inches(5.0), Inches(12), Inches(0.6),
             "Eine Firma. Ein Ziel: Mars.",
             size=24, color=ACCENT_GREY, font="Calibri Light")
    add_line(s, Inches(0.8), Inches(6.4), Inches(4.5), Inches(6.4),
             color=ACCENT_RED, weight=1.5)
    add_text(s, Inches(0.8), Inches(6.5), Inches(6), Inches(0.4),
             "JULIUS MÜLLER   ·   BEN HAMMEN",
             size=14, bold=True, color=ACCENT_WHITE,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(0.8), Inches(6.85), Inches(6), Inches(0.35),
             "Mai 2026", size=11, color=ACCENT_GREY, font="Calibri")
    add_image_placeholder(s, Inches(9.3), Inches(0.5), Inches(3.5), Inches(2.0),
                          "SpaceX-Logo (transparent, weiß)",
                          ratio_hint="ca. 16:9")


# =====================================================
# 2. BEGRUESSUNG (BEN)
# =====================================================
def s02_welcome():
    s = new_slide("Begrüßung", "BEN")
    add_stars(s, count=30, seed=2)
    add_kicker(s, Inches(0.9), "Willkommen")
    add_underline(s, Inches(1.25))
    add_text(s, Inches(0.5), Inches(1.8), Inches(12), Inches(3),
             "Heute fliegen wir\nmit SpaceX.",
             size=80, bold=True, color=ACCENT_WHITE,
             font="Calibri Light", line_spacing=1.0)
    keys = ["RAKETEN.", "VISIONEN.", "DISKUSSION."]
    for i, k in enumerate(keys):
        x = Inches(0.5 + i * 4.3)
        add_rect(s, x, Inches(5.7), Inches(0.4), Inches(0.05),
                 fill=ACCENT_RED, line=None)
        add_text(s, x, Inches(5.85), Inches(4), Inches(0.6),
                 k, size=24, bold=True, color=ACCENT_WHITE,
                 font="Calibri Light", letter_spacing=200)


# =====================================================
# 3. AGENDA (JULIUS)
# =====================================================
def s03_agenda():
    s = new_slide("Agenda", "JULIUS")
    add_kicker(s, Inches(0.9), "Themenübersicht")
    add_underline(s, Inches(1.25))
    add_title(s, Inches(1.6), "Unsere Reise", size=46)

    items_left = [
        ("01", "Was ist SpaceX?"),
        ("02", "Geschichte"),
        ("03", "Elon Musk"),
        ("04", "Vision & Mission"),
        ("05", "Die Raketen"),
        ("06", "Wiederverwendbarkeit"),
    ]
    items_right = [
        ("07", "Dragon"),
        ("08", "Starlink"),
        ("09", "SpaceX in Zahlen"),
        ("10", "Mond & Mars"),
        ("11", "Was bringt's uns?"),
        ("12", "Kritik & Diskussion"),
    ]
    def render(col_items, x0):
        for i, (n, t) in enumerate(col_items):
            y = Inches(2.8 + i * 0.6)
            add_text(s, x0, y, Inches(0.6), Inches(0.4),
                     n, size=14, bold=True, color=ACCENT_RED,
                     font="Calibri", letter_spacing=200)
            add_text(s, x0 + Inches(0.8), y, Inches(5), Inches(0.4),
                     t, size=20, color=ACCENT_WHITE, font="Calibri Light")
    render(items_left, Inches(0.5))
    render(items_right, Inches(7.0))
    add_text(s, Inches(0.5), Inches(6.5), Inches(8), Inches(0.4),
             "≈ 30 Minuten   ·   Fragen am Ende",
             size=12, color=ACCENT_GREY, italic=True, font="Calibri")


# =====================================================
# 4. WAS IST SPACEX (BEN)
# =====================================================
def s04_what():
    s = new_slide("01 · Was ist SpaceX", "BEN")
    add_top_marker(s, 1, "Was ist SpaceX?")
    add_date_chip(s, Inches(11.0), Inches(0.5), "Stand: 2025")
    add_title(s, Inches(1.2), "Vier Worte.", size=54)
    add_underline(s, Inches(2.4))

    keys = ["PRIVAT", "WIEDER-\nVERWENDBAR", "SCHNELL", "AMBITIONIERT"]
    box_w = Inches(2.95)
    gap   = Inches(0.2)
    start = Inches(0.5)
    for i, k in enumerate(keys):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(3.2), box_w, Inches(2.6),
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_text(s, x + Inches(0.25), Inches(3.4), Inches(1), Inches(0.4),
                 f"0{i+1}", size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, x, Inches(3.9), box_w, Inches(1.7),
                 k, size=30, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font="Calibri Light")
    add_text(s, Inches(0.5), Inches(6.2), Inches(12), Inches(0.4),
             "Space Exploration Technologies Corp.   ·   Sitz: Hawthorne, Kalifornien",
             size=14, italic=True, color=ACCENT_BLUE, font="Calibri Light")


# =====================================================
# 5. GESCHICHTE / ZEITSTRAHL (JULIUS)
# =====================================================
def s05_history():
    s = new_slide("02 · Geschichte", "JULIUS")
    add_top_marker(s, 2, "Geschichte")
    add_date_chip(s, Inches(11.0), Inches(0.5), "2002 – heute")
    add_title(s, Inches(1.2), "23 Jahre. 6 Sprünge.", size=44)
    add_underline(s, Inches(2.25))

    timeline_y = Inches(5.0)
    add_line(s, Inches(0.7), timeline_y, Inches(12.6), timeline_y,
             color=ACCENT_RED, weight=2.0)
    events = [
        ("2002", "Gründung"),
        ("2008", "Falcon 1\nim Orbit"),
        ("2012", "Dragon\nan ISS"),
        ("2015", "1. Falcon-\nLandung"),
        ("2020", "Crew zur\nISS"),
        ("2024", "Starship-\nFangarm"),
    ]
    n = len(events)
    span = Inches(11.9)
    start_x = Inches(0.7)
    step = span / (n - 1)
    for i, (year, label) in enumerate(events):
        x = start_x + step * i
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                  x - Inches(0.15), timeline_y - Inches(0.15),
                                  Inches(0.3), Inches(0.3))
        dot.fill.solid(); dot.fill.fore_color.rgb = ACCENT_RED
        dot.line.color.rgb = BG_DARK
        dot.line.width = Pt(2)
        add_text(s, x - Inches(1), timeline_y - Inches(1.4),
                 Inches(2), Inches(0.5),
                 year, size=22, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(s, x - Inches(1.2), timeline_y + Inches(0.3),
                 Inches(2.4), Inches(1.3),
                 label, size=12, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri Light")
    add_text(s, Inches(0.5), Inches(2.8), Inches(12), Inches(0.4),
             "Startkapital: 100 Mio. $ aus Musks PayPal-Verkauf (2002).",
             size=14, color=ACCENT_BLUE, italic=True, font="Calibri Light")


# =====================================================
# 6. ELON MUSK (BEN)
# =====================================================
def s06_musk():
    s = new_slide("03 · Elon Musk", "BEN")
    add_top_marker(s, 3, "Elon Musk")
    add_date_chip(s, Inches(11.0), Inches(0.5), "geb. 1971")
    add_title(s, Inches(1.2), "Der Kopf dahinter.", size=44)
    add_underline(s, Inches(2.25))
    add_image_placeholder(s, Inches(0.5), Inches(3.0), Inches(5.5), Inches(3.8),
                          "Portrait Elon Musk\n(z. B. Wikipedia – CC-BY)",
                          ratio_hint="Hochformat / quadratisch")
    add_text(s, Inches(6.5), Inches(3.0), Inches(6), Inches(0.4),
             "FAKTEN", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    bullets = [
        "Geboren 1971 in Pretoria (Südafrika)",
        "Studium: Physik & BWL (USA)",
        "PayPal-Mitgründer → Verkauf 2002",
        "CEO: SpaceX · Tesla · xAI · X",
        "Vermögen: > 300 Mrd. $ (Forbes 2025)",
    ]
    for i, b in enumerate(bullets):
        add_bullet(s, Inches(6.5), Inches(3.6 + i * 0.55), Inches(6.5), b, size=16)


# =====================================================
# 7. MISSION (JULIUS)
# =====================================================
def s07_mission():
    s = new_slide("04 · Vision", "JULIUS")
    add_stars(s, count=60, seed=7)
    add_top_marker(s, 4, "Mission & Vision")
    add_title(s, Inches(1.2), "Warum?", size=72)
    add_underline(s, Inches(2.6))
    add_text(s, Inches(0.5), Inches(3.3), Inches(12.3), Inches(1.8),
             "„Die Menschheit\nmultiplanetar machen.“",
             size=46, italic=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, font="Calibri Light",
             line_spacing=1.1)
    add_text(s, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.5),
             "— Elon Musk, SpaceX-Mission Statement",
             size=14, color=ACCENT_GREY, italic=True,
             align=PP_ALIGN.CENTER, font="Calibri Light")
    pillars = [("MISSION", "billiger fliegen"),
               ("VISION", "Mars besiedeln"),
               ("METHODE", "wiederverwenden")]
    for i, (k, v) in enumerate(pillars):
        x = Inches(0.5 + i * 4.3)
        add_rect(s, x, Inches(6.05), Inches(0.5), Inches(0.04),
                 fill=ACCENT_RED, line=None)
        add_text(s, x, Inches(6.15), Inches(4), Inches(0.35),
                 k, size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=400, font="Calibri")
        add_text(s, x, Inches(6.5), Inches(4), Inches(0.5),
                 v, size=18, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# 8. RAKETEN-FAMILIE (BEN)
# =====================================================
def s08_rockets():
    s = new_slide("05 · Raketen", "BEN")
    add_top_marker(s, 5, "Die Raketen-Familie")
    add_date_chip(s, Inches(11.0), Inches(0.5), "2006 – heute")
    add_title(s, Inches(1.2), "Vier Größen.", size=44)
    add_underline(s, Inches(2.25))

    rockets = [
        ("FALCON 1",     21,   "21 m",  "0,67 t",  "ab 2006", ACCENT_GREY),
        ("FALCON 9",     70,   "70 m",  "22,8 t",  "ab 2010", ACCENT_WHITE),
        ("FALCON HEAVY", 70,   "70 m",  "63,8 t",  "ab 2018", ACCENT_BLUE),
        ("STARSHIP",    121,  "121 m", ">100 t",   "ab 2023", ACCENT_RED),
    ]
    max_h = 121
    base_y = Inches(6.2)
    max_pic_h = Inches(3.0)
    col_w = Inches(2.8)
    gap = Inches(0.3)
    start = Inches(0.7)
    for i, (name, h, h_lbl, payload, year, color) in enumerate(rockets):
        x = start + (col_w + gap) * i
        pic_h = Emu(int(max_pic_h * (h / max_h)))
        body_w = Inches(0.55)
        add_rect(s, x + col_w/2 - body_w/2, base_y - pic_h,
                 body_w, pic_h, fill=color, line=None)
        tip = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                  x + col_w/2 - body_w/2,
                                  base_y - pic_h - Inches(0.45),
                                  body_w, Inches(0.5))
        tip.fill.solid(); tip.fill.fore_color.rgb = color
        tip.line.fill.background()
        add_line(s, x, base_y, x + col_w, base_y,
                 color=LINE_GREY, weight=0.75)
        add_text(s, x, base_y + Inches(0.1), col_w, Inches(0.4),
                 name, size=13, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, font="Calibri", letter_spacing=200)
        add_text(s, x, base_y + Inches(0.5), col_w, Inches(0.3),
                 h_lbl, size=11, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(s, x, base_y + Inches(0.8), col_w, Inches(0.3),
                 payload, size=11, color=ACCENT_GREY,
                 align=PP_ALIGN.CENTER, font="Calibri")
        add_text(s, x, base_y + Inches(1.1), col_w, Inches(0.3),
                 year, size=9, color=ACCENT_RED,
                 align=PP_ALIGN.CENTER, font="Calibri", letter_spacing=200)


# =====================================================
# 9. FALCON 9 DETAIL (JULIUS)
# =====================================================
def s09_falcon9():
    s = new_slide("05 · Falcon 9", "JULIUS")
    add_top_marker(s, 5, "Falcon 9 – das Arbeitspferd")
    add_date_chip(s, Inches(11.0), Inches(0.5), "seit 2010")
    add_title(s, Inches(1.2), "Falcon 9.", size=64)
    add_text(s, Inches(0.5), Inches(2.65), Inches(12), Inches(0.5),
             "Die meistgenutzte Rakete der Welt.",
             size=18, color=ACCENT_GREY, italic=True, font="Calibri Light")
    add_image_placeholder(s, Inches(0.5), Inches(3.4), Inches(6.5), Inches(3.3),
                          "Falcon 9 Start bei Nacht\n(SpaceX Flickr, CC-BY-NC)",
                          ratio_hint="Querformat 16:9 oder 4:3")
    add_text(s, Inches(7.3), Inches(3.4), Inches(5), Inches(0.4),
             "DATEN", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    bullets = [
        "70 m hoch",
        "9 × Merlin-Triebwerke",
        "Nutzlast: 22,8 t (LEO)",
        "Booster: bis zu 20 Flüge",
        "Preis: ~ 67 Mio. $ / Start",
    ]
    for i, b in enumerate(bullets):
        add_bullet(s, Inches(7.3), Inches(4.0 + i * 0.55), Inches(6), b, size=18)


# =====================================================
# 10. WIEDERVERWENDBARKEIT + HISTOGRAMM KOSTEN (BEN)
# =====================================================
def s10_reuse():
    s = new_slide("06 · Wiederverwendbarkeit", "BEN")
    add_top_marker(s, 6, "Wiederverwendbarkeit")
    add_title(s, Inches(1.2), "Die Revolution.", size=44)
    add_underline(s, Inches(2.25))
    add_text(s, Inches(0.5), Inches(2.9), Inches(12), Inches(0.4),
             "KOSTEN PRO KG IN DEN ORBIT (USD)", size=11, bold=True,
             color=ACCENT_GREY, letter_spacing=300, font="Calibri")
    data = [
        ("Space Shuttle\n1981–2011", 54500),
        ("Ariane 5\n1996–2023", 10000),
        ("Falcon 9\nseit 2010", 2720),
        ("Falcon Heavy\nseit 2018", 1400),
        ("Starship\nZiel", 200),
    ]
    add_bar_chart(s, Inches(0.7), Inches(3.4), Inches(8.0), Inches(2.8),
                  data, max_val=55000, color=ACCENT_RED, highlight_last=True)
    add_rect(s, Inches(9.2), Inches(3.4), Inches(3.6), Inches(2.8),
             fill=RGBColor(0x18, 0x0E, 0x12), line=ACCENT_RED)
    add_text(s, Inches(9.4), Inches(3.55), Inches(3.4), Inches(0.4),
             "KERNAUSSAGE", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(9.4), Inches(4.0), Inches(3.3), Inches(2),
             "−95 %\nKosten",
             size=44, bold=True, color=ACCENT_WHITE,
             font="Calibri Light", line_spacing=1.0)
    add_text(s, Inches(9.4), Inches(5.8), Inches(3.3), Inches(0.4),
             "vs. Space Shuttle",
             size=12, color=ACCENT_GREY, italic=True, font="Calibri Light")
    add_text(s, Inches(0.5), Inches(6.4), Inches(12), Inches(0.4),
             "Quelle: NASA, CSIS Aerospace Reports 2024.",
             size=10, italic=True, color=ACCENT_DIM, font="Calibri")


# =====================================================
# 11. STARSHIP (JULIUS)
# =====================================================
def s11_starship():
    s = new_slide("05 · Starship", "JULIUS")
    add_stars(s, count=60, seed=11)
    add_top_marker(s, 5, "Starship – die Zukunft")
    add_date_chip(s, Inches(11.0), Inches(0.5), "Test-Flüge seit 2023")
    add_title(s, Inches(1.2), "Starship.", size=80)
    add_text(s, Inches(0.5), Inches(2.7), Inches(12), Inches(0.5),
             "Die größte Rakete aller Zeiten.",
             size=18, color=ACCENT_GREY, italic=True, font="Calibri Light")
    add_image_placeholder(s, Inches(7.3), Inches(3.4), Inches(5.5), Inches(3.4),
                          "Starship Vollstack auf Starbase\n(SpaceX Flickr)",
                          ratio_hint="Hochformat ideal")
    add_text(s, Inches(0.5), Inches(3.4), Inches(5), Inches(0.4),
             "ZAHLEN", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    facts = [
        ("121 m",  "höher als Freiheitsstatue"),
        (">100 t", "Nutzlast in den Orbit"),
        ("33",     "Raptor-Triebwerke"),
        ("0 $",    "Wegwerf-Hardware (Ziel)"),
    ]
    for i, (n, t) in enumerate(facts):
        y = Inches(3.95 + i * 0.65)
        add_text(s, Inches(0.5), y, Inches(2.0), Inches(0.5),
                 n, size=26, bold=True, color=ACCENT_RED, font="Calibri")
        add_text(s, Inches(2.7), y + Inches(0.15), Inches(4.3), Inches(0.5),
                 t, size=14, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# 12. DRAGON (BEN)
# =====================================================
def s12_dragon():
    s = new_slide("07 · Dragon", "BEN")
    add_top_marker(s, 7, "Dragon & Crew Dragon")
    add_date_chip(s, Inches(11.0), Inches(0.5), "seit 2012 / 2020")
    add_title(s, Inches(1.2), "Dragon.", size=64)
    add_text(s, Inches(0.5), Inches(2.7), Inches(12), Inches(0.5),
             "Fracht und Menschen ins All.",
             size=18, color=ACCENT_GREY, italic=True, font="Calibri Light")

    add_rect(s, Inches(0.5), Inches(3.4), Inches(6.0), Inches(3.5),
             fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
    add_image_placeholder(s, Inches(0.7), Inches(3.6), Inches(5.6), Inches(2.0),
                          "Cargo Dragon (Andocken ISS)",
                          ratio_hint="16:9")
    add_text(s, Inches(0.7), Inches(5.7), Inches(2), Inches(0.35),
             "CARGO DRAGON", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(0.7), Inches(6.05), Inches(5.6), Inches(0.4),
             "Versorgt ISS · seit 2012", size=15, color=ACCENT_WHITE,
             font="Calibri Light")
    add_text(s, Inches(0.7), Inches(6.45), Inches(5.6), Inches(0.4),
             "bis 6.000 kg Fracht", size=12, color=ACCENT_GREY,
             font="Calibri Light")

    add_rect(s, Inches(6.83), Inches(3.4), Inches(6.0), Inches(3.5),
             fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
    add_image_placeholder(s, Inches(7.03), Inches(3.6), Inches(5.6), Inches(2.0),
                          "Crew Dragon mit Astronauten",
                          ratio_hint="16:9")
    add_text(s, Inches(7.03), Inches(5.7), Inches(2), Inches(0.35),
             "CREW DRAGON", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(7.03), Inches(6.05), Inches(5.6), Inches(0.4),
             "Astronauten · seit 2020", size=15, color=ACCENT_WHITE,
             font="Calibri Light")
    add_text(s, Inches(7.03), Inches(6.45), Inches(5.6), Inches(0.4),
             "4 Sitzplätze · Inspiration4 (zivil)", size=12,
             color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# 13. STARLINK + HISTOGRAMM (JULIUS)
# =====================================================
def s13_starlink():
    s = new_slide("08 · Starlink", "JULIUS")
    add_stars(s, count=80, seed=13)
    add_top_marker(s, 8, "Starlink – Internet aus dem All")
    add_date_chip(s, Inches(11.0), Inches(0.5), "seit 2019")
    add_title(s, Inches(1.2), "Starlink.", size=64)
    add_text(s, Inches(0.5), Inches(2.65), Inches(12), Inches(0.5),
             "Wachstum, das es noch nie gab.",
             size=18, color=ACCENT_GREY, italic=True, font="Calibri Light")
    add_text(s, Inches(0.5), Inches(3.3), Inches(12), Inches(0.4),
             "AKTIVE STARLINK-SATELLITEN IM ORBIT", size=11, bold=True,
             color=ACCENT_GREY, letter_spacing=300, font="Calibri")
    data = [
        ("2019", 60), ("2020", 955), ("2021", 1944),
        ("2022", 3271), ("2023", 5289), ("2024", 6800), ("2025", 7500),
    ]
    add_bar_chart(s, Inches(0.5), Inches(3.8), Inches(8.5), Inches(2.6),
                  data, max_val=8000, color=ACCENT_BLUE, highlight_last=True)
    add_text(s, Inches(9.4), Inches(3.3), Inches(3.5), Inches(0.4),
             "WAS DAS BEDEUTET", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    bullets = [
        "Empfang in 100+ Ländern",
        "~ 50 ms Latenz",
        "~ 550 km Flughöhe",
        "wichtig im Ukraine-Krieg",
    ]
    for i, b in enumerate(bullets):
        add_bullet(s, Inches(9.4), Inches(3.85 + i * 0.55), Inches(3.5), b, size=13)
    add_text(s, Inches(0.5), Inches(6.55), Inches(12), Inches(0.4),
             "Quelle: Jonathan McDowell / planet4589.org · SpaceX.",
             size=10, italic=True, color=ACCENT_DIM, font="Calibri")


# =====================================================
# 14. QUIZ (BEN)
# =====================================================
def s14_quiz():
    s = new_slide("Interaktiv · Quiz", "BEN")
    add_kicker(s, Inches(0.9), "Quiz · für euch!")
    add_underline(s, Inches(1.25))
    add_title(s, Inches(1.6), "Wie hoch ist Starship?", size=46)
    options = [
        ("A", "72 m   – wie Falcon 9"),
        ("B", "98 m   – wie eine Saturn V"),
        ("C", "121 m – höher als Freiheitsstatue"),
        ("D", "180 m – höher als Kölner Dom"),
    ]
    for i, (l, t) in enumerate(options):
        y = Inches(3.8 + i * 0.7)
        add_rect(s, Inches(2.5), y, Inches(0.85), Inches(0.6),
                 fill=ACCENT_RED, line=None)
        add_text(s, Inches(2.5), y, Inches(0.85), Inches(0.6),
                 l, size=20, bold=True, color=ACCENT_WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font="Calibri")
        add_text(s, Inches(3.6), y + Inches(0.1), Inches(8), Inches(0.5),
                 t, size=20, color=ACCENT_WHITE, font="Calibri Light")
    add_text(s, Inches(0.5), Inches(6.7), Inches(12), Inches(0.4),
             "→ Auflösung kommt auf der nächsten Folie.",
             size=12, italic=True, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# 15. MEILENSTEINE + HISTOGRAMM STARTS/JAHR (JULIUS)
# =====================================================
def s15_milestones():
    s = new_slide("09 · SpaceX in Zahlen", "JULIUS")
    add_top_marker(s, 9, "SpaceX in Zahlen")
    add_date_chip(s, Inches(11.0), Inches(0.5), "2010 – 2025")
    add_rect(s, Inches(0.5), Inches(1.2), Inches(2.0), Inches(0.5),
             fill=ACCENT_RED, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(0.5), Inches(1.2), Inches(2.0), Inches(0.5),
             "✓ Lösung: C", size=14, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             font="Calibri", letter_spacing=200)
    add_title(s, Inches(1.85), "Starts pro Jahr.", size=42)
    add_underline(s, Inches(2.85))
    data = [
        ("'10", 2),  ("'11", 0),  ("'12", 2),  ("'13", 3),  ("'14", 6),
        ("'15", 7),  ("'16", 9),  ("'17", 18), ("'18", 21), ("'19", 13),
        ("'20", 26), ("'21", 31), ("'22", 61), ("'23", 96), ("'24", 134),
        ("'25", 170),
    ]
    add_bar_chart(s, Inches(0.5), Inches(3.4), Inches(8.5), Inches(2.8),
                  data, max_val=180, color=ACCENT_WHITE, highlight_last=True)
    add_text(s, Inches(9.4), Inches(3.4), Inches(3.5), Inches(0.4),
             "DAS GROSSE BILD", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    stats = [
        ("450+", "Falcon-Starts"),
        ("400+", "Landungen"),
        ("~87 %", "Welt-Orbit-Anteil 2024"),
    ]
    for i, (n, t) in enumerate(stats):
        y = Inches(3.95 + i * 0.85)
        add_text(s, Inches(9.4), y, Inches(3.5), Inches(0.5),
                 n, size=28, bold=True, color=ACCENT_WHITE, font="Calibri")
        add_text(s, Inches(9.4), y + Inches(0.55), Inches(3.5), Inches(0.4),
                 t, size=11, color=ACCENT_GREY, font="Calibri Light")
    add_text(s, Inches(0.5), Inches(6.55), Inches(12), Inches(0.4),
             "Quelle: SpaceX-Statistik · BryceTech Start Report 2024.",
             size=10, italic=True, color=ACCENT_DIM, font="Calibri")


# =====================================================
# 16. WETTBEWERB (BEN)
# =====================================================
def s16_competition():
    s = new_slide("Wettbewerb", "BEN")
    add_top_marker(s, 9, "Wer macht Konkurrenz?")
    add_title(s, Inches(1.2), "Im Vergleich.", size=44)
    add_underline(s, Inches(2.25))
    cols = ["Akteur", "Heimat", "Stärke", "Schwäche"]
    rows = [
        ["SpaceX",       "USA · privat",   "Wiederverwendung",  "Macht in einer Hand"],
        ["NASA",         "USA · Staat",    "Wissenschaft",      "langsam, teuer"],
        ["Blue Origin",  "USA · privat",   "Investor Bezos",    "wenig Orbit-Erfahrung"],
        ["ESA",          "Europa",         "Forschung",         "kaum wiederverwendbar"],
        ["CNSA",         "China",          "eigene Raumstation","wenig Transparenz"],
    ]
    col_x = [Inches(0.5), Inches(3.7), Inches(6.7), Inches(9.7)]
    col_w = [Inches(3.0), Inches(2.8), Inches(2.8), Inches(3.0)]
    row_y0 = Inches(3.0)
    row_h  = Inches(0.7)
    for j, c in enumerate(cols):
        add_text(s, col_x[j], row_y0, col_w[j], Inches(0.4),
                 c.upper(), size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
    add_line(s, Inches(0.5), row_y0 + Inches(0.45),
             Inches(12.83), row_y0 + Inches(0.45),
             color=ACCENT_RED, weight=1.0)
    for i, row in enumerate(rows):
        y = row_y0 + Inches(0.55) + row_h * i
        for j, cell in enumerate(row):
            color = ACCENT_WHITE if j == 0 else ACCENT_GREY
            size  = 16 if j == 0 else 13
            bold  = j == 0
            add_text(s, col_x[j], y + Inches(0.1), col_w[j], row_h,
                     cell, size=size, bold=bold, color=color,
                     font="Calibri Light")
        add_line(s, Inches(0.5), y + row_h - Inches(0.05),
                 Inches(12.83), y + row_h - Inches(0.05),
                 color=LINE_GREY, weight=0.5)


# =====================================================
# 17. MARS (JULIUS)
# =====================================================
def s17_mars():
    s = new_slide("10 · Mars", "JULIUS")
    add_stars(s, count=60, seed=17)
    add_top_marker(s, 10, "Zukunft – Mars")
    add_date_chip(s, Inches(11.0), Inches(0.5), "Ziel ab ~2030")
    add_title(s, Inches(1.2), "Ziel: Mars.", size=70)
    add_text(s, Inches(0.5), Inches(2.7), Inches(12), Inches(0.5),
             "Eine Stadt auf einem anderen Planeten.",
             size=18, color=ACCENT_GREY, italic=True, font="Calibri Light")
    add_image_placeholder(s, Inches(7.3), Inches(3.4), Inches(5.5), Inches(3.4),
                          "Mars / SpaceX-Mars-Konzept\n(NASA Image Library, CC0)",
                          ratio_hint="Querformat")
    phases = [
        ("PHASE 1", "Vorrats-Flüge"),
        ("PHASE 2", "Erste Crew"),
        ("PHASE 3", "Treibstoff auf Mars"),
        ("PHASE 4", "Selbsterhaltende Kolonie"),
    ]
    for i, (k, t) in enumerate(phases):
        y = Inches(3.5 + i * 0.85)
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y,
                                Inches(0.55), Inches(0.55))
        c.fill.background()
        c.line.color.rgb = ACCENT_RED
        c.line.width = Pt(1.5)
        add_text(s, Inches(0.5), y, Inches(0.55), Inches(0.55),
                 str(i+1), size=18, bold=True, color=ACCENT_RED,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font="Calibri")
        add_text(s, Inches(1.3), y - Inches(0.05), Inches(5.5), Inches(0.4),
                 k, size=10, bold=True, color=ACCENT_RED,
                 letter_spacing=400, font="Calibri")
        add_text(s, Inches(1.3), y + Inches(0.25), Inches(5.5), Inches(0.5),
                 t, size=18, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# 18. MOND (BEN)
# =====================================================
def s18_moon():
    s = new_slide("10 · Mond", "BEN")
    add_top_marker(s, 10, "Vorher: zum Mond")
    add_date_chip(s, Inches(11.0), Inches(0.5), "Artemis III ab 2027")
    add_title(s, Inches(1.2), "Zuerst zum Mond.", size=44)
    add_underline(s, Inches(2.25))
    add_image_placeholder(s, Inches(0.5), Inches(3.0), Inches(5.5), Inches(3.7),
                          "Starship-HLS Mondlander (NASA-Konzept)",
                          ratio_hint="Querformat")
    add_text(s, Inches(6.5), Inches(3.0), Inches(6), Inches(0.5),
             "ARTEMIS III", size=24, bold=True, color=ACCENT_WHITE,
             font="Calibri Light")
    add_text(s, Inches(6.5), Inches(3.5), Inches(6), Inches(0.4),
             "NASA-Programm · SpaceX baut Lander",
             size=12, color=ACCENT_GREY, italic=True, font="Calibri Light")
    bullets = [
        "1. Mondlandung seit 1972",
        "Vertrag SpaceX: ~4 Mrd. $",
        "Erstmals eine Frau auf dem Mond",
        "Ziel: dauerhafte Mond-Basis (Südpol)",
        "Starship als „Mondfähre“ (HLS)",
    ]
    for i, b in enumerate(bullets):
        add_bullet(s, Inches(6.5), Inches(4.15 + i * 0.5), Inches(6.3), b, size=15)


# =====================================================
# 19. PRAXIS - WAS HABEN WIR DAVON (JULIUS)
# =====================================================
def s19_use():
    s = new_slide("11 · Praxis", "JULIUS")
    add_top_marker(s, 11, "Was bringt uns das?")
    add_title(s, Inches(1.2), "Was haben wir davon?", size=42)
    add_underline(s, Inches(2.25))
    items = [
        ("📡", "Internet überall",   "Starlink für Dörfer, Wüsten, Krisengebiete."),
        ("🧪", "Forschung im All",    "Experimente in Schwerelosigkeit auf der ISS."),
        ("🌍", "Erdbeobachtung",     "Wetter, Klima, Landwirtschaft – günstig."),
        ("🚀", "Raumfahrt für alle", "Unis & Start-ups können sich Starts leisten."),
        ("💼", "Neue Jobs",          "Ingenieure, Informatiker, Materialwissen."),
        ("🛰️", "GPS & Co.",          "Wir nutzen Satelliten täglich – meist unsichtbar."),
    ]
    box_w = Inches(4.05)
    box_h = Inches(1.6)
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)
    start_x = Inches(0.5)
    start_y = Inches(2.85)
    for i, (icon, k, body) in enumerate(items):
        col = i % 3
        row = i // 3
        x = start_x + (box_w + gap_x) * col
        y = start_y + (box_h + gap_y) * row
        add_rect(s, x, y, box_w, box_h,
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_text(s, x + Inches(0.3), y + Inches(0.15), Inches(0.7), Inches(0.7),
                 icon, size=24, color=ACCENT_RED, font="Calibri")
        add_text(s, x + Inches(1.1), y + Inches(0.2), box_w - Inches(1.3),
                 Inches(0.5),
                 k, size=15, bold=True, color=ACCENT_WHITE, font="Calibri")
        add_text(s, x + Inches(1.1), y + Inches(0.7), box_w - Inches(1.3),
                 Inches(0.8),
                 body, size=11, color=ACCENT_GREY, font="Calibri Light")


# =====================================================
# 20. PRAXIS-TIPPS (BEN)
# =====================================================
def s20_tips():
    s = new_slide("11 · Tipps", "BEN")
    add_top_marker(s, 11, "Tipps · selbst aktiv werden")
    add_title(s, Inches(1.2), "Probiert es aus.", size=44)
    add_underline(s, Inches(2.25))
    tips = [
        ("01", "BEOBACHTEN",
         "Starlink am Himmel sehen",
         "App „Heavens-Above“ oder „Starlink Tracker“"),
        ("02", "ZUSCHAUEN",
         "Raketenstarts live",
         "SpaceX-Stream auf X / YouTube"),
        ("03", "LERNEN",
         "Selber bauen / simulieren",
         "Estes-Modellraketen · Kerbal Space Program"),
        ("04", "MITREDEN",
         "Quellen vergleichen",
         "SpaceX vs. NASA vs. ESA prüfen"),
    ]
    box_w = Inches(2.95)
    gap   = Inches(0.2)
    start = Inches(0.5)
    for i, (n, k, head, body) in enumerate(tips):
        x = start + (box_w + gap) * i
        add_rect(s, x, Inches(3.0), box_w, Inches(3.7),
                 fill=RGBColor(0x0E, 0x14, 0x22), line=LINE_GREY)
        add_rect(s, x, Inches(3.0), Inches(0.5), Inches(0.05),
                 fill=ACCENT_RED, line=None)
        add_text(s, x + Inches(0.3), Inches(3.15), Inches(2), Inches(0.4),
                 n, size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, x + Inches(0.3), Inches(3.6), Inches(2.6), Inches(0.4),
                 k, size=11, bold=True, color=ACCENT_WHITE,
                 letter_spacing=400, font="Calibri")
        add_text(s, x + Inches(0.3), Inches(4.1), Inches(2.6), Inches(1.2),
                 head, size=18, color=ACCENT_WHITE, font="Calibri Light")
        add_text(s, x + Inches(0.3), Inches(5.7), Inches(2.6), Inches(1.0),
                 body, size=11, color=ACCENT_GREY, italic=True,
                 font="Calibri Light")


# =====================================================
# 21. KRITIK (JULIUS)
# =====================================================
def s21_criticism():
    s = new_slide("12 · Kritik", "JULIUS")
    add_top_marker(s, 12, "Kritik & Risiken")
    add_title(s, Inches(1.2), "Nicht alles ist Glanz.", size=42)
    add_underline(s, Inches(2.25))
    items = [
        ("WELTRAUMMÜLL",   "7.000+ Satelliten → Kollisions-Risiko."),
        ("ASTRONOMIE",      "Helle Streifen auf Teleskopbildern."),
        ("UMWELT",          "Lärm, Schutzgebiete, CO₂ in Texas."),
        ("MONOPOL",         "Eine Firma dominiert den Markt."),
        ("PERSON MUSK",     "Politische Aussagen polarisieren."),
        ("VERSPRECHEN",     "„Mars bis 2024“ – immer wieder verschoben."),
    ]
    box_w = Inches(4.05)
    box_h = Inches(1.6)
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
                 Inches(0.4),
                 k, size=11, bold=True, color=ACCENT_RED,
                 letter_spacing=300, font="Calibri")
        add_text(s, x + Inches(0.3), y + Inches(0.7), box_w - Inches(0.6),
                 box_h - Inches(0.9),
                 t, size=14, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# 22. DISKUSSION (BEIDE)
# =====================================================
def s22_discussion():
    s = new_slide("Interaktiv · Diskussion", "BEIDE")
    add_stars(s, count=40, seed=22)
    add_kicker(s, Inches(0.9), "Diskutiert mit")
    add_underline(s, Inches(1.25))
    add_text(s, Inches(0.5), Inches(1.9), Inches(12.3), Inches(3),
             "Darf eine\nprivate Firma\nden Weltraum lenken?",
             size=52, bold=True, color=ACCENT_WHITE,
             font="Calibri Light", line_spacing=1.0)
    add_rect(s, Inches(0.5), Inches(5.7), Inches(6.0), Inches(1.4),
             fill=RGBColor(0x0C, 0x18, 0x10), line=ACCENT_BLUE, line_weight=0.75)
    add_text(s, Inches(0.8), Inches(5.85), Inches(5), Inches(0.4),
             "PRO", size=11, bold=True, color=ACCENT_BLUE,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(0.8), Inches(6.2), Inches(5.5), Inches(1.0),
             "schneller · günstiger · innovativer",
             size=18, color=ACCENT_WHITE, font="Calibri Light")
    add_rect(s, Inches(6.83), Inches(5.7), Inches(6.0), Inches(1.4),
             fill=RGBColor(0x18, 0x0C, 0x10), line=ACCENT_RED, line_weight=0.75)
    add_text(s, Inches(7.13), Inches(5.85), Inches(5), Inches(0.4),
             "CONTRA", size=11, bold=True, color=ACCENT_RED,
             letter_spacing=400, font="Calibri")
    add_text(s, Inches(7.13), Inches(6.2), Inches(5.5), Inches(1.0),
             "Abhängigkeit · Macht · Politik",
             size=18, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# 23. FAZIT (BEIDE)
# =====================================================
def s23_summary():
    s = new_slide("Fazit", "BEIDE")
    add_top_marker(s, 13, "Fazit")
    add_title(s, Inches(1.2), "Drei Sätze zum Mitnehmen.", size=38)
    add_underline(s, Inches(2.15))
    points = [
        ("01", "SpaceX hat Raumfahrt billiger gemacht."),
        ("02", "Heute der größte Akteur im Orbit."),
        ("03", "Mond, Mars, Internet: das Jahrzehnt entscheidet."),
    ]
    for i, (n, t) in enumerate(points):
        y = Inches(3.0 + i * 1.3)
        add_text(s, Inches(0.5), y, Inches(1.5), Inches(0.8),
                 n, size=48, bold=True, color=ACCENT_RED, font="Calibri")
        add_text(s, Inches(2.3), y + Inches(0.1), Inches(10.3), Inches(1.2),
                 t, size=24, color=ACCENT_WHITE, font="Calibri Light")


# =====================================================
# 24. QUELLEN (JULIUS)
# =====================================================
def s24_sources():
    s = new_slide("Quellen", "JULIUS")
    add_top_marker(s, 14, "Quellen")
    add_title(s, Inches(1.2), "Quellen & weiterführende Links",
              size=30)
    add_underline(s, Inches(2.05))
    sources = [
        ("SpaceX – Offiziell",                "https://www.spacex.com"),
        ("NASA – Commercial Crew",            "https://www.nasa.gov/commercial-crew"),
        ("NASA – Artemis",                    "https://www.nasa.gov/humans-in-space/artemis"),
        ("ESA",                               "https://www.esa.int"),
        ("DLR (Deutschland)",                 "https://www.dlr.de"),
        ("BryceTech – Start Reports",         "https://brycetech.com/reports"),
        ("Spiegel – Weltall",                 "https://www.spiegel.de/wissenschaft/weltall"),
        ("Tagesschau – Raumfahrt",            "https://www.tagesschau.de/wissen/raumfahrt"),
        ("Reuters – Aerospace",               "https://www.reuters.com/business/aerospace-defense"),
        ("BBC – Science",                     "https://www.bbc.com/news/science_and_environment"),
        ("Wikipedia – SpaceX (DE)",           "https://de.wikipedia.org/wiki/SpaceX"),
        ("J. McDowell – Satellite Tracker",   "https://planet4589.org/space/stats"),
    ]
    col1 = sources[:6]
    col2 = sources[6:]
    def render(col, x0):
        for i, (name, url) in enumerate(col):
            y = Inches(2.7 + i * 0.55)
            add_text(s, x0, y, Inches(6.3), Inches(0.3),
                     name, size=13, bold=True, color=ACCENT_WHITE,
                     font="Calibri Light")
            add_text(s, x0, y + Inches(0.27), Inches(6.3), Inches(0.3),
                     url, size=10, color=ACCENT_BLUE, font="Calibri Light")
    render(col1, Inches(0.5))
    render(col2, Inches(6.8))
    add_text(s, Inches(0.5), Inches(6.6), Inches(12), Inches(0.4),
             "Abgerufen: Mai 2026.",
             size=10, italic=True, color=ACCENT_GREY, font="Calibri")


# =====================================================
# 25. FRAGEN (BEIDE)
# =====================================================
def s25_questions():
    s = new_slide("Fragen?", "BEIDE")
    add_stars(s, count=220, seed=99)
    add_text(s, Inches(0.5), Inches(2.6), Inches(12.3), Inches(2),
             "Fragen?",
             size=200, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, font="Calibri")
    add_rect(s, Inches(6.16), Inches(5.2), Inches(1.0), Inches(0.05),
             fill=ACCENT_RED, line=None)
    add_text(s, Inches(0.5), Inches(5.4), Inches(12.3), Inches(0.5),
             "Wir freuen uns auf eure Diskussion.",
             size=20, color=ACCENT_GREY, align=PP_ALIGN.CENTER,
             italic=True, font="Calibri Light")
    add_text(s, Inches(0.5), Inches(6.8), Inches(12.3), Inches(0.4),
             "JULIUS MÜLLER   ·   BEN HAMMEN",
             size=13, bold=True, color=ACCENT_WHITE,
             align=PP_ALIGN.CENTER, letter_spacing=400, font="Calibri")


# =====================================================
# BAUEN
# =====================================================
s01_title(); s02_welcome(); s03_agenda(); s04_what(); s05_history()
s06_musk();  s07_mission(); s08_rockets(); s09_falcon9(); s10_reuse()
s11_starship(); s12_dragon(); s13_starlink(); s14_quiz(); s15_milestones()
s16_competition(); s17_mars(); s18_moon(); s19_use(); s20_tips()
s21_criticism(); s22_discussion(); s23_summary(); s24_sources(); s25_questions()

total = len(SLIDES)
for i, (slide, chapter, speaker) in enumerate(SLIDES):
    if i == 0 or i == total - 1:
        continue
    add_footer(slide, i + 1, total, chapter, speaker)


def add_fade_transition(slide):
    sld = slide._element
    for t in sld.findall(qn("p:transition")):
        sld.remove(t)
    transition_xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'spd="med" advClick="1"><p:fade/></p:transition>'
    )
    el = etree.fromstring(transition_xml)
    sld.append(el)
for slide, _, _ in SLIDES:
    add_fade_transition(slide)

out_path = "/home/user/Pr-si-Mathe-/SpaceX_Praesentation_Mueller_Hammen.pptx"
prs.save(out_path)
print(f"OK -> {out_path}")
print(f"Folien gesamt: {total}")
print(f"Bild-Platzhalter: {IMG_COUNTER[0]}")
