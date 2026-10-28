from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Color Palette - Nature/Environment theme
DARK_GREEN = RGBColor(0x1B, 0x4D, 0x3E)       # Deep forest green
MED_GREEN  = RGBColor(0x2E, 0x86, 0x5A)       # Medium green
LIGHT_GREEN= RGBColor(0xA8, 0xD5, 0xBA)       # Pale mint
ACCENT     = RGBColor(0xF4, 0xA2, 0x61)       # Warm amber
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT  = RGBColor(0x1A, 0x1A, 0x2E)
GRAY_BG    = RGBColor(0xF4, 0xF7, 0xF4)
GRAY_TEXT  = RGBColor(0x55, 0x65, 0x5A)
LIGHT_AMBER= RGBColor(0xFF, 0xF3, 0xE0)

W  = Inches(13.33)   # WIDE layout width
H  = Inches(7.5)     # WIDE layout height

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

def blank_slide(prs):
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)

def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def rect(slide, x, y, w, h, color, transparency=0):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def txt(slide, text, x, y, w, h, size=16, color=RGBColor(0x1A,0x1A,0x2E),
        bold=False, italic=False, align=PP_ALIGN.LEFT, font="Calibri"):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    return txBox

def add_para(tf, text, size=14, color=RGBColor(0x1A,0x1A,0x2E), bold=False,
             italic=False, align=PP_ALIGN.LEFT, bullet=False, space_before=0):
    p = tf.add_paragraph()
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = "Calibri"
    return p

def stat_box(slide, x, y, w, h, value, label, bg_color=MED_GREEN, text_color=WHITE):
    rect(slide, x, y, w, h, bg_color)
    txt(slide, value, x, y+0.1, w, h*0.55, size=28, color=text_color, bold=True, align=PP_ALIGN.CENTER)
    txt(slide, label, x, y+h*0.55, w, h*0.45, size=10, color=text_color, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────
# SLIDE 1 — Title Slide
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, DARK_GREEN)

# Large accent rect right side
rect(s, 9.0, 0, 4.33, 7.5, MED_GREEN)

# Decorative circles (water drops style)
for ci, (cx, cy, cr, col) in enumerate([
    (9.4, 0.3, 1.2, LIGHT_GREEN),
    (11.8, 5.8, 1.5, RGBColor(0x24,0x6A,0x4A)),
]):
    shape = s.shapes.add_shape(9, Inches(cx), Inches(cy), Inches(cr), Inches(cr))  # oval
    shape.fill.solid()
    shape.fill.fore_color.rgb = col
    shape.line.fill.background()

# Title text
txt(s, "DAIRY WASTEWATER TREATMENT", 0.5, 0.8, 8.2, 1.2,
    size=34, color=WHITE, bold=True, align=PP_ALIGN.LEFT, font="Calibri")
txt(s, "Using Biochar–Clay Disc", 0.5, 1.9, 8.2, 0.8,
    size=24, color=ACCENT, bold=False, italic=True, align=PP_ALIGN.LEFT)

# Amber divider line
rect(s, 0.5, 2.75, 5.0, 0.06, ACCENT)

# Subtitle info
txb = s.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(8.2), Inches(2.5))
tf = txb.text_frame
tf.word_wrap = True
for line, sz, bold, col in [
    ("M.Sc. Environmental Science & Technology Dissertation", 13, False, LIGHT_GREEN),
    ("", 6, False, WHITE),
    ("Jaisal Johnson  |  Reg. No. 24100721", 14, True, WHITE),
    ("School of Environmental Studies, CUSAT", 12, False, LIGHT_GREEN),
    ("", 6, False, WHITE),
    ("Guide: Dr. Sylas V P, ACESSD, MG University, Kottayam", 12, False, LIGHT_GREEN),
    ("April 2026", 12, True, ACCENT),
]:
    add_para(tf, line, size=sz, color=col, bold=bold)

# Right side graphic text
txt(s, "🌿", 9.8, 1.5, 2.5, 2.0, size=80, color=WHITE, align=PP_ALIGN.CENTER)
txt(s, "Nature-Based\nWastewater\nSolution", 9.4, 3.8, 3.5, 2.5,
    size=12, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

# ─────────────────────────────────────────────────────────────
# SLIDE 2 — Introduction / Problem Statement
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, GRAY_BG)

rect(s, 0, 0, 13.33, 1.0, DARK_GREEN)
txt(s, "INTRODUCTION", 0.5, 0.15, 10, 0.7, size=28, color=WHITE, bold=True)
txt(s, "The Global Challenge", 0.5, 0.5, 10, 0.5, size=14, color=LIGHT_GREEN, italic=True)

# Two columns
# Left column - problem
rect(s, 0.4, 1.2, 5.8, 5.6, WHITE)
txt(s, "The Problem", 0.7, 1.3, 5.2, 0.5, size=16, color=DARK_GREEN, bold=True)
rect(s, 0.4, 1.75, 0.07, 4.8, ACCENT)

txb = s.shapes.add_textbox(Inches(0.7), Inches(1.8), Inches(5.2), Inches(4.8))
tf = txb.text_frame; tf.word_wrap = True
items = [
    ("Global Water Crisis", "Over 50% of humanity faces water shortages for at least 1 month per year", True),
    ("Dairy Sector Impact", "India produces ~230 million tonnes of milk annually, generating massive effluent volumes", True),
    ("Effluent Characteristics", "High BOD (1,000–4,000 mg/L), COD (2,000–8,000 mg/L), suspended solids, nutrients", True),
    ("Environmental Threat", "Untreated discharge causes eutrophication, oxygen depletion & aquatic harm", True),
    ("Treatment Gap", "Small dairies lack access to costly centralised treatment infrastructure", True),
]
first = True
for title, desc, _ in items:
    if not first:
        add_para(tf, "", size=4, color=DARK_TEXT)
    p = tf.add_paragraph()
    run = p.add_run(); run.text = f"▶ {title}"
    run.font.size = Pt(13); run.font.bold = True; run.font.color.rgb = MED_GREEN; run.font.name = "Calibri"
    p2 = tf.add_paragraph()
    run2 = p2.add_run(); run2.text = f"   {desc}"
    run2.font.size = Pt(11); run2.font.color.rgb = GRAY_TEXT; run2.font.name = "Calibri"
    first = False

# Right column - solution angle
rect(s, 6.8, 1.2, 6.1, 5.6, DARK_GREEN)
txt(s, "Why This Study?", 7.1, 1.3, 5.5, 0.5, size=16, color=ACCENT, bold=True)
txb2 = s.shapes.add_textbox(Inches(7.1), Inches(1.9), Inches(5.5), Inches(4.5))
tf2 = txb2.text_frame; tf2.word_wrap = True
solution_pts = [
    "Kerala produces 5,000+ million coconuts/year — shells are agricultural waste",
    "Coconut shell biochar: high surface area (>500 m²/g), excellent adsorption",
    "Clay: abundant, low cost, structural integrity for filter fabrication",
    "Gravity-driven system: no energy input needed",
    "Ideal for decentralised dairy treatment in rural India",
    "Aligned with UN SDG 6 (Clean Water) & SDG 12 (Responsible Consumption)",
]
for pt in solution_pts:
    p = tf2.add_paragraph()
    run = p.add_run(); run.text = f"✓  {pt}"
    run.font.size = Pt(11.5); run.font.color.rgb = LIGHT_GREEN; run.font.name = "Calibri"
    p.space_before = Pt(5)

# ─────────────────────────────────────────────────────────────
# SLIDE 3 — Aim & Objectives
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, WHITE)

rect(s, 0, 0, 13.33, 1.0, MED_GREEN)
txt(s, "AIM & OBJECTIVES", 0.5, 0.15, 12, 0.7, size=28, color=WHITE, bold=True)

# Aim box
rect(s, 0.4, 1.2, 12.5, 1.1, LIGHT_GREEN)
txt(s, "AIM", 0.65, 1.28, 1.2, 0.5, size=13, color=DARK_GREEN, bold=True)
txt(s, "To design and evaluate a low-cost clay-biochar disc filter for treating dairy industry wastewater, focusing on removing organic load, suspended solids, dissolved ions, and nutrients to meet environmental discharge standards.",
    2.0, 1.25, 10.7, 1.0, size=13, color=DARK_TEXT)

# Three objective boxes
obj_data = [
    ("01", "Biochar Production",
     "Produce biochar from coconut shells via controlled pyrolysis at 900°C under inert N₂ atmosphere and develop biochar-clay filters with varying compositions"),
    ("02", "Material Characterisation",
     "Characterise structural morphology and chemical properties using FESEM and ATR-FTIR spectroscopy before and after filtration"),
    ("03", "System Development & Analysis",
     "Develop a gravity-fed PVC column filtration system and conduct comprehensive post-filtration physicochemical analysis"),
]
for i, (num, title, desc) in enumerate(obj_data):
    x = 0.4 + i * 4.3
    rect(s, x, 2.5, 4.0, 4.5, DARK_GREEN)
    txt(s, num, x+0.15, 2.6, 1.0, 0.8, size=36, color=ACCENT, bold=True)
    txt(s, title, x+0.2, 3.35, 3.6, 0.6, size=14, color=WHITE, bold=True)
    rect(s, x+0.2, 3.95, 3.5, 0.04, ACCENT)
    txb = s.shapes.add_textbox(Inches(x+0.2), Inches(4.1), Inches(3.6), Inches(2.7))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; run = p.add_run(); run.text = desc
    run.font.size = Pt(12); run.font.color.rgb = LIGHT_GREEN; run.font.name = "Calibri"

# ─────────────────────────────────────────────────────────────
# SLIDE 4 — Methodology (Biochar Production)
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, GRAY_BG)

rect(s, 0, 0, 13.33, 1.0, DARK_GREEN)
txt(s, "METHODOLOGY", 0.5, 0.15, 10, 0.7, size=28, color=WHITE, bold=True)
txt(s, "Biochar Production & Disc Fabrication", 0.5, 0.5, 10, 0.45, size=14, color=LIGHT_GREEN, italic=True)

# Flow process boxes
steps = [
    ("1", "Coconut Shell\nCollection", "Cocos nucifera shells from MG University campus; washed with distilled water"),
    ("2", "Drying", "Hot-air oven at 105°C for 24 hours to remove residual moisture"),
    ("3", "Crushing", "Manually crushed into uniform pieces using a steel hammer"),
    ("4", "Pyrolysis", "900°C for 2 hours in N₂ atmosphere muffle furnace — prevents oxidation"),
    ("5", "Biochar\nPreparation", "Ground & sieved through 0.25 mm mesh for uniform particle size"),
]
box_w = 2.2
gap = 0.25
for i, (num, title, desc) in enumerate(steps):
    x = 0.4 + i * (box_w + gap)
    rect(s, x, 1.2, box_w, 3.8, WHITE)
    rect(s, x, 1.2, box_w, 0.55, MED_GREEN)
    txt(s, num, x+0.05, 1.2, 0.5, 0.55, size=20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, title, x+0.5, 1.25, box_w-0.6, 0.5, size=11, color=WHITE, bold=True)
    txb = s.shapes.add_textbox(Inches(x+0.1), Inches(1.85), Inches(box_w-0.2), Inches(3.0))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; run = p.add_run(); run.text = desc
    run.font.size = Pt(10.5); run.font.color.rgb = GRAY_TEXT; run.font.name = "Calibri"
    if i < len(steps)-1:
        txt(s, "→", x+box_w, 2.7, gap+0.05, 0.4, size=18, color=ACCENT, bold=True, align=PP_ALIGN.CENTER)

# Disc Fabrication box
rect(s, 0.4, 5.2, 12.5, 2.0, DARK_GREEN)
txt(s, "Disc Fabrication", 0.7, 5.28, 3, 0.5, size=14, color=ACCENT, bold=True)
disc_info = "Four formulations prepared (biochar:clay ratios) → Mixed with fine sand & sawdust → Moulded into discs (Ø 6.5 cm, thickness 0.5 cm) → Air-dried 24 hrs → Sintered at 850°C for 2 hrs → Assembled in gravity-fed PVC column"
txt(s, disc_info, 0.7, 5.8, 12.0, 1.2, size=12, color=LIGHT_GREEN)

# ─────────────────────────────────────────────────────────────
# SLIDE 5 — Methodology (Filter System & Analytical Methods)
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, WHITE)

rect(s, 0, 0, 13.33, 1.0, MED_GREEN)
txt(s, "FILTER SYSTEM & ANALYTICAL METHODS", 0.5, 0.15, 12, 0.7, size=24, color=WHITE, bold=True)

# Left: Four disc formulations table
rect(s, 0.4, 1.2, 5.8, 5.8, GRAY_BG)
txt(s, "Four Disc Formulations Tested", 0.6, 1.3, 5.4, 0.5, size=14, color=DARK_GREEN, bold=True)

headers = ["Disc", "Clay %", "Biochar %", "Sawdust %", "Sand %"]
rows = [
    ["A", "50", "30", "10", "10"],
    ["B", "60", "25", "10", "5"],
    ["C ★", "65", "16", "11", "8"],
    ["D", "70", "20", "5", "5"],
]
col_ws = [0.7, 1.0, 1.1, 1.1, 0.9]
col_xs = [0.5, 1.2, 2.2, 3.3, 4.4]
# header
for j, (hdr, cx, cw) in enumerate(zip(headers, col_xs, col_ws)):
    rect(s, cx, 1.85, cw-0.05, 0.38, DARK_GREEN)
    txt(s, hdr, cx+0.02, 1.88, cw-0.07, 0.35, size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
for i, row in enumerate(rows):
    row_color = LIGHT_GREEN if row[0].startswith("C") else WHITE
    for j, (cell, cx, cw) in enumerate(zip(row, col_xs, col_ws)):
        rect(s, cx, 2.28+i*0.42, cw-0.05, 0.4, row_color)
        fc = DARK_GREEN if row[0].startswith("C") else DARK_TEXT
        txt(s, cell, cx+0.02, 2.3+i*0.42, cw-0.07, 0.38, size=11,
            color=fc, bold=row[0].startswith("C"), align=PP_ALIGN.CENTER)

txt(s, "★ Optimal formulation selected for full analysis", 0.6, 4.1, 5.4, 0.4,
    size=10, color=MED_GREEN, italic=True)

# System description
rect(s, 0.4, 4.55, 5.8, 2.3, MED_GREEN)
txt(s, "Filtration System", 0.65, 4.65, 5.3, 0.4, size=13, color=WHITE, bold=True)
sys_pts = [
    "Gravity-fed PVC column — no energy required",
    "Sample: Milma Factory, Vadavathoor, Kottayam",
    "Stored at 4°C to prevent degradation",
    "Flow rate measured volumetrically (mL/hr)",
]
txb = s.shapes.add_textbox(Inches(0.65), Inches(5.1), Inches(5.3), Inches(1.65))
tf = txb.text_frame; tf.word_wrap = True
for pt in sys_pts:
    p = tf.add_paragraph(); run = p.add_run(); run.text = f"• {pt}"
    run.font.size = Pt(11); run.font.color.rgb = WHITE; run.font.name = "Calibri"

# Right: Analytical methods
rect(s, 6.8, 1.2, 6.1, 5.9, DARK_GREEN)
txt(s, "Water Quality Parameters Tested", 7.0, 1.3, 5.7, 0.5, size=14, color=ACCENT, bold=True)
methods = [
    ("pH & EC", "Digital pH/conductivity meter"),
    ("Turbidity", "Nephelometric turbidimeter (NTU)"),
    ("TDS & Salinity", "TDS meter + Mohr's titration"),
    ("Hardness (Ca/Mg)", "EDTA complexometric titration"),
    ("Chloride", "Argentometric (Mohr's) titration"),
    ("Nitrate", "Brucine sulphanilic acid method, UV 410 nm"),
    ("Phosphate", "Stannous chloride reduction, 690 nm"),
    ("COD", "Closed reflux dichromate method"),
    ("BOD", "5-day incubation at 20°C (BOD₅)"),
    ("FESEM", "Surface morphology — MAIA3 TESCAN, 10 kV"),
    ("FTIR", "ATR-FTIR, 400–4000 cm⁻¹ — functional groups"),
    ("PLR", "Pollution Load Removal = Q × (C_before − C_after)"),
]
txb2 = s.shapes.add_textbox(Inches(7.0), Inches(1.9), Inches(5.7), Inches(4.9))
tf2 = txb2.text_frame; tf2.word_wrap = True
for param, method in methods:
    p = tf2.add_paragraph()
    run1 = p.add_run(); run1.text = f"{param}: "
    run1.font.size = Pt(11); run1.font.bold = True; run1.font.color.rgb = ACCENT; run1.font.name = "Calibri"
    run2 = p.add_run(); run2.text = method
    run2.font.size = Pt(11); run2.font.color.rgb = LIGHT_GREEN; run2.font.name = "Calibri"
    p.space_before = Pt(2)

# ─────────────────────────────────────────────────────────────
# SLIDE 6 — Results: Disc Performance Comparison
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, GRAY_BG)

rect(s, 0, 0, 13.33, 1.0, DARK_GREEN)
txt(s, "RESULTS: DISC PERFORMANCE COMPARISON", 0.5, 0.15, 12, 0.7, size=24, color=WHITE, bold=True)

# Bar chart — BOD removal %
chart_data_labels = ["Disc A\n(50% clay)", "Disc B\n(60% clay)", "Disc C★\n(65% clay)", "Disc D\n(70% clay)"]
removal_pct = [70, 78, 86.25, 83.75]
flow_rates   = [8.0, 5.2, 4.8, 2.0]
colors_bars  = [RGBColor(0x84,0xB5,0x9F), RGBColor(0x2E,0x86,0x5A), RGBColor(0x1B,0x4D,0x3E), RGBColor(0x84,0xB5,0x9F)]

chart_x, chart_y, chart_w, chart_h = 0.5, 1.2, 7.5, 5.5
max_pct = 100
bar_area_h = 3.8
bar_area_y_bottom = chart_y + 0.5 + bar_area_h
bar_w = 1.2
bar_gap = 0.4
bar_start_x = chart_x + 0.8

for i, (label, pct, color) in enumerate(zip(chart_data_labels, removal_pct, colors_bars)):
    bx = bar_start_x + i * (bar_w + bar_gap)
    bh = (pct / max_pct) * bar_area_h
    by = bar_area_y_bottom - bh
    fc = ACCENT if label.startswith("Disc C") else color
    rect(s, bx, by, bar_w, bh, fc)
    txt(s, f"{pct}%", bx, by-0.35, bar_w, 0.35, size=13, color=DARK_GREEN, bold=True, align=PP_ALIGN.CENTER)
    txt(s, label, bx, bar_area_y_bottom+0.05, bar_w, 0.55, size=9, color=GRAY_TEXT, align=PP_ALIGN.CENTER)

# Y axis label
txt(s, "BOD Removal Efficiency (%)", chart_x, chart_y+0.5, 0.7, bar_area_h, size=10,
    color=GRAY_TEXT, align=PP_ALIGN.CENTER)
# Y axis line
rect(s, chart_x+0.8, chart_y+0.5, 0.03, bar_area_h+0.05, GRAY_TEXT)
# X axis line
rect(s, chart_x+0.8, bar_area_y_bottom, 7.0, 0.03, GRAY_TEXT)

txt(s, "BOD Removal Efficiency by Disc Formulation", 0.5, 5.0, 7.5, 0.4,
    size=11, color=GRAY_TEXT, italic=True, align=PP_ALIGN.CENTER)

# Right panel — key insights
rect(s, 8.4, 1.2, 4.5, 5.5, WHITE)
txt(s, "Key Findings", 8.65, 1.3, 4.0, 0.5, size=15, color=DARK_GREEN, bold=True)
rect(s, 8.65, 1.8, 3.8, 0.04, ACCENT)

insights = [
    ("Disc A (30% biochar)", "Highest flow 8 mL/h but lowest removal 70% — short contact time"),
    ("Disc B (25% biochar)", "78% BOD removal at 5.2 mL/h — improved with more clay"),
    ("Disc C ★ OPTIMAL", "86.25% BOD removal at 4.8 mL/h — best balance of adsorption & flow"),
    ("Disc D (20% biochar)", "83.75% removal but very slow flow 2 mL/h — too restrictive"),
    ("Key Insight", "Inverse relationship: ↑ clay content → ↓ flow rate → ↑ contact time"),
]
txb = s.shapes.add_textbox(Inches(8.65), Inches(1.95), Inches(4.0), Inches(4.5))
tf = txb.text_frame; tf.word_wrap = True
for title, desc in insights:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = title
    r.font.size = Pt(11.5); r.font.bold = True; r.font.name = "Calibri"
    r.font.color.rgb = MED_GREEN if "OPTIMAL" in title else DARK_GREEN
    p.space_before = Pt(7)
    p2 = tf.add_paragraph()
    r2 = p2.add_run(); r2.text = f"   {desc}"
    r2.font.size = Pt(10.5); r2.font.color.rgb = GRAY_TEXT; r2.font.name = "Calibri"

# ─────────────────────────────────────────────────────────────
# SLIDE 7 — Results: Physicochemical Removal (Key Parameters)
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, WHITE)

rect(s, 0, 0, 13.33, 1.0, MED_GREEN)
txt(s, "RESULTS: PHYSICOCHEMICAL PARAMETER REMOVAL", 0.5, 0.15, 12, 0.7, size=22, color=WHITE, bold=True)

# Stat boxes — top row
stats = [
    ("98.59%", "Turbidity\n814.67 → 11.47 NTU", DARK_GREEN),
    ("94.35%", "COD\n1770 → 100 mg/L", MED_GREEN),
    ("86.25%", "BOD\n400 → 55 mg/L", RGBColor(0x2A,0x7A,0x55)),
    ("90.44%", "Total Hardness\n52.33 → 5.00 mg/L", RGBColor(0x1E,0x6E,0x4A)),
    ("88.22%", "Sulphate\n254.67 → 30 mg/L", ACCENT),
]
bw = 2.4; by = 1.15
for i, (val, lbl, col) in enumerate(stats):
    bx = 0.35 + i*(bw+0.09)
    stat_box(s, bx, by, bw, 1.4, val, lbl, bg_color=col)

# Table — all parameters
rect(s, 0.3, 2.75, 12.7, 0.42, DARK_GREEN)
col_headers = ["Parameter", "Before Treatment", "After Treatment", "Removal Efficiency"]
col_xs_t = [0.35, 3.5, 6.3, 9.2]
col_ws_t = [3.1, 2.7, 2.9, 3.7]
for hdr, cx, cw in zip(col_headers, col_xs_t, col_ws_t):
    txt(s, hdr, cx+0.05, 2.8, cw-0.1, 0.35, size=11, color=WHITE, bold=True)

table_rows = [
    ("pH", "11.95", "8.50", "28.87%"),
    ("TDS (ppt)", "3.20", "1.20", "62.50%"),
    ("Salinity (ppt)", "3.76", "1.48", "60.64%"),
    ("EC (mS/cm)", "5.02", "2.20", "56.17%"),
    ("Nitrate (mg/L)", "26.49", "13.66", "48.43%"),
    ("Sulphate (mg/L)", "254.67", "30.00", "88.22%"),
    ("Phosphate (mg/L)", "13.05", "4.00", "69.35%"),
    ("Chloride (mg/L)", "77.53", "12.43", "83.97%"),
    ("Ca Hardness (mg/L)", "12.00", "2.00", "83.33%"),
    ("Mg Hardness (mg/L)", "40.00", "3.00", "92.50%"),
]
for ri, row in enumerate(table_rows):
    bg_col = GRAY_BG if ri % 2 == 0 else WHITE
    row_y = 3.2 + ri * 0.37
    rect(s, 0.3, row_y, 12.7, 0.37, bg_col)
    for cell, cx, cw in zip(row, col_xs_t, col_ws_t):
        fc = MED_GREEN if "%" in cell and float(cell.replace("%","")) >= 80 else DARK_TEXT
        txt(s, cell, cx+0.05, row_y+0.05, cw-0.1, 0.3, size=10.5,
            color=fc, bold="%" in cell and float(cell.replace("%","")) >= 80)

# ─────────────────────────────────────────────────────────────
# SLIDE 8 — Results: Turbidity, COD, BOD Deep Dive
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, GRAY_BG)

rect(s, 0, 0, 13.33, 1.0, DARK_GREEN)
txt(s, "RESULTS: ORGANIC LOAD & TURBIDITY REMOVAL", 0.5, 0.15, 12, 0.7, size=24, color=WHITE, bold=True)

# Large visual callouts
callouts = [
    ("98.59%", "Turbidity Removed", "814.67 NTU → 11.47 NTU", ACCENT),
    ("94.35%", "COD Removed", "1770 mg/L → 100 mg/L", MED_GREEN),
    ("86.25%", "BOD Removed", "400 mg/L → 55 mg/L", DARK_GREEN),
]
for i, (pct, title, detail, col) in enumerate(callouts):
    x = 0.4 + i*4.3
    rect(s, x, 1.2, 4.0, 2.5, col)
    txt(s, pct, x, 1.35, 4.0, 1.2, size=46, color=WHITE, bold=True, align=PP_ALIGN.CENTER, font="Calibri")
    txt(s, title, x, 2.5, 4.0, 0.5, size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, detail, x, 3.0, 4.0, 0.5, size=11, color=LIGHT_GREEN, italic=True, align=PP_ALIGN.CENTER)

# Explanation boxes below
explanations = [
    ("Turbidity — Physical Filtration",
     "Near-complete turbidity removal demonstrates exceptional physical filtration. "
     "The clay–biochar pore network mechanically sieves and retains suspended particulates, fat globules, and colloidal matter."),
    ("COD — Adsorption Mechanisms",
     "94.35% COD reduction reflects effective adsorption of dissolved and colloidal organic matter. "
     "Biochar's high surface area and diverse functional groups (–OH, –COOH, C=O) bind both polar and non-polar organics."),
    ("BOD — Organic Matter Removal",
     "86.25% BOD removal confirms effective breakdown of biodegradable fraction. "
     "Low BOD/COD ratio (0.226) of influent makes adsorption-based removal particularly appropriate."),
]
for i, (title, desc) in enumerate(explanations):
    x = 0.4 + i*4.3
    rect(s, x, 3.7, 4.0, 3.5, WHITE)
    rect(s, x, 3.7, 4.0, 0.45, callouts[i][3])
    txt(s, title, x+0.1, 3.75, 3.8, 0.4, size=11, color=WHITE, bold=True)
    txb = s.shapes.add_textbox(Inches(x+0.1), Inches(4.25), Inches(3.8), Inches(2.8))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; run = p.add_run(); run.text = desc
    run.font.size = Pt(11); run.font.color.rgb = GRAY_TEXT; run.font.name = "Calibri"

# ─────────────────────────────────────────────────────────────
# SLIDE 9 — Results: Nutrients & Pollution Load Removal
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, WHITE)

rect(s, 0, 0, 13.33, 1.0, MED_GREEN)
txt(s, "RESULTS: NUTRIENTS & POLLUTION LOAD REMOVAL (PLR)", 0.5, 0.15, 12, 0.7, size=22, color=WHITE, bold=True)

# Left: Nutrient removal
rect(s, 0.3, 1.2, 5.8, 6.0, GRAY_BG)
txt(s, "Nutrient Removal", 0.55, 1.3, 5.3, 0.5, size=15, color=DARK_GREEN, bold=True)

nutrients = [
    ("Sulphate", 88.22, "Anion adsorption onto +ve biochar surfaces"),
    ("Phosphate", 69.35, "Ligand exchange on clay mineral surfaces"),
    ("Nitrate", 48.43, "Weakest — needs additional denitrification"),
    ("Chloride", 83.97, "Co-precipitation & physical entrapment"),
]
bar_h_unit = 0.45
for i, (param, pct, mech) in enumerate(nutrients):
    y = 1.9 + i * 1.1
    rect(s, 0.5, y, 5.3, bar_h_unit, WHITE)
    col = ACCENT if param == "Nitrate" else MED_GREEN
    bw = (pct/100) * 4.5
    rect(s, 0.5, y, bw, bar_h_unit, col)
    txt(s, f"{param}: {pct}%", 0.55, y+0.02, 4.8, bar_h_unit-0.05, size=11,
        color=WHITE, bold=True)
    txt(s, mech, 0.55, y+bar_h_unit+0.02, 5.1, 0.45, size=10, color=GRAY_TEXT, italic=True)

rect(s, 0.3, 5.55, 5.8, 1.45, DARK_GREEN)
txt(s, "⚠  Note on Nitrate (48.43%)", 0.55, 5.65, 5.3, 0.45, size=12, color=ACCENT, bold=True)
txt(s, "Nitrate shows lowest removal due to weak affinity for filter matrix. "
       "Biological denitrification or ion-exchange steps are recommended for complete nitrogen removal.",
    0.55, 6.1, 5.3, 0.8, size=10.5, color=LIGHT_GREEN)

# Right: PLR table
rect(s, 6.5, 1.2, 6.5, 6.0, DARK_GREEN)
txt(s, "Pollution Load Removal (PLR)", 6.75, 1.3, 6.0, 0.5, size=14, color=ACCENT, bold=True)
txt(s, "PLR = Q × (C_before − C_after)", 6.75, 1.8, 6.0, 0.35, size=11, color=LIGHT_GREEN, italic=True)
txt(s, "Flow rate Q = 4.80 mL/hr", 6.75, 2.1, 6.0, 0.35, size=11, color=LIGHT_GREEN)

plr_data = [
    ("Salinity", "10,944", "262.66"),
    ("TDS", "9,600", "230.40"),
    ("COD", "8,016", "192.38"),
    ("Turbidity", "3,855", "92.53"),
    ("BOD", "1,656", "39.74"),
    ("Sulphate", "1,078", "25.88"),
    ("Chloride", "312", "7.50"),
    ("Total Hardness", "227", "5.45"),
    ("Phosphate", "43", "1.04"),
]
rect(s, 6.5, 2.5, 6.5, 0.38, MED_GREEN)
for hdr, hx, hw in [("Parameter", 6.55, 2.5), ("PLR (μg/hr)", 9.1, 1.7), ("PLR (mg/day)", 10.85, 2.1)]:
    txt(s, hdr, hx, 2.55, hw, 0.32, size=10, color=WHITE, bold=True)
for i, (param, plr_hr, plr_day) in enumerate(plr_data):
    ry = 2.9 + i*0.36
    bg_c = RGBColor(0x24,0x6A,0x4A) if i%2==0 else RGBColor(0x1B,0x4D,0x3E)
    rect(s, 6.5, ry, 6.5, 0.36, bg_c)
    txt(s, param, 6.55, ry+0.03, 2.5, 0.32, size=10, color=WHITE)
    txt(s, plr_hr, 9.1, ry+0.03, 1.7, 0.32, size=10, color=LIGHT_GREEN, bold=True, align=PP_ALIGN.RIGHT)
    txt(s, plr_day, 10.85, ry+0.03, 2.1, 0.32, size=10, color=ACCENT, bold=True, align=PP_ALIGN.RIGHT)

# ─────────────────────────────────────────────────────────────
# SLIDE 10 — Results: Material Characterisation (FESEM & FTIR)
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, GRAY_BG)

rect(s, 0, 0, 13.33, 1.0, DARK_GREEN)
txt(s, "MATERIAL CHARACTERISATION: FESEM & FTIR", 0.5, 0.15, 12, 0.7, size=24, color=WHITE, bold=True)

# Physical properties
props = [
    ("50.70%", "Apparent Porosity", "Open interconnected pores\nfor fluid flow & adsorption"),
    ("36.89%", "Water Absorption", "Well-developed pore network\nretains wastewater for contact"),
    ("4.80 mL/hr", "Flow Rate", "Low flow → long hydraulic\nretention time → high removal"),
]
for i, (val, title, desc) in enumerate(props):
    x = 0.3 + i*4.35
    rect(s, x, 1.15, 4.1, 2.0, WHITE)
    rect(s, x, 1.15, 4.1, 0.5, MED_GREEN)
    txt(s, title, x+0.1, 1.2, 3.9, 0.45, size=13, color=WHITE, bold=True)
    txt(s, val, x, 1.7, 4.1, 0.65, size=28, color=DARK_GREEN, bold=True, align=PP_ALIGN.CENTER)
    txt(s, desc, x+0.1, 2.38, 3.9, 0.7, size=10.5, color=GRAY_TEXT, italic=True)

# FESEM findings
rect(s, 0.3, 3.3, 6.1, 3.9, WHITE)
txt(s, "FESEM Analysis", 0.55, 3.4, 5.6, 0.45, size=14, color=DARK_GREEN, bold=True)
txb = s.shapes.add_textbox(Inches(0.55), Inches(3.9), Inches(5.6), Inches(3.1))
tf = txb.text_frame; tf.word_wrap = True
fesem_pts = [
    ("Before Filtration:", "Clean, smooth elongated lamellar clay platelets; open unobstructed pore channels; no surface deposition — confirms fresh, unloaded state"),
    ("After Filtration:", "Heavily obscured surface with amorphous deposits, bright crystalline aggregates; pore spaces partially occluded; increased surface roughness — confirms active pollutant adsorption"),
]
for title, desc in fesem_pts:
    p = tf.add_paragraph(); run = p.add_run(); run.text = title
    run.font.size = Pt(12); run.font.bold = True; run.font.color.rgb = MED_GREEN; run.font.name = "Calibri"
    p.space_before = Pt(8)
    p2 = tf.add_paragraph(); run2 = p2.add_run(); run2.text = desc
    run2.font.size = Pt(11); run2.font.color.rgb = GRAY_TEXT; run2.font.name = "Calibri"

# FTIR findings
rect(s, 6.8, 3.3, 6.2, 3.9, DARK_GREEN)
txt(s, "FTIR Analysis", 7.05, 3.4, 5.7, 0.45, size=14, color=ACCENT, bold=True)
txb2 = s.shapes.add_textbox(Inches(7.05), Inches(3.9), Inches(5.7), Inches(3.1))
tf2 = txb2.text_frame; tf2.word_wrap = True
ftir_data = [
    ("~3300 cm⁻¹", "O–H stretching — hydroxyl/phenolic groups"),
    ("2907–3001 cm⁻¹", "C–H stretching — aliphatic hydrocarbons"),
    ("1519–1592 cm⁻¹", "C=C aromatic / N–H bending"),
    ("1042–1196 cm⁻¹", "C–O stretching — alcohols, ethers, polysaccharides"),
    ("After filtration:", "Reduced O–H peak (adsorption consumed hydroxyl groups); dominant aromatic C=C at 1612 cm⁻¹ from carbonisation; strong C–O at 1051 cm⁻¹"),
]
for peak, desc in ftir_data:
    p = tf2.add_paragraph()
    r1 = p.add_run(); r1.text = f"{peak}: "
    r1.font.size = Pt(11); r1.font.bold = True; r1.font.color.rgb = ACCENT; r1.font.name = "Calibri"
    r2 = p.add_run(); r2.text = desc
    r2.font.size = Pt(11); r2.font.color.rgb = LIGHT_GREEN; r2.font.name = "Calibri"
    p.space_before = Pt(5)

# ─────────────────────────────────────────────────────────────
# SLIDE 11 — Advantages & Limitations
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, WHITE)

rect(s, 0, 0, 13.33, 1.0, MED_GREEN)
txt(s, "ADVANTAGES & LIMITATIONS", 0.5, 0.15, 12, 0.7, size=28, color=WHITE, bold=True)

# Advantages - left
rect(s, 0.3, 1.2, 6.0, 5.8, DARK_GREEN)
txt(s, "✔  ADVANTAGES", 0.55, 1.3, 5.5, 0.5, size=15, color=ACCENT, bold=True)

advantages = [
    ("Low Cost", "Uses coconut shell agricultural waste & local clay — near-zero raw material cost"),
    ("No Energy Input", "Gravity-driven filtration system requires no electricity or pumping"),
    ("High Performance", "98.59% turbidity, 94.35% COD, 86.25% BOD removal achieved"),
    ("Decentralised", "Suitable for small/medium dairies and rural cooperative dairy farms"),
    ("Locally Appropriate", "Ideal for Kerala — abundant coconut shells from 5,000+ million nuts/year"),
    ("Sustainable", "Converts agricultural waste to functional filter; aligns with circular economy"),
    ("Simple Operation", "No skilled operators required; easy fabrication from common materials"),
    ("Scalable", "Modular disc design allows scale-up by increasing surface area or stacking"),
]
txb = s.shapes.add_textbox(Inches(0.55), Inches(1.9), Inches(5.5), Inches(4.9))
tf = txb.text_frame; tf.word_wrap = True
for title, desc in advantages:
    p = tf.add_paragraph()
    r1 = p.add_run(); r1.text = f"▶ {title}: "
    r1.font.size = Pt(11.5); r1.font.bold = True; r1.font.color.rgb = LIGHT_GREEN; r1.font.name = "Calibri"
    r2 = p.add_run(); r2.text = desc
    r2.font.size = Pt(11); r2.font.color.rgb = WHITE; r2.font.name = "Calibri"
    p.space_before = Pt(5)

# Limitations - right
rect(s, 6.8, 1.2, 6.2, 5.8, GRAY_BG)
txt(s, "✘  LIMITATIONS", 7.05, 1.3, 5.7, 0.5, size=15, color=RGBColor(0xC0,0x39,0x2B), bold=True)

limitations = [
    ("Low Flow Rate", "4.8 mL/hr — gravity operation limits throughput; scale-up required for larger volumes"),
    ("Nitrate Removal", "Only 48.43% removal — biological denitrification needed for full N removal"),
    ("Filter Saturation", "Adsorption sites get saturated over time; regeneration adds operational complexity"),
    ("pH Sensitivity", "Performance varies with influent pH, temperature, and competing ions"),
    ("Pyrolysis Complexity", "900°C controlled atmosphere pyrolysis requires specialised equipment & safety"),
    ("Clogging Risk", "Pores may gradually occlude with heavy use — lifespan/regeneration needs study"),
    ("No Pathogen Removal", "System does not target microbial/pathogen removal — additional disinfection needed"),
]
txb2 = s.shapes.add_textbox(Inches(7.05), Inches(1.9), Inches(5.7), Inches(4.9))
tf2 = txb2.text_frame; tf2.word_wrap = True
for title, desc in limitations:
    p = tf2.add_paragraph()
    r1 = p.add_run(); r1.text = f"▶ {title}: "
    r1.font.size = Pt(11.5); r1.font.bold = True; r1.font.color.rgb = RGBColor(0xC0,0x39,0x2B); r1.font.name = "Calibri"
    r2 = p.add_run(); r2.text = desc
    r2.font.size = Pt(11); r2.font.color.rgb = DARK_TEXT; r2.font.name = "Calibri"
    p.space_before = Pt(5)

# ─────────────────────────────────────────────────────────────
# SLIDE 12 — Conclusion
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, DARK_GREEN)

rect(s, 0, 0, 13.33, 1.0, RGBColor(0x0D,0x2B,0x20))
txt(s, "CONCLUSION", 0.5, 0.15, 12, 0.7, size=30, color=WHITE, bold=True)

# Summary statement
rect(s, 0.4, 1.15, 12.5, 1.2, RGBColor(0x0D,0x2B,0x20))
txt(s, "The biochar–clay composite disc filter is an effective, economical, and environmentally sustainable solution for decentralised dairy wastewater treatment.",
    0.65, 1.25, 12.0, 0.9, size=14, color=ACCENT, italic=True, align=PP_ALIGN.CENTER)

# Key conclusion points
conclusions = [
    ("Optimal Composition Found",
     "Disc C (65% clay, 16% biochar, 11% sawdust, 8% sand) achieved best performance at 86.25% BOD removal, 4.8 mL/hr flow"),
    ("Multi-Parameter Effectiveness",
     "Turbidity 98.59%, COD 94.35%, Total Hardness 90.44%, Sulphate 88.22%, Chloride 83.97% — far exceeding discharge thresholds"),
    ("Dual Mechanism Validated",
     "Clay provides structural integrity + ion exchange; Biochar provides high surface adsorption — synergistic effect confirmed by FESEM & FTIR"),
    ("Material Properties Confirmed",
     "50.70% apparent porosity and 36.89% water absorption support high filtration efficiency and long hydraulic retention time"),
    ("Practical Applicability",
     "Gravity-driven, no energy input, uses local agricultural waste — highly suitable for Kerala's rural dairy sector"),
    ("Pathway to SDG Alignment",
     "Supports UN SDG 6 (Clean Water), SDG 12 (Responsible Consumption) through circular economy approach"),
]
box_w = 4.0; box_h = 1.55
for i, (title, desc) in enumerate(conclusions):
    col = i % 3
    row = i // 3
    x = 0.4 + col*(box_w+0.1)
    y = 2.55 + row*(box_h+0.12)
    rect(s, x, y, box_w, box_h, RGBColor(0x24,0x6A,0x4A))
    txt(s, title, x+0.12, y+0.1, box_w-0.2, 0.45, size=12, color=ACCENT, bold=True)
    txb = s.shapes.add_textbox(Inches(x+0.12), Inches(y+0.55), Inches(box_w-0.2), Inches(0.9))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; run = p.add_run(); run.text = desc
    run.font.size = Pt(10.5); run.font.color.rgb = LIGHT_GREEN; run.font.name = "Calibri"

# ─────────────────────────────────────────────────────────────
# SLIDE 13 — Future Directions
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, GRAY_BG)

rect(s, 0, 0, 13.33, 1.0, MED_GREEN)
txt(s, "FUTURE DIRECTIONS", 0.5, 0.15, 12, 0.7, size=28, color=WHITE, bold=True)

future = [
    ("🔬", "Biochar Modification",
     "Chemical activation (KOH/NaOH) or metal oxide (MgO) functionalisation to enhance nitrate & phosphate removal"),
    ("⚗️", "System Scale-Up",
     "Design multi-layered or modular disc units to treat larger dairy effluent volumes at field scale"),
    ("🔄", "Regeneration Studies",
     "Evaluate thermal & chemical regeneration cycles to determine disc lifespan and long-term economic viability"),
    ("🌱", "Hybrid Integration",
     "Combine biochar-clay filtration with constructed wetlands or biological reactors for complete treatment"),
    ("🧪", "Advanced Characterisation",
     "XRD, BET surface area analysis, and zeta potential measurements for deeper mechanistic insight"),
    ("♻️", "Spent Biochar as Fertiliser",
     "Assess bioavailability of recovered phosphate & ammonium when spent biochar is land-applied — circular economy value"),
]
box_w = 4.0; box_h = 2.4
for i, (icon, title, desc) in enumerate(future):
    col = i % 3
    row = i // 3
    x = 0.3 + col*(box_w+0.2)
    y = 1.2 + row*(box_h+0.2)
    rect(s, x, y, box_w, box_h, WHITE)
    rect(s, x, y, box_w, 0.6, DARK_GREEN)
    txt(s, icon + "  " + title, x+0.1, y+0.1, box_w-0.2, 0.5, size=13, color=WHITE, bold=True)
    txb = s.shapes.add_textbox(Inches(x+0.15), Inches(y+0.75), Inches(box_w-0.25), Inches(1.5))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; run = p.add_run(); run.text = desc
    run.font.size = Pt(11.5); run.font.color.rgb = GRAY_TEXT; run.font.name = "Calibri"

# ─────────────────────────────────────────────────────────────
# SLIDE 14 — Summary Infographic / Key Numbers
# ─────────────────────────────────────────────────────────────
s = blank_slide(prs)
bg(s, WHITE)

rect(s, 0, 0, 13.33, 1.0, DARK_GREEN)
txt(s, "KEY NUMBERS AT A GLANCE", 0.5, 0.15, 12, 0.7, size=28, color=WHITE, bold=True)

numbers = [
    ("900°C", "Pyrolysis Temperature", MED_GREEN),
    ("850°C", "Disc Sintering Temp", DARK_GREEN),
    ("4.8 mL/hr", "Optimal Flow Rate", RGBColor(0x2A,0x7A,0x55)),
    ("98.59%", "Turbidity Removed", ACCENT),
    ("94.35%", "COD Removed", MED_GREEN),
    ("86.25%", "BOD Removed", DARK_GREEN),
    ("50.70%", "Apparent Porosity", RGBColor(0x2A,0x7A,0x55)),
    ("36.89%", "Water Absorption", ACCENT),
    ("10,944 μg/hr", "PLR — Salinity", MED_GREEN),
]
for i, (val, label, col) in enumerate(numbers):
    col_idx = i % 3
    row_idx = i // 3
    x = 0.35 + col_idx * 4.32
    y = 1.2 + row_idx * 2.0
    bx = x; by = y; bw = 4.1; bh = 1.75
    rect(s, bx, by, bw, bh, col)
    txt(s, val, bx, by+0.15, bw, 0.85, size=34, color=WHITE, bold=True, align=PP_ALIGN.CENTER, font="Calibri")
    txt(s, label, bx, by+1.0, bw, 0.65, size=12, color=LIGHT_GREEN, align=PP_ALIGN.CENTER)

# Footer
rect(s, 0, 7.15, 13.33, 0.35, DARK_GREEN)
txt(s, "Jaisal Johnson  |  M.Sc. Environmental Science & Technology  |  CUSAT  |  Guide: Dr. Sylas V P, ACESSD  |  April 2026",
    0.3, 7.17, 12.7, 0.3, size=9.5, color=LIGHT_GREEN, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────────
out = "/mnt/user-data/outputs/Jaisal_Biochar_Clay_Disc_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}")