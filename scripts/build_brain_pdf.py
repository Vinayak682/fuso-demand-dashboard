#!/usr/bin/env python3
"""
FUSO MEA Demand Planning — Model Brain Document
Generates: FUSO_Model_Brain.pdf
Covers: Every formula, assumption, logic, and real-world example
"""

import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.platypus import ListFlowable, ListItem

# ── PATHS ─────────────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "FUSO_Model_Brain.pdf")

# ── COLORS ────────────────────────────────────────────────────────────────────
NAVY       = colors.HexColor("#0D1B2A")
NAVY_MID   = colors.HexColor("#1B2E45")
RED        = colors.HexColor("#C0392B")
RED_LIGHT  = colors.HexColor("#E74C3C")
GOLD       = colors.HexColor("#E67E22")
GREEN      = colors.HexColor("#27AE60")
TEAL       = colors.HexColor("#1ABC9C")
BLUE       = colors.HexColor("#2980B9")
PURPLE     = colors.HexColor("#8E44AD")
GREY       = colors.HexColor("#7F8C8D")
LIGHT_BG   = colors.HexColor("#F8F9FA")
PALE_BLUE  = colors.HexColor("#EBF5FB")
PALE_GREEN = colors.HexColor("#EAFAF1")
PALE_RED   = colors.HexColor("#FDEDEC")
PALE_GOLD  = colors.HexColor("#FEF9E7")
WHITE      = colors.white
BLACK      = colors.black

W, H = A4

# ── STYLES ────────────────────────────────────────────────────────────────────
def make_styles():
    s = getSampleStyleSheet()

    def add(name, **kw):
        s.add(ParagraphStyle(name=name, **kw))

    add("ChapterTitle",
        fontName="Helvetica-Bold", fontSize=20, textColor=WHITE,
        leading=26, spaceAfter=6, backColor=NAVY,
        leftIndent=-1.5*cm, rightIndent=-1.5*cm,
        borderPad=14)

    add("SectionTitle",
        fontName="Helvetica-Bold", fontSize=13, textColor=RED,
        leading=18, spaceBefore=18, spaceAfter=8,
        borderPadding=(0,0,4,0))

    add("SubTitle",
        fontName="Helvetica-Bold", fontSize=11, textColor=NAVY,
        leading=15, spaceBefore=12, spaceAfter=6)

    add("Body",
        fontName="Helvetica", fontSize=9.5, textColor=BLACK,
        leading=15, spaceAfter=6, alignment=TA_JUSTIFY)

    add("BodyBold",
        fontName="Helvetica-Bold", fontSize=9.5, textColor=BLACK,
        leading=15, spaceAfter=4)

    add("Formula",
        fontName="Courier-Bold", fontSize=10, textColor=NAVY,
        leading=15, spaceAfter=4, backColor=PALE_BLUE,
        borderPadding=8, leftIndent=8)

    add("FormulaSmall",
        fontName="Courier", fontSize=9, textColor=NAVY_MID,
        leading=14, spaceAfter=3, backColor=PALE_BLUE,
        borderPadding=6, leftIndent=8)

    add("Example",
        fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#1a4a2e"),
        leading=14, spaceAfter=4, backColor=PALE_GREEN,
        borderPadding=8, leftIndent=8)

    add("Warning",
        fontName="Helvetica-Bold", fontSize=9, textColor=RED,
        leading=14, spaceAfter=4, backColor=PALE_RED,
        borderPadding=8, leftIndent=8)

    add("Insight",
        fontName="Helvetica-Oblique", fontSize=9.5, textColor=colors.HexColor("#5d3a00"),
        leading=14, spaceAfter=4, backColor=PALE_GOLD,
        borderPadding=8, leftIndent=8)

    add("Caption",
        fontName="Helvetica", fontSize=8, textColor=GREY,
        leading=12, spaceAfter=4, alignment=TA_CENTER)

    add("BulletItem",
        fontName="Helvetica", fontSize=9.5, textColor=BLACK,
        leading=14, spaceAfter=3, leftIndent=14, bulletIndent=0)

    add("TableHeader",
        fontName="Helvetica-Bold", fontSize=8.5, textColor=WHITE,
        leading=12, alignment=TA_CENTER)

    add("TOCEntry",
        fontName="Helvetica", fontSize=10, textColor=NAVY,
        leading=16, spaceAfter=2)

    add("TOCEntryBold",
        fontName="Helvetica-Bold", fontSize=11, textColor=NAVY,
        leading=18, spaceAfter=4, spaceBefore=8)

    add("PageNum",
        fontName="Helvetica", fontSize=8, textColor=GREY,
        leading=12, alignment=TA_CENTER)

    return s

# ── HELPERS ───────────────────────────────────────────────────────────────────
def divider(color=RED):
    return HRFlowable(width="100%", thickness=1.5, color=color, spaceAfter=8, spaceBefore=4)

def thin_divider():
    return HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#D5D8DC"),
                      spaceAfter=6, spaceBefore=6)

def chapter_box(title, subtitle, color=NAVY):
    data = [[Paragraph(f"<font size=18><b>{title}</b></font><br/>"
                       f"<font size=10 color='#95A5A6'>{subtitle}</font>",
                       ParagraphStyle("cb", fontName="Helvetica-Bold", fontSize=18,
                                      textColor=WHITE, leading=24))]]
    t = Table(data, colWidths=[W - 4*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), color),
        ("LEFTPADDING", (0,0), (-1,-1), 18),
        ("RIGHTPADDING", (0,0), (-1,-1), 18),
        ("TOPPADDING", (0,0), (-1,-1), 16),
        ("BOTTOMPADDING", (0,0), (-1,-1), 16),
        ("ROUNDEDCORNERS", (0,0), (-1,-1), 6),
    ]))
    return t

def section_box(title, color=RED):
    data = [[Paragraph(f"<b>{title}</b>",
                       ParagraphStyle("sb", fontName="Helvetica-Bold", fontSize=12,
                                      textColor=WHITE, leading=16))]]
    t = Table(data, colWidths=[W - 4*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), color),
        ("LEFTPADDING", (0,0), (-1,-1), 12),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    return t

def kv_table(rows, col_widths=None):
    """Key-value two-column table"""
    s = make_styles()
    if col_widths is None:
        col_widths = [6*cm, W - 10*cm]
    data = []
    for k, v in rows:
        data.append([
            Paragraph(f"<b>{k}</b>", s["Body"]),
            Paragraph(v, s["Body"])
        ])
    t = Table(data, colWidths=col_widths, repeatRows=0)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), PALE_BLUE),
        ("BACKGROUND", (1,0), (1,-1), WHITE),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#BDC3C7")),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    return t

def data_table(headers, rows, col_widths=None, highlight_rows=None):
    """Standard data table with navy header"""
    s = make_styles()
    header_row = [Paragraph(f"<b>{h}</b>", s["TableHeader"]) for h in headers]
    all_rows = [header_row]
    for i, row in enumerate(rows):
        cells = [Paragraph(str(c), ParagraphStyle("tc", fontName="Helvetica",
                           fontSize=8.5, leading=12, alignment=TA_LEFT)) for c in row]
        all_rows.append(cells)

    if col_widths is None:
        col_widths = [(W - 4*cm) / len(headers)] * len(headers)

    t = Table(all_rows, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0,0), (-1,0), NAVY),
        ("TEXTCOLOR", (0,0), (-1,0), WHITE),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#BDC3C7")),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT_BG]),
    ]
    if highlight_rows:
        for ri, color in highlight_rows:
            style.append(("BACKGROUND", (0, ri+1), (-1, ri+1), color))
    t.setStyle(TableStyle(style))
    return t

# ── PAGE TEMPLATE ─────────────────────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    # Header bar
    canvas.setFillColor(NAVY)
    canvas.rect(0, H - 1*cm, W, 1*cm, fill=1, stroke=0)
    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawString(1.5*cm, H - 0.65*cm, "FUSO MEA DEMAND PLANNING — MODEL BRAIN DOCUMENT")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(W - 1.5*cm, H - 0.65*cm, "DSV Logistics × Mitsubishi FUSO | Jafza Free Zone, Dubai")
    # Footer
    canvas.setFillColor(GREY)
    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(1.5*cm, 0.6*cm, "CONFIDENTIAL — Internal Reference Document — Vinayak Bhadani")
    canvas.drawRightString(W - 1.5*cm, 0.6*cm, f"Page {doc.page}")
    canvas.setStrokeColor(RED)
    canvas.setLineWidth(1.5)
    canvas.line(1.5*cm, 0.9*cm, W - 1.5*cm, 0.9*cm)
    canvas.restoreState()

# ── BUILD ─────────────────────────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(
        OUT,
        pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm,
        title="FUSO MEA Demand Planning — Model Brain",
        author="Vinayak Bhadani"
    )
    s = make_styles()
    story = []

    def P(text, style="Body"):
        return Paragraph(text, s[style])

    def SP(n=6):
        return Spacer(1, n)

    def bullet(text):
        return Paragraph(f"• {text}", s["BulletItem"])

    # ══════════════════════════════════════════════════════════
    # COVER PAGE
    # ══════════════════════════════════════════════════════════
    story.append(SP(40))
    cover = Table([[
        Paragraph("<font size=28 color='#0D1B2A'><b>FUSO MEA</b></font><br/>"
                  "<font size=18 color='#C0392B'><b>Demand Planning</b></font><br/>"
                  "<font size=18 color='#C0392B'><b>Model Brain Document</b></font>",
                  ParagraphStyle("cv", fontName="Helvetica-Bold", fontSize=28,
                                 textColor=NAVY, leading=36, alignment=TA_CENTER))
    ]], colWidths=[W - 3*cm])
    cover.setStyle(TableStyle([
        ("LEFTPADDING", (0,0), (-1,-1), 20),
        ("RIGHTPADDING", (0,0), (-1,-1), 20),
        ("TOPPADDING", (0,0), (-1,-1), 30),
        ("BOTTOMPADDING", (0,0), (-1,-1), 30),
        ("BOX", (0,0), (-1,-1), 2, NAVY),
        ("LINEBELOW", (0,0), (-1,0), 4, RED),
    ]))
    story.append(cover)
    story.append(SP(16))

    story.append(P(
        "<b>The complete formula logic, business assumptions, real-world examples, "
        "and decision rationale behind every calculation in the FUSO MEA "
        "Advanced Demand Planning Model.</b>",
        "SubTitle"
    ))
    story.append(SP(8))
    meta_data = [
        ["Model File", "FUSO_Advanced_Model_v2.xlsx"],
        ["Scope", "58 MEA Markets · Jafza Free Zone Hub · Japan / Chennai / GPC Halberstadt"],
        ["Portfolio", "AED 38M · 50 Active SKUs · 3 Vehicle Lines"],
        ["Prepared by", "Vinayak Bhadani — DSV × FUSO Final Round Interview"],
        ["Date", "March 2026"],
    ]
    story.append(kv_table(meta_data, col_widths=[4.5*cm, W - 9*cm]))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # TABLE OF CONTENTS
    # ══════════════════════════════════════════════════════════
    story.append(section_box("TABLE OF CONTENTS", NAVY))
    story.append(SP(10))
    toc = [
        ("1", "ABC ANALYSIS — Annual Value Classification"),
        ("2", "XYZ ANALYSIS — Demand Variability & Coefficient of Variation"),
        ("3", "REPLENISHMENT STRATEGY — VLOOKUP Matrix Logic"),
        ("4", "SMOB FLAG & DISPOSITION — Slow Moving & Obsolete Stock"),
        ("5", "SBA ENGINE — Syntetos-Boylan Approximation"),
        ("6", "SAFETY STOCK v2 — Standard vs Enhanced Formula (3 Scenarios)"),
        ("7", "SERVICE LEVEL OPTIMIZER — NORM.S.INV & Working Capital"),
        ("8", "eCanter CAUSAL FORECASTING — VOR Risk Triggers"),
        ("9", "GPC HALBERSTADT BRIDGE STOCK — Transition Calculator"),
    ]
    for num, title in toc:
        story.append(P(f"<b>Chapter {num}</b> &nbsp;&nbsp; {title}", "TOCEntry"))
        story.append(thin_divider())

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 1 — ABC ANALYSIS
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 1", "ABC Analysis — Annual Value Classification", NAVY))
    story.append(SP(14))

    story.append(section_box("1.1  What is ABC Analysis?"))
    story.append(SP(8))
    story.append(P(
        "ABC Analysis is the foundational classification framework for spare parts inventory. "
        "It is rooted in the <b>Pareto Principle (80/20 rule)</b> — the observation that in any "
        "inventory, a small number of parts account for the vast majority of total value. "
        "By classifying parts into three tiers (A, B, C), a demand planner can concentrate "
        "management attention and working capital where it matters most, without over-investing "
        "in low-value items."
    ))
    story.append(SP(6))
    story.append(P(
        "In the FUSO MEA context, managing 58 markets with a AED 38M parts portfolio from "
        "three supply origins (Japan 45-day LT, Chennai 21-day LT, Germany 30-day LT), "
        "ABC classification is the first gate — it determines service level targets, "
        "safety stock Z-scores, review frequency, and replenishment policies for every SKU."
    ))

    story.append(SP(10))
    story.append(section_box("1.2  Step 1 — Cumulative Value %  |  Formula: =G3/SUM($G$3:$G$52)"))
    story.append(SP(8))
    story.append(P(
        "<b>Column G (Annual Value AED)</b> = Unit Cost × 12-Month Demand. "
        "Column H then computes what <i>share</i> of the total portfolio value this SKU represents "
        "<i>cumulatively</i> — i.e., if you sorted all 50 SKUs from most expensive to least, "
        "what running percentage of total value have you reached?"
    ))
    story.append(SP(6))
    story.append(P("Formula breakdown:", "BodyBold"))
    story.append(P("=G3 / SUM($G$3:$G$52)", "Formula"))
    story.append(SP(4))
    story.append(kv_table([
        ("G3", "Annual Value of this specific SKU (Cost × Annual Demand)"),
        ("SUM($G$3:$G$52)", "Total Annual Value of ALL 50 SKUs combined (absolute reference — does not shift when copied down)"),
        ("Result", "A decimal between 0 and 1. When sorted descending by value, this accumulates toward 1.0 (100%)"),
    ], col_widths=[4*cm, W - 8.5*cm]))
    story.append(SP(8))
    story.append(P(
        "<b>Why cumulative and not individual %?</b> Because ABC is about concentration, not "
        "individual weight. You want to know: 'At what point in the sorted list have I covered "
        "70% of total spend?' That is the A-boundary. Individual percentages do not tell you this."
    ))
    story.append(SP(8))

    story.append(P("Real-World Example:", "SubTitle"))
    story.append(P(
        "Imagine 5 SKUs with annual values: AED 50,000 / 30,000 / 10,000 / 6,000 / 4,000. "
        "Total = AED 100,000. Sorted descending and cumulated:",
        "Body"
    ))
    story.append(SP(4))
    story.append(data_table(
        ["SKU", "Annual Value (AED)", "Individual %", "Cumul %", "Observation"],
        [
            ["Brake Pad", "50,000", "50%", "50%", "Halfway with just 1 SKU"],
            ["Alternator", "30,000", "30%", "80%", "Crossed 70% boundary at 2 SKUs"],
            ["Fuel Filter", "10,000", "10%", "90%", "Crossed 90% boundary"],
            ["Oil Filter", "6,000", "6%", "96%", "Low value"],
            ["Gasket", "4,000", "4%", "100%", "Tail end"],
        ],
        col_widths=[3.5*cm, 3.5*cm, 2.5*cm, 2.5*cm, W - 16.5*cm],
        highlight_rows=[(0, PALE_GREEN), (1, PALE_GREEN)]
    ))
    story.append(SP(4))
    story.append(P(
        "✦ Brake Pad and Alternator (2 out of 5 = 40% of SKUs) account for AED 80,000 "
        "(80% of total value). This is the ABC Pareto effect. In FUSO's 50-SKU model, "
        "21 SKUs (42%) are A-class and account for ~70% of AED 38M = AED 26.6M.",
        "Example"
    ))

    story.append(SP(10))
    story.append(section_box("1.3  Step 2 — ABC Classification  |  =IF(H3<=0.7,\"A\",IF(H3<=0.9,\"B\",\"C\"))"))
    story.append(SP(8))
    story.append(P(
        "Once cumulative value % is computed (Column H), this IF formula assigns the class. "
        "The thresholds 0.7 / 0.9 / 1.0 are the <b>standard industry convention</b> for "
        "industrial spare parts, not arbitrary numbers. Here is the business logic behind each:"
    ))
    story.append(SP(6))
    story.append(data_table(
        ["Class", "Cumul Value Threshold", "% of Total Value", "Typical SKU Count", "Business Meaning"],
        [
            ["A", "<= 70%  (H3 <= 0.7)", "~70% of value", "15-25% of SKUs", "Critical parts — highest financial impact if stocked out"],
            ["B", "<= 90%  (H3 <= 0.9)", "Next 20% of value", "~25-35% of SKUs", "Important but not critical — moderate management"],
            ["C", "> 90%   (H3 > 0.9)", "Bottom 10% of value", "40-55% of SKUs", "Low-value parts — minimal management needed"],
        ],
        col_widths=[1.5*cm, 3.5*cm, 3*cm, 3*cm, W - 15.5*cm],
        highlight_rows=[(0, PALE_GREEN), (1, PALE_BLUE), (2, LIGHT_BG)]
    ))
    story.append(SP(6))
    story.append(P(
        "<b>Why 70/90 specifically?</b> This is the APICS/CSCMP standard for industrial parts. "
        "Some companies use 80/95 for retail FMCG. FUSO uses 70/90 because spare parts have "
        "extremely skewed value distributions — engines (AED 18,000) vs gaskets (AED 45). "
        "A tighter 70% A-boundary ensures the most expensive assemblies get the highest Z-scores "
        "and tightest safety stock management.", "Body"
    ))
    story.append(SP(6))
    story.append(P(
        "FUSO MEA Result: A=21 SKUs, B=13 SKUs, C=16 SKUs. "
        "A-class parts include Battery Cell Module (AED 18,000 × 6 = AED 108,000/yr), "
        "On-Board Charger (AED 9,500), eCanter ECU (AED 8,500). "
        "C-class includes Oil Filters (AED 45 × 1,800 = AED 81,000/yr — high volume, low unit cost).",
        "Example"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 2 — XYZ ANALYSIS
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 2", "XYZ Analysis — Demand Variability & Coefficient of Variation", NAVY))
    story.append(SP(14))

    story.append(section_box("2.1  What is XYZ Analysis?"))
    story.append(SP(8))
    story.append(P(
        "ABC tells you WHAT a part is worth. XYZ tells you HOW PREDICTABLE its demand is. "
        "This is critical for spare parts because two parts can have identical annual spend "
        "but completely different demand patterns. "
        "An Oil Filter with steady 150 units/month demand (predictable) requires a very different "
        "forecasting method than a Turbocharger with 0,0,0,48,0,0,0,0,48 units "
        "(intermittent — sporadic)."
    ))
    story.append(SP(6))
    story.append(P(
        "XYZ uses the <b>Coefficient of Variation (CV)</b> — a dimensionless ratio that normalises "
        "demand variability across parts of different volumes. Without normalisation, a high-volume "
        "part would always appear more variable simply because its absolute standard deviation "
        "is larger."
    ))

    story.append(SP(10))
    story.append(section_box("2.2  Coefficient of Variation  |  =IF(K3=0, 9999, J3/K3)"))
    story.append(SP(8))
    story.append(P("Formula:", "BodyBold"))
    story.append(P("CV = Standard Deviation (monthly) / Mean Monthly Demand", "Formula"))
    story.append(P("Excel: =IF(K3=0, 9999, J3/K3)    where K3=Mean Monthly, J3=Std Dev", "FormulaSmall"))
    story.append(SP(6))
    story.append(kv_table([
        ("J3 — Std Dev", "Monthly demand standard deviation (how much demand fluctuates around the mean)"),
        ("K3 — Mean", "Average monthly demand = Annual Demand / 12"),
        ("IF(K3=0, 9999)", "Protection against division by zero. If a part has ZERO mean demand it is already a dead SKU — assign CV=9999 (extreme Z-class)"),
        ("Result", "A pure ratio. CV=0.05 means demand is very stable. CV=1.5 means demand is highly erratic"),
    ], col_widths=[4.5*cm, W - 9*cm]))
    story.append(SP(8))

    story.append(section_box("2.3  XYZ Classification  |  =IF(L3<0.1,\"X\",IF(L3<0.3,\"Y\",\"Z\"))"))
    story.append(SP(8))
    story.append(data_table(
        ["Class", "CV Range", "Demand Pattern", "FUSO SKU Count", "Forecasting Implication"],
        [
            ["X", "CV < 0.10", "Highly stable, predictable", "0 in model", "Simple Moving Average / WMA works well"],
            ["Y", "0.10 <= CV < 0.30", "Moderate variability, some trend", "15 SKUs", "Exponential Smoothing, seasonal adjustment"],
            ["Z", "CV >= 0.30", "Intermittent / sporadic", "35 SKUs", "SBA (Syntetos-Boylan) or Croston's Method required"],
        ],
        col_widths=[1.5*cm, 3*cm, 4*cm, 3*cm, W - 16*cm],
        highlight_rows=[(2, PALE_RED)]
    ))
    story.append(SP(6))
    story.append(P(
        "<b>Key finding for FUSO:</b> 35 out of 50 SKUs (70%) are Z-class (CV > 0.30). "
        "This is <i>expected</i> for a commercial vehicle spare parts portfolio. "
        "Engines fail sporadically. Turbochargers are replaced infrequently. "
        "This is precisely why SAP APO's standard moving average baseline (designed for "
        "retail FMCG with stable demand) achieves only 70% accuracy on FUSO's portfolio. "
        "Z-class parts need intermittent forecasting methods — SBA/Croston's — not averages.",
        "Warning"
    ))
    story.append(SP(8))

    story.append(P("Real-World Example — CV Calculation:", "SubTitle"))
    story.append(data_table(
        ["Part", "Monthly Demand Pattern (12 months)", "Std Dev", "Mean", "CV", "Class"],
        [
            ["Oil Filter", "145,150,148,155,142,149,151,147,153,146,150,164", "5.8", "150", "0.039", "X — Stable"],
            ["Shock Absorber", "22,18,35,8,41,12,28,15,33,10,25,20", "10.4", "22.3", "0.466", "Z — Intermittent"],
            ["Turbocharger", "0,0,0,4,0,0,0,0,4,0,0,0", "1.7", "0.67", "2.56", "Z — Very Sparse"],
        ],
        col_widths=[3*cm, 5.5*cm, 1.8*cm, 1.6*cm, 1.6*cm, W - 17*cm],
        highlight_rows=[(0, PALE_GREEN), (1, PALE_GOLD), (2, PALE_RED)]
    ))
    story.append(SP(4))
    story.append(P(
        "The Oil Filter barely deviates — SAP APO handles it fine. "
        "The Turbocharger has 10 months of zero demand then 2 bursts of 4 units — "
        "a moving average would forecast ~0.67/month and always be wrong. "
        "SBA separates the 'when demand occurs' from 'how much' — this is the breakthrough.",
        "Example"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 3 — REPLENISHMENT STRATEGY
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 3", "Replenishment Strategy — VLOOKUP Matrix Logic", NAVY))
    story.append(SP(14))

    story.append(section_box("3.1  The Combined ABC-XYZ Class"))
    story.append(SP(8))
    story.append(P(
        "Once ABC (value) and XYZ (variability) are known, they are concatenated into a "
        "<b>Combined Class</b> (Column N): AX, AY, AZ, BX, BY, BZ, CX, CY, CZ. "
        "This 9-cell matrix is the master decision framework for all inventory policy. "
        "Each cell has a distinct replenishment strategy because the combination of value "
        "and variability creates fundamentally different risk profiles."
    ))

    story.append(SP(10))
    story.append(section_box("3.2  VLOOKUP Strategy Assignment  |  =IFERROR(VLOOKUP(N3,$T$2:$U$10,2,FALSE),\"Manual\")"))
    story.append(SP(8))
    story.append(P("Formula breakdown:", "BodyBold"))
    story.append(kv_table([
        ("N3", "Combined Class (e.g., 'AX', 'CZ')"),
        ("$T$2:$U$10", "Strategy lookup table — 9 rows (one per combined class), 2 columns (class, strategy). Absolute reference."),
        ("2", "Return column 2 from the lookup table — the strategy description"),
        ("FALSE", "Exact match only — 'AX' must match exactly 'AX', no approximation"),
        ("IFERROR(...,'Manual')", "If the combined class is not in the table (e.g., data entry error), return 'Manual' instead of #N/A error"),
    ], col_widths=[4.5*cm, W - 9*cm]))
    story.append(SP(8))

    story.append(P("The 9-cell Strategy Matrix:", "SubTitle"))
    story.append(data_table(
        ["Combined", "Replenishment Strategy", "Why This Strategy", "FUSO SKUs"],
        [
            ["AX", "Min-Max / Continuous Review", "High value + stable = can optimise tightly. Never stockout of a AED 3,800 part.", "0"],
            ["AY", "Periodic Review (R,S)", "High value + moderate variability. Review every R days, order up to S. Catches trends.", "5"],
            ["AZ", "Demand Sensing + Buffer", "High value + intermittent. Buffer stock + real-time triggers. Cannot rely on forecast alone.", "16"],
            ["BX", "Kanban / Fixed-Period", "Medium value + stable = textbook Kanban. Visual signal, no complex calculation.", "0"],
            ["BY", "Periodic Review Quarterly", "Medium value + variable. Quarterly review balances cost of holding vs review effort.", "6"],
            ["BZ", "Order-on-Demand + Emergency", "Medium value + sporadic. Don't hold unless triggered. Keep emergency supplier contact.", "7"],
            ["CX", "Min-Max Light / Annual Review", "Low value + stable. Set a generous min-max once a year and leave it.", "0"],
            ["CY", "Periodic Review Annual", "Low value + moderate. Annual review, bulk buy if economics allow.", "4"],
            ["CZ", "SMOB Candidate / Order-on-Demand", "Low value + sporadic = classic obsolescence risk. Do NOT replenish proactively.", "12"],
        ],
        col_widths=[2*cm, 4.5*cm, 6*cm, W - 17*cm],
        highlight_rows=[(0, PALE_GREEN), (2, PALE_GOLD), (8, PALE_RED)]
    ))
    story.append(SP(6))
    story.append(P(
        "FUSO MEA finding: 16 AZ parts + 12 CZ parts = 28 SKUs (56%) in high-risk categories. "
        "AZ parts are high-value with unpredictable demand (Turbocharger AED 3,800, Injector AED 1,200). "
        "CZ parts are prime SMOB candidates — low value, sporadic demand, high obsolescence risk.",
        "Insight"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 4 — SMOB
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 4", "SMOB Flag & Disposition — Slow Moving & Obsolete Stock", NAVY))
    story.append(SP(14))

    story.append(section_box("4.1  What is SMOB and Why Does It Matter?"))
    story.append(SP(8))
    story.append(P(
        "SMOB = <b>Slow Moving and OBsolete</b> stock. It is inventory that has not moved "
        "(been consumed or sold) for an extended period, or inventory for discontinued "
        "vehicle models with no remaining active fleet. "
        "SMOB is a silent destroyer of working capital — it ties up AED in warehouse space, "
        "insurance, and depreciation while generating zero revenue."
    ))
    story.append(SP(6))
    story.append(P(
        "FUSO MEA current state: <b>AED 6.84M locked in SMOB (18% of AED 38M portfolio)</b>. "
        "Industry benchmark is 3-5%. Every AED locked in SMOB is capital that cannot be used "
        "to stock fast-moving A-class parts, reducing fill rate. "
        "The 3-year target is to reduce SMOB below 5% — recovering approximately AED 4.9M.",
        "Warning"
    ))

    story.append(SP(10))
    story.append(section_box("4.2  SMOB Flag  |  =IF(AND(M3=\"Z\",I3=\"C\"),\"SMOB RISK\",\"OK\")"))
    story.append(SP(8))
    story.append(P("Formula logic:", "BodyBold"))
    story.append(P('=IF(AND(M3="Z", I3="C"), "SMOB RISK", "OK")', "Formula"))
    story.append(SP(6))
    story.append(kv_table([
        ("M3 = Z", "XYZ class is Z = demand is highly intermittent (CV > 0.30). The part barely moves."),
        ("I3 = C", "ABC class is C = the part is in the bottom 10% of annual value. Low financial impact."),
        ("AND()", "BOTH conditions must be true simultaneously. A high-value sporadic part (AZ) is NOT flagged — it needs buffer stock, not liquidation."),
        ("'SMOB RISK'", "The combined CZ profile is the classic obsolescence signature: low value + sporadic demand"),
        ("'OK'", "Any other combination — even AZ (high value, sporadic) — is managed differently"),
    ], col_widths=[4*cm, W - 8.5*cm]))
    story.append(SP(6))
    story.append(P(
        "<b>Business rationale:</b> A CZ part costs little per unit AND barely moves. "
        "Every month it sits in the warehouse costs holding cost (~2.5% of value/month for Dubai climate-controlled storage). "
        "After 6 months, you have paid 15% of the part's value just in holding costs. "
        "After 12 months, 30%. At some point, scrapping it or bundling it is cheaper "
        "than continuing to store it."
    ))

    story.append(SP(10))
    story.append(section_box("4.3  Months Zero Demand — What If Stock Was Missing (Not Demand)?"))
    story.append(SP(8))
    story.append(P(
        "Column Q counts the number of months in the last 12 with zero consumption recorded. "
        "This is manually populated from the SAP transaction history (MMBE / MB52 reports)."
    ))
    story.append(SP(6))
    story.append(P(
        "<b>The key limitation you raise:</b> Zero demand months can mean two very different things:",
        "BodyBold"
    ))
    story.append(data_table(
        ["Reason for Zero", "What it looks like in SAP", "Correct interpretation", "Action"],
        [
            ["True zero demand", "Part in stock, no GI (goods issue) recorded", "No customer need — genuine slow mover", "Proceed with SMOB flagging"],
            ["Stockout (hidden)", "Part shows 0 stock (ZS in MRP), no GI possible", "Demand suppressed by unavailability", "Do NOT flag as SMOB — investigate fill rate"],
            ["Seasonal pattern", "0 demand for 4 months, then 8 in next month", "Seasonal vehicle model", "Use seasonal index, not SMOB flag"],
        ],
        col_widths=[3.5*cm, 4.5*cm, 4*cm, W - 16.5*cm],
        highlight_rows=[(1, PALE_RED)]
    ))
    story.append(SP(6))
    story.append(P(
        "In this model, the assumption is that JAFZA warehouse records are accurate and "
        "stock availability can be verified. To properly distinguish, a demand planner must "
        "cross-reference SAP MRP stock reports with the consumption history. "
        "If a part shows 0 stock AND 0 demand, it is a stockout scenario — not SMOB. "
        "This is a known limitation of purely formula-based models.",
        "Insight"
    ))

    story.append(SP(10))
    story.append(section_box("4.4  Disposition Strategy  |  =IF(AND(P3=\"SMOB RISK\",Q3>6),\"LIQUIDATE\",IF(Q3>3,\"REVIEW\",\"ACTIVE\"))"))
    story.append(SP(8))
    story.append(P("Formula logic:", "BodyBold"))
    story.append(P('=IF(AND(P3="SMOB RISK", Q3>6), "LIQUIDATE", IF(Q3>3, "REVIEW", "ACTIVE"))', "Formula"))
    story.append(SP(6))
    story.append(data_table(
        ["Output", "Conditions", "Threshold Logic", "Business Action"],
        [
            ["LIQUIDATE", "SMOB RISK flag AND >6 months no movement", ">6 months = holding cost exceeds expected recovery at normal price. Discount or scrap now.", "Bundle (20% off), Rotate to other markets, or Write-off"],
            ["REVIEW", "Any part with >3 months no movement (regardless of class)", "3 months is the early warning trigger. Not yet SMOB but trending.", "Investigate: is it seasonal? Model discontinuation? Supply gap?"],
            ["ACTIVE", "Less than 3 months with zero demand", "Normal sporadic pattern within Z-class tolerance", "Continue monitoring, no action"],
        ],
        col_widths=[2.5*cm, 4.5*cm, 4.5*cm, W - 16*cm],
        highlight_rows=[(0, PALE_RED), (1, PALE_GOLD)]
    ))
    story.append(SP(6))
    story.append(P(
        "The 3-step FUSO SMOB Disposition Framework: "
        "(1) BUNDLE — pair with fast-moving service kits at 20% discount; "
        "(2) ROTATE — offer to 58 MEA market distributors via inter-market swap (amnesty stock programme); "
        "(3) SCRAP / WRITE-OFF — when holding cost exceeds recovery value or the vehicle model "
        "is discontinued with no fleet remaining.",
        "Example"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 5 — SBA ENGINE
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 5", "SBA Engine — Syntetos-Boylan Approximation", NAVY))
    story.append(SP(14))

    story.append(section_box("5.1  Why SBA? The Problem with Standard Forecasting on Sparse Demand"))
    story.append(SP(8))
    story.append(P(
        "Standard forecasting methods (Moving Average, Exponential Smoothing, SAP APO's "
        "baseline) were designed for <b>regular, continuous demand</b> like retail FMCG. "
        "They assume demand occurs every period. When applied to intermittent spare parts "
        "(where demand is 0,0,0,40,0,0 across 6 months), they produce systematically "
        "incorrect forecasts."
    ))
    story.append(SP(6))
    story.append(P(
        "The specific problem: <b>Croston's Method</b> (1972) was the first attempt to fix this. "
        "It separates intermittent demand into two components: <i>demand size</i> (how much when "
        "a demand event occurs) and <i>demand interval</i> (how long between events). "
        "However, Croston's method was later proven by Syntetos & Boylan (2001) to have a "
        "systematic <b>upward bias</b> — it consistently over-forecasts by a factor of approximately "
        "α/2 (where α is the smoothing constant). "
        "SBA applies a simple bias correction: multiply by (1 - α/2)."
    ))

    story.append(SP(10))
    story.append(section_box("5.2  SBA Core Formula"))
    story.append(SP(8))
    story.append(P("The SBA (Syntetos-Boylan Approximation) forecast:", "BodyBold"))
    story.append(P("F_SBA = (1 - alpha/2) × (Z_hat / P_hat)", "Formula"))
    story.append(SP(4))
    story.append(kv_table([
        ("Z_hat", "Exponentially smoothed estimate of DEMAND SIZE (how many units when a demand event occurs)"),
        ("P_hat", "Exponentially smoothed estimate of DEMAND INTERVAL (how many periods between demand events)"),
        ("alpha", "Smoothing constant (0 to 1). In the model: alpha = 0.1 (set low for intermittent — overreacts otherwise)"),
        ("(1 - alpha/2)", "The bias correction factor. With alpha=0.1: factor = (1 - 0.05) = 0.95. This reduces Croston's overestimate."),
        ("Final result", "Demand rate per period — expected units demanded per review period"),
    ], col_widths=[4*cm, W - 8.5*cm]))

    story.append(SP(10))
    story.append(section_box("5.3  The Three SBA Inputs: Alpha, Service Level, Review Period"))
    story.append(SP(8))

    story.append(P("<b>Alpha (Smoothing Parameter) = 0.1</b>", "SubTitle"))
    story.append(P(
        "Alpha controls how quickly the forecast responds to new demand observations. "
        "Range: 0 (never update) to 1 (100% weight on latest observation only)."
    ))
    story.append(data_table(
        ["Alpha Value", "Behaviour", "Risk", "Best For"],
        [
            ["0.05 - 0.10", "Very slow to respond. Smooths out noise heavily.", "Misses genuine demand step-changes", "Very stable intermittent parts (FUSO spare parts)"],
            ["0.15 - 0.25", "Moderate response speed", "Balanced", "Spare parts with some trend"],
            ["0.30 - 0.50", "Fast response to new data", "Over-reacts to noise — one demand spike inflates forecast", "High-frequency consumer goods"],
        ],
        col_widths=[2.5*cm, 4*cm, 4*cm, W - 15*cm],
        highlight_rows=[(0, PALE_GREEN)]
    ))
    story.append(SP(6))
    story.append(P(
        "FUSO uses alpha=0.1 because FUSO spare parts have long inter-arrival times (months). "
        "A single VOR emergency order for 40 water pumps should NOT permanently inflate the "
        "long-run forecast to 40/month. Low alpha damps this noise appropriately.",
        "Insight"
    ))

    story.append(SP(8))
    story.append(P("<b>Service Level = 0.95 (95%)</b>", "SubTitle"))
    story.append(P(
        "Service level is the <b>probability of not experiencing a stockout</b> during any given "
        "replenishment cycle. Setting it to 0.95 means: we accept a 5% chance of running out "
        "of stock before the next replenishment arrives. "
        "Service level drives the Z-score used in safety stock — higher service level requires "
        "more safety stock (covered in detail in Chapter 7)."
    ))

    story.append(SP(8))
    story.append(P("<b>Review Period = 7 days</b>", "SubTitle"))
    story.append(P(
        "The review period is how frequently you check stock levels and potentially place an order. "
        "7 days (weekly review) is used in the SBA sheet as the calculation base period. "
        "This means the SBA output is a weekly demand rate. For the safety stock calculation "
        "(Chapter 6), the review period shifts to 30 days (monthly S&OP cycle)."
    ))
    story.append(SP(6))
    story.append(kv_table([
        ("Daily review (R=1)", "Maximum responsiveness. High ordering cost. Used for A-class critical parts in some companies."),
        ("Weekly review (R=7)", "SBA engine default. Balances responsiveness with procurement effort."),
        ("Monthly review (R=30)", "Used for SS calculation and S&OP. Standard for MEA distribution networks."),
        ("Quarterly (R=90)", "C-class, low-value parts. Annual review sometimes used for CZ SMOB candidates."),
    ], col_widths=[4.5*cm, W - 9*cm]))

    story.append(SP(10))
    story.append(section_box("5.4  Why Only 4 Parts in the SBA Sheet?"))
    story.append(SP(8))
    story.append(P(
        "The SBA sheet contains 4 representative case study parts — not all 50. This is a "
        "<b>deliberate design choice</b> for clarity and interview demonstration purposes:"
    ))
    for pt in [
        "<b>FUSO-WPM-007 Water Pump (AY)</b> — the headline comparison: SBA vs SAP APO on [0,0,0,40,0,0] demand pattern",
        "<b>FUSO-TRB-021 Turbocharger (AZ)</b> — high-value sporadic: demonstrates why buffer stock + SBA is needed",
        "<b>FUSO-BAT-032 Battery Module (AZ)</b> — eCanter EV part: causal trigger overrides SBA entirely",
        "<b>FUSO-INJ-013 Injector Nozzle (AZ)</b> — shows fleet age modifier interaction with SBA base forecast",
    ]:
        story.append(bullet(pt))
    story.append(SP(6))
    story.append(P(
        "In a production deployment, ALL Z-class and Y-class parts would run through the SBA "
        "engine (45 out of 50 SKUs). The 4-part model is the proof-of-concept that demonstrates "
        "the methodology to the interview panel.",
        "Insight"
    ))

    story.append(SP(10))
    story.append(section_box("5.5  Fleet Age Modifier Lookup Table"))
    story.append(SP(8))
    story.append(P(
        "The Fleet Age Modifier is a <b>demand amplifier</b> applied to the SBA base forecast "
        "based on the age distribution of the operating fleet in each market. "
        "As vehicles age, their failure rates increase following the <b>Weibull distribution</b> "
        "(a standard reliability engineering model). Older vehicles need more spare parts."
    ))
    story.append(SP(6))
    story.append(data_table(
        ["Fleet Age", "Modifier", "Business Rationale", "Real Example"],
        [
            ["0-2 years", "0.8×", "New vehicles under warranty. OEM covers failures. Dealer demand suppressed by 20%.", "New Saudi fleet (2024 Canter FE): apply 0.8× to base forecast"],
            ["3-4 years", "1.0×", "Post-warranty. Normal operating wear. Base forecast is correct.", "UAE fleet (2021-22 models): no adjustment"],
            ["5-6 years", "1.2×", "Accelerating wear. Brake pads, filters, belts need 20% more frequent replacement.", "Kuwait fleet (2018-19 models): multiply forecast by 1.2"],
            ["7+ years", "1.5×", "Heavy wear phase. Clutches, shock absorbers, water pumps fail at 50% higher rate.", "Egypt fleet (2016-17 models): multiply forecast by 1.5"],
        ],
        col_widths=[2.5*cm, 2*cm, 5.5*cm, W - 14.5*cm],
        highlight_rows=[(3, PALE_RED)]
    ))
    story.append(SP(6))
    story.append(P(
        "Example: Water Pump SBA base forecast = 4 units/month. "
        "Egypt market fleet is 8 years old. Modifier = 1.5. "
        "Egypt-adjusted forecast = 4 × 1.5 = 6 units/month. "
        "This explains why a single 'global' forecast fails — the 58 MEA markets have "
        "vastly different fleet age profiles, from brand-new GCC fleets to aging East African fleets.",
        "Example"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 6 — SAFETY STOCK
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 6", "Safety Stock v2 — Standard vs Enhanced Formula", NAVY))
    story.append(SP(14))

    story.append(section_box("6.1  What is Safety Stock and Why Does FUSO Need It?"))
    story.append(SP(8))
    story.append(P(
        "Safety stock (SS) is buffer inventory held <i>above</i> the average expected demand "
        "during lead time. It exists to protect against two sources of uncertainty: "
        "(1) demand variability — customers needing more parts than expected, and "
        "(2) lead time variability — suppliers delivering later than expected. "
        "Without safety stock, any upward demand spike or supplier delay causes a stockout."
    ))
    story.append(SP(6))
    story.append(P(
        "FUSO operates 58 markets from a single Jafza hub. A stockout at the hub means ALL "
        "58 distributors are affected simultaneously. A VOR (Vehicle Off Road) situation "
        "in Saudi Arabia, UAE, or East Africa cannot wait 45 days for the next Japan shipment. "
        "Safety stock is the insurance policy that bridges that gap."
    ))

    story.append(SP(10))
    story.append(section_box("6.2  Block A — Standard Safety Stock Formula"))
    story.append(SP(8))
    story.append(P("Standard Formula:", "BodyBold"))
    story.append(P("SS_standard = Z × sigma_demand × SQRT(Lead_Time + Review_Period)", "Formula"))
    story.append(SP(6))
    story.append(kv_table([
        ("Z", "Z-score from service level (see Chapter 7). A=2.05 (97.9% SL), B=1.65 (95%), C=1.28 (90%)"),
        ("sigma_demand", "Standard deviation of DAILY demand (monthly std dev ÷ 30)"),
        ("SQRT(LT + R)", "Square root of Lead Time + Review Period in days. Time horizon over which uncertainty accumulates."),
        ("Assumption", "This formula assumes lead time is CONSTANT (no lead time variability). Valid for Japan supply (sigma_LT = 2 days only)."),
    ], col_widths=[3.5*cm, W - 8*cm]))

    story.append(SP(10))
    story.append(section_box("6.3  Block B — Enhanced Safety Stock (GPC Halberstadt Case)"))
    story.append(SP(8))
    story.append(P("Enhanced Formula — includes lead time variability:", "BodyBold"))
    story.append(P("SS_enhanced = Z × SQRT( LT × sigma_d^2 + D_avg^2 × sigma_LT^2 )", "Formula"))
    story.append(SP(6))
    story.append(kv_table([
        ("LT", "Mean lead time in days"),
        ("sigma_d", "Standard deviation of DAILY demand"),
        ("D_avg", "Average DAILY demand"),
        ("sigma_LT", "Standard deviation of LEAD TIME in days — the critical addition vs Block A"),
        ("LT × sigma_d^2", "Uncertainty from demand variability DURING the lead time period"),
        ("D_avg^2 × sigma_LT^2", "Uncertainty from lead time variability — how much extra stock needed because the shipment might arrive late"),
    ], col_widths=[4.5*cm, W - 9*cm]))
    story.append(SP(6))
    story.append(P(
        "<b>When to use enhanced vs standard?</b> Use enhanced (Block B) whenever lead time "
        "variability is significant (sigma_LT > 3 days). For Japan supply (sigma_LT=2 days), "
        "Block A and Block B give nearly identical results. For GPC Halberstadt (sigma_LT=10 days "
        "during ramp-up), Block B gives dramatically higher safety stock — and is correct."
    ))

    story.append(SP(10))
    story.append(section_box("6.4  The Three Scenarios — Japan vs GPC Halberstadt vs Chennai"))
    story.append(SP(8))
    story.append(data_table(
        ["Parameter", "Scenario 1: Japan (Stable)", "Scenario 2: GPC (Volatile)", "Scenario 3: Chennai (Medium)"],
        [
            ["Lead Time (LT)", "45 days", "30 days", "21 days"],
            ["LT Std Dev (sigma_LT)", "2 days", "10 days", "3 days"],
            ["Avg Daily Demand", "1.64 units/day (Clutch Disc)", "1.64 units/day", "5.0 units/day (Oil Filter)"],
            ["Demand Std Dev (monthly)", "8 units", "8 units", "15 units"],
            ["Z-Score (A-class)", "2.05", "2.05", "1.65 (B-class example)"],
            ["SS Standard (Block A)", "~47 units", "~35 units", "~52 units"],
            ["SS Enhanced (Block B)", "~48 units", "~167 units", "~62 units"],
            ["Delta", "+1 unit (negligible)", "+132 units — TRIPLES!", "+10 units"],
            ["Key Insight", "LT stable — SS barely changes", "Short LT but HIGH variance explodes SS", "Medium LT, moderate variance"],
        ],
        col_widths=[5*cm, 4.5*cm, 3.5*cm, W - 17.5*cm],
        highlight_rows=[(6, PALE_RED), (7, PALE_RED)]
    ))
    story.append(SP(6))
    story.append(P(
        "The GPC Halberstadt counterintuitive result: lead time drops from 45 to 30 days "
        "(a 33% improvement) but safety stock TRIPLES from ~47 to ~167 units. "
        "The reason: during the 2026 ramp-up, Germany's supply reliability is unpredictable. "
        "A shipment that could arrive anywhere from day 20 to day 40 creates enormous "
        "uncertainty. The enhanced formula captures this. The standard formula would "
        "incorrectly show that shorter LT means LESS safety stock — leading to chronic "
        "stockouts during the transition period.",
        "Warning"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 7 — SERVICE LEVEL OPTIMIZER
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 7", "Service Level Optimizer — NORM.S.INV & Working Capital", NAVY))
    story.append(SP(14))

    story.append(section_box("7.1  What is Service Level and Why Does It Cost Money?"))
    story.append(SP(8))
    story.append(P(
        "Service level (SL) is the <b>probability that all demand is fulfilled from stock</b> "
        "during a replenishment cycle. A 95% service level means that in 95 out of 100 "
        "replenishment cycles, no stockout occurs. The remaining 5% of cycles experience "
        "at least one unfulfilled order. "
        "Higher service levels require more safety stock — which costs more working capital."
    ))
    story.append(SP(6))
    story.append(P(
        "The relationship is non-linear. Going from 90% to 95% SL requires X more safety stock. "
        "Going from 95% to 99% requires 4× more safety stock than that step. "
        "Going from 99% to 99.9% costs enormously. "
        "This is why a <b>tiered approach by ABC class</b> (A-parts get 99%, C-parts get 80%) "
        "is dramatically more capital-efficient than blanket 99% across all SKUs.",
        "Insight"
    ))

    story.append(SP(10))
    story.append(section_box("7.2  NORM.S.INV — The Z-Score Function  |  =NORM.S.INV(0.95)"))
    story.append(SP(8))
    story.append(P(
        "Safety stock uses the <b>Normal distribution</b> to model demand uncertainty. "
        "The Z-score represents 'how many standard deviations above the mean' you need to "
        "stock to achieve a given service level. NORM.S.INV() is Excel's inverse normal "
        "distribution function — it takes a probability and returns the corresponding Z-score."
    ))
    story.append(SP(6))
    story.append(P("Formula:", "BodyBold"))
    story.append(P("Z = NORM.S.INV(service_level_probability)", "Formula"))
    story.append(P("Example: =NORM.S.INV(0.95) → returns 1.645", "FormulaSmall"))
    story.append(SP(6))
    story.append(data_table(
        ["Service Level", "NORM.S.INV Result", "Interpretation", "Used For"],
        [
            ["80%", "0.842", "Stock 0.84 std devs above mean demand", "CZ SMOB candidates — cost only"],
            ["85%", "1.036", "Stock 1.04 std devs above mean", "Slow movers CZ, CY low value"],
            ["90%", "1.282", "Stock 1.28 std devs above mean", "C-class parts"],
            ["95%", "1.645", "Stock 1.64 std devs above mean — Base", "B-class standard"],
            ["97.5%", "1.960", "Stock 1.96 std devs above mean", "B-class critical"],
            ["99%", "2.326", "Stock 2.33 std devs above mean — costly", "A-class parts: AX, AY, AZ"],
            ["99.9%", "3.090", "Stock 3.09 std devs above mean", "Life-critical safety parts only"],
        ],
        col_widths=[2.5*cm, 3*cm, 5*cm, W - 15*cm],
        highlight_rows=[(3, PALE_BLUE), (5, PALE_GREEN)]
    ))

    story.append(SP(10))
    story.append(section_box("7.3  SS Multiplier vs 95% Base  |  =NORM.S.INV(0.8)/NORM.S.INV(0.95)"))
    story.append(SP(8))
    story.append(P("Formula:", "BodyBold"))
    story.append(P("SS_Multiplier = NORM.S.INV(target_SL) / NORM.S.INV(0.95)", "Formula"))
    story.append(SP(6))
    story.append(P(
        "This normalises all safety stock values relative to the 95% baseline, "
        "showing how much MORE or LESS safety stock a different service level requires "
        "compared to the 95% standard. It is a management communication tool — "
        "executives understand '40% less safety stock for B-parts' better than raw Z-scores."
    ))
    story.append(SP(6))
    story.append(P("SS Value at a given service level:", "BodyBold"))
    story.append(P("SS_AED = Current_SS_Value × SS_Multiplier = $B$4 × C9", "Formula"))
    story.append(SP(6))
    story.append(P(
        "Example: Current SS value (at 95% base) = AED 10,000,000. "
        "SS at 80% SL = 10,000,000 × (0.842/1.645) = 10,000,000 × 0.512 = AED 5,120,000. "
        "Additional Working Capital vs 95% = 5,120,000 - 10,000,000 = <b>-AED 4,880,000</b> "
        "(a saving if you accept 80% SL for these parts).",
        "Example"
    ))

    story.append(SP(10))
    story.append(section_box("7.4  S&OP Compromise — Tiered vs Blanket 99% Approach"))
    story.append(SP(8))
    story.append(P(
        "The proposed tiered service level approach assigns different SL targets by ABC-XYZ class, "
        "reflecting the true business priority of each segment. "
        "Blanket 99% for all parts wastes working capital on CZ SMOB candidates "
        "while over-stocking gaskets instead of turbochargers."
    ))
    story.append(SP(6))
    story.append(data_table(
        ["Class", "Proposed SL", "Z-Score", "Business Rationale"],
        [
            ["AX / AY", "99%", "2.326", "High-value, high-impact. Stockout costs far exceed holding cost."],
            ["AZ", "97.5%", "1.960", "High value but sporadic. Cannot justify full 99% — too expensive with high SS."],
            ["BX / BY", "95%", "1.645", "Standard baseline. Balanced cost-service trade-off."],
            ["BZ", "92%", "1.405", "Medium value, unpredictable. Slightly below standard."],
            ["CX / CY", "90%", "1.282", "Low value, stable. Generous stock levels affordable, but not critical."],
            ["CZ", "80%", "0.842", "Low value, sporadic. SMOB risk. Accept higher stockout risk — order on demand."],
        ],
        col_widths=[2.5*cm, 2.5*cm, 2*cm, W - 11.5*cm],
        highlight_rows=[(0, PALE_GREEN), (5, PALE_RED)]
    ))
    story.append(SP(8))
    story.append(P("Working Capital Comparison:", "SubTitle"))
    story.append(data_table(
        ["Approach", "Total SS Investment (AED)", "Interpretation"],
        [
            ["Blanket 99% for all SKUs", "70,715,954", "Over-stocks low-value parts to same level as critical assemblies"],
            ["Proposed tiered ABC-XYZ", "44,369,535", "Concentrates SS investment where it matters most"],
            ["Working Capital Released", "26,346,419", "AED 26.3M freed up — can fund fleet growth or fill rate improvement"],
        ],
        col_widths=[5.5*cm, 4.5*cm, W - 14.5*cm],
        highlight_rows=[(2, PALE_GREEN)]
    ))
    story.append(SP(6))
    story.append(P(
        "AED 26.3M released by right-sizing service levels is not a compromise on quality — "
        "it is a smarter allocation of the same capital. A-class parts are protected at 99%+ "
        "while C-class parts accept 80-90% SL. The net effect: higher fill rate for critical "
        "parts, lower working capital overall. This is the core S&OP trade-off conversation.",
        "Insight"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 8 — eCanter CAUSAL FORECASTING
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 8", "eCanter Causal Forecasting — VOR Risk Triggers", NAVY))
    story.append(SP(14))

    story.append(section_box("8.1  Why eCanter Parts Need a Different Approach"))
    story.append(SP(8))
    story.append(P(
        "The Mitsubishi FUSO eCanter is an <b>electric commercial vehicle</b>. Its spare parts "
        "(Battery Cell Module, On-Board Charger, ECU, Inverter, Coolant Pump) have "
        "<b>no historical SAP APO demand data</b> — the eCanter is new to the MEA market. "
        "Standard statistical forecasting is impossible without history. "
        "Instead, demand is <b>causal</b> — driven directly by measurable vehicle condition data "
        "from the telematics system."
    ))
    story.append(SP(6))
    story.append(P(
        "This is a forward-thinking approach: instead of waiting for failure events to "
        "generate demand (reactive), telematics data predicts failure BEFORE it happens "
        "(predictive/proactive). Battery SoH (State of Health) below a threshold means "
        "replacement is imminent — even if the vehicle hasn't failed yet."
    ))

    story.append(SP(10))
    story.append(section_box("8.2  VOR Risk Flag  |  =IF(AND(C6>150,D6<0.85),\"URGENT VOR\",IF(D6<0.9,\"WATCHLIST\",\"NORMAL\"))"))
    story.append(SP(8))
    story.append(P("Formula:", "BodyBold"))
    story.append(P('=IF(AND(Odometer>150000, SoH<0.85), "URGENT VOR",\n   IF(SoH<0.90, "WATCHLIST", "NORMAL"))', "Formula"))
    story.append(SP(6))
    story.append(kv_table([
        ("C6 > 150 (odometer)", "Vehicle has driven more than 150,000 km. Battery wear accelerates significantly after this threshold (similar to 7+ year fleet age modifier)."),
        ("D6 < 0.85 (SoH)", "Battery State of Health below 85%. This is the OEM-defined critical threshold. Below 85%, range drops >30% and replacement is operationally necessary."),
        ("AND()", "BOTH conditions required for URGENT VOR. High mileage alone (new battery swap) OR low SoH alone (low-mileage damage) don't trigger URGENT."),
        ("'URGENT VOR'", "Vehicle Off Road risk is HIGH. Pre-position parts NOW. Trigger emergency stock order at next replenishment cycle."),
        ("SoH < 0.90", "Second threshold: SoH between 85-90% = degrading but not critical yet. Monitor and plan, don't order yet."),
        ("'WATCHLIST'", "Flag for next monthly S&OP review. If SoH continues declining, will hit URGENT within 1-2 months."),
        ("'NORMAL'", "SoH above 90% — battery operating within acceptable range. No immediate parts planning action needed."),
    ], col_widths=[4.5*cm, W - 9*cm]))

    story.append(SP(10))
    story.append(section_box("8.3  Trigger Logic — When to Order"))
    story.append(SP(8))
    story.append(data_table(
        ["Trigger Condition", "Parts Action", "Rationale"],
        [
            ["URGENT VOR count > 3", "Add 5× Coolant Pump (CPM-033) + 1× Battery Module (BAT-032) to next inbound order", "If 3+ vehicles simultaneously at URGENT, demand burst is imminent. Pre-position before they fail."],
            ["WATCHLIST count > 5", "Add 3× Coolant Pump to watchlist replenishment", "5 vehicles deteriorating = likely 1-2 URGENT VOR within 60 days. Get ahead of it."],
            ["Both counts normal", "No action — monitor monthly", "No imminent demand signal from telematics"],
        ],
        col_widths=[4*cm, 5.5*cm, W - 14*cm],
        highlight_rows=[(0, PALE_RED), (1, PALE_GOLD)]
    ))

    story.append(SP(10))
    story.append(section_box("8.4  HAZMAT Note — Why Battery Logistics is Critical"))
    story.append(SP(8))
    story.append(P(
        "eCanter battery modules are classified as <b>Class 9 Dangerous Goods (DG) "
        "under UN3480 (lithium ion batteries) / UN3481 (lithium ion batteries in equipment)</b>. "
        "This creates additional supply chain complexity specific to the FUSO eCanter:"
    ))
    story.append(SP(6))
    for item in [
        "<b>SoC Monitoring Every 90 Days:</b> Lithium batteries in long-term storage must be maintained at 40-60% State of Charge (SoC). Below 20% SoC risks permanent capacity loss (over-discharge). Dubai's ambient warehouse temperatures (35-45°C without climate control) accelerate degradation — Jafza bonded warehouse requires temperature-controlled storage.",
        "<b>Return Logistics Pre-Arranged:</b> Failed/replaced eCanter batteries cannot be disposed of as regular waste. Hazmat regulations require OEM-approved return logistics. This must be contracted with FUSO Japan before the first eCanter spare battery is imported — lead time for return logistics contracts is 3-6 months.",
        "<b>Air Freight Restrictions:</b> Large lithium batteries (>100Wh) have strict IATA air freight restrictions. eCanter battery modules are high-capacity — sea freight Japan→Jafza (21-30 days) is the only compliant route for large orders, unlike small parts which can be expedited by air.",
    ]:
        story.append(bullet(item))
        story.append(SP(4))
    story.append(SP(4))
    story.append(P(
        "This is why the eCanter causal model must trigger demand signals 45-60 days in advance — "
        "there is no air freight option for emergency battery modules. A demand planner who "
        "understands this constraint will pre-position based on telematics signals, "
        "not wait for VOR failures.",
        "Warning"
    ))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # CHAPTER 9 — GPC BRIDGE STOCK
    # ══════════════════════════════════════════════════════════
    story.append(chapter_box("Chapter 9", "GPC Halberstadt Bridge Stock — Transition Calculator", NAVY))
    story.append(SP(14))

    story.append(section_box("9.1  What is the GPC Halberstadt Transition?"))
    story.append(SP(8))
    story.append(P(
        "GPC Halberstadt is <b>Daimler Truck's Global Parts Center in Germany</b>. "
        "From Q1 2026, FUSO MEA will begin sourcing a portion of spare parts from GPC Halberstadt "
        "instead of exclusively from Japan. The business case is lead time reduction: "
        "Japan = 45 days, Germany = 30 days (a 33% improvement). "
        "However, during the transition ramp-up period (2026), GPC's reliability and "
        "shipping schedules to Jafza have high variability — sigma_LT = 10 days vs Japan's 2 days."
    ))
    story.append(SP(6))
    story.append(P(
        "<b>The counterintuitive finding:</b> Despite shorter average lead time, "
        "GPC supply requires MORE safety stock during the transition, not less. "
        "This is a critical insight for the FUSO finance team who may expect immediate "
        "working capital benefits from the shorter lead time. "
        "The enhanced SS formula (Chapter 6, Block B) quantifies exactly how much bridge "
        "stock is needed to maintain fill rate during the transition.",
        "Warning"
    ))

    story.append(SP(10))
    story.append(section_box("9.2  How Bridge Stock is Calculated"))
    story.append(SP(8))
    story.append(P("For each high-runner AX/BX SKU:", "BodyBold"))
    story.append(P("Bridge Stock = GPC Enhanced SS - Japan Standard SS = Delta Additional Units", "Formula"))
    story.append(SP(6))
    story.append(kv_table([
        ("Japan Standard SS", "Z × sigma_d × SQRT(LT_Japan + Review_Period) = using LT=45, sigma_LT=2"),
        ("GPC Enhanced SS", "Z × SQRT(LT_GPC × sigma_d^2 + D_avg^2 × sigma_LT_GPC^2) = using LT=30, sigma_LT=10"),
        ("Delta", "GPC Enhanced SS - Japan Standard SS = additional units that must be pre-positioned at Jafza BEFORE first GPC shipment arrives"),
        ("Bridge Cost AED", "Delta units × Unit Cost = financial commitment required for safe transition"),
        ("Priority", "CRITICAL = AX parts (top 70% value). HIGH = important BX/AX secondary. MEDIUM = BY/BX tertiary."),
    ], col_widths=[4.5*cm, W - 9*cm]))

    story.append(SP(10))
    story.append(section_box("9.3  Top 20 Bridge Stock Calculation Results"))
    story.append(SP(8))
    story.append(data_table(
        ["Part #", "Description", "Japan SS", "GPC SS", "Delta", "Bridge Cost (AED)", "Priority"],
        [
            ["FUSO-BRK-001", "Brake Pad Front", "28", "182", "154", "43,120", "CRITICAL"],
            ["FUSO-OIL-002", "Oil Filter Std", "42", "273", "231", "41,580", "CRITICAL"],
            ["FUSO-AIR-003", "Air Filter Elem.", "34", "221", "187", "44,880", "CRITICAL"],
            ["FUSO-CLT-006", "Clutch Disc", "14", "91", "77", "129,360", "CRITICAL"],
            ["FUSO-SPK-008", "Spark Plug Set", "55", "358", "303", "36,360", "CRITICAL"],
            ["FUSO-SHK-012", "Shock Absorber", "8", "52", "44", "119,680", "HIGH"],
            ["FUSO-ALT-011", "Alternator 24V", "6", "39", "33", "125,400", "HIGH"],
            ["FUSO-OIL-037", "Oil Filter HD", "21", "137", "116", "25,520", "MEDIUM"],
            ["... 12 more SKUs", "—", "—", "—", "—", "—", "—"],
            ["TOTAL", "All 20 AX/BX SKUs", "—", "—", "2,214 units", "1,070,220", "Pre-committed"],
        ],
        col_widths=[3*cm, 3.5*cm, 2*cm, 2*cm, 2*cm, 3*cm, W - 20*cm],
        highlight_rows=[(9, PALE_GREEN)]
    ))

    story.append(SP(10))
    story.append(section_box("9.4  Why This Analysis is Critical for the Business Case"))
    story.append(SP(8))
    for item in [
        "<b>Finance needs the number:</b> AED 1,070,220 bridge stock investment must be approved before GPC goes live. Without this analysis, finance has no basis for the budget request.",
        "<b>Prevents fill rate collapse:</b> If bridge stock is NOT pre-positioned and Japan shipments stop while GPC ramps up, the Jafza hub will experience stockouts across all 58 markets simultaneously — a catastrophic fill rate event.",
        "<b>Transition window:</b> The bridge stock covers the 3-6 month period when GPC LT variability is high (sigma_LT=10). Once GPC stabilises (sigma_LT drops to 4-5 days), the enhanced SS formula naturally reduces, and bridge stock can be drawn down.",
        "<b>Only AX/BX parts:</b> The bridge stock calculation focuses on the top 20 high-runner parts (AX/BX class). Lower-value CZ/CY parts are NOT included — their stockout impact is minimal and the bridge investment cost would be unjustified.",
    ]:
        story.append(bullet(item))
        story.append(SP(4))

    story.append(SP(10))
    story.append(section_box("9.5  Summary — The Five Model Innovations vs SAP APO Baseline"))
    story.append(SP(8))
    story.append(data_table(
        ["Innovation", "Problem Solved", "Impact"],
        [
            ["SBA (Syntetos-Boylan)", "SAP APO over-forecasts intermittent demand by ~20% (Croston's bias)", "Forecast accuracy: 70% → 85% target"],
            ["Enhanced SS Formula", "Standard SS ignores LT variability — understates SS for GPC transition", "Prevents fill rate collapse during 2026 transition"],
            ["eCanter Causal", "No history for EV parts — statistical methods fail completely", "Zero stockouts on AED 18,000 battery modules"],
            ["Fleet Age Modifier", "Global forecast ignores that Egypt's 2016 fleet needs 50% more parts than UAE's 2024 fleet", "Market-specific accuracy improvement"],
            ["Tiered Service Levels", "Blanket 99% wastes AED 26.3M on low-value parts", "AED 26.3M working capital released for better use"],
        ],
        col_widths=[4.5*cm, 5.5*cm, W - 14.5*cm],
        highlight_rows=[(0, PALE_GREEN), (1, PALE_BLUE), (2, PALE_GOLD), (3, PALE_GREEN), (4, PALE_GREEN)]
    ))

    story.append(SP(12))
    story.append(divider(NAVY))
    story.append(SP(8))
    story.append(P(
        "<b>This document is the complete intellectual framework behind the FUSO MEA "
        "Advanced Demand Planning Model. Every formula has a business reason. Every "
        "assumption is defensible. Every threshold is industry-standard or data-derived. "
        "The model is not a black box — it is a transparent, auditable decision system "
        "built for a AED 38M, 58-market, 50-SKU spare parts operation.</b>",
        "SubTitle"
    ))

    # ── BUILD ──────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f"PDF written: {OUT}")
    print(f"Size: {os.path.getsize(OUT):,} bytes")

if __name__ == "__main__":
    build()
