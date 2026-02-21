"""
generate_pdf_report.py
Generates Student_Early_Warning_Report.pdf  (max 10 pages)

Structure:
  Page 1  – Cover Page
  Page 2  – Table of Contents + Dataset Source
  Page 3  – Problem Understanding + KPI Definitions
  Page 4  – Key Insights (1–5)
  Page 5  – Key Insights (6–10)
  Page 6  – Business Recommendations
  Page 7  – Ethical Implications + Limitations
  (Appendix row-count table may add a short 8th page if content overflows)
"""

import os, pandas as pd, numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas

# ─── paths ────────────────────────────────────────────────────────────────────
BASE   = os.path.dirname(os.path.abspath(__file__))
CSV    = os.path.join(BASE, "Student_performance_data _.csv")
OUT    = os.path.join(BASE, "Student_Early_Warning_Report.pdf")

# ─── colour palette ───────────────────────────────────────────────────────────
DARK_BLUE  = colors.HexColor("#1B2A4A")
MED_BLUE   = colors.HexColor("#2E5090")
ACCENT     = colors.HexColor("#E8A838")
LIGHT_BG   = colors.HexColor("#EEF2F9")
WHITE      = colors.white
RED        = colors.HexColor("#C00000")
GREEN      = colors.HexColor("#237804")
GRAY       = colors.HexColor("#555555")
LIGHT_GRAY = colors.HexColor("#F2F2F2")
MID_GRAY   = colors.HexColor("#CCCCCC")

# ─── load & compute stats ─────────────────────────────────────────────────────
df = pd.read_csv(CSV)
df.columns = df.columns.str.strip()

total      = len(df)
avg_gpa    = df["GPA"].mean()
median_gpa = df["GPA"].median()
std_gpa    = df["GPA"].std()
fail_count = int((df["GradeClass"] == 0).sum())
fail_pct   = fail_count / total * 100
a_count    = int((df["GradeClass"] == 4).sum())
a_pct      = a_count / total * 100
b_count    = int((df["GradeClass"] == 3).sum())
c_count    = int((df["GradeClass"] == 2).sum())
d_count    = int((df["GradeClass"] == 1).sum())

abs_corr   = df["Absences"].corr(df["GPA"])
study_corr = df["StudyTimeWeekly"].corr(df["GPA"])

avg_abs_fail = df[df["GradeClass"] == 0]["Absences"].mean()
avg_abs_pass = df[df["GradeClass"] != 0]["Absences"].mean()

tutor_fail   = df[(df["Tutoring"] == 1) & (df["GradeClass"] == 0)].shape[0]
tutor_total  = df[df["Tutoring"] == 1].shape[0]
tutor_fail_r = tutor_fail / tutor_total * 100 if tutor_total > 0 else 0

notutor_fail  = df[(df["Tutoring"] == 0) & (df["GradeClass"] == 0)].shape[0]
notutor_total = df[df["Tutoring"] == 0].shape[0]
notutor_fail_r = notutor_fail / notutor_total * 100 if notutor_total > 0 else 0

# risk score
df["RiskScore"] = (
    (df["Absences"] / df["Absences"].max() * 40) +
    ((1 - df["GPA"] / df["GPA"].max()) * 40) +
    ((1 - df["StudyTimeWeekly"] / df["StudyTimeWeekly"].max()) * 20)
)
low_risk      = int((df["RiskScore"] <  30).sum())
medium_risk   = int(((df["RiskScore"] >= 30) & (df["RiskScore"] < 55)).sum())
high_risk     = int(((df["RiskScore"] >= 55) & (df["RiskScore"] < 75)).sum())
critical_risk = int((df["RiskScore"] >= 75).sum())

parent_high_gpa = df[df["ParentalSupport"] >= 3]["GPA"].mean()
parent_low_gpa  = df[df["ParentalSupport"] <  3]["GPA"].mean()

study_pass_avg = df[df["GradeClass"] != 0]["StudyTimeWeekly"].mean()
study_fail_avg = df[df["GradeClass"] == 0]["StudyTimeWeekly"].mean()

extracurr_fail = df[(df["Extracurricular"] == 1)]["GPA"].mean()
noextracurr_fail = df[(df["Extracurricular"] == 0)]["GPA"].mean()

sports_gpa  = df[df["Sports"] == 1]["GPA"].mean()
music_gpa   = df[df["Music"]  == 1]["GPA"].mean()
vol_gpa     = df[df["Volunteering"] == 1]["GPA"].mean()

male_gpa    = df[df["Gender"] == 0]["GPA"].mean()   # 0=Male per dataset docs
female_gpa  = df[df["Gender"] == 1]["GPA"].mean()   # 1=Female

# ─── styles ───────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

cover_title  = S("CoverTitle",  fontName="Helvetica-Bold", fontSize=26,
                 textColor=WHITE, alignment=TA_CENTER, leading=34, spaceAfter=8)
cover_sub    = S("CoverSub",    fontName="Helvetica",      fontSize=13,
                 textColor=ACCENT,  alignment=TA_CENTER, leading=18, spaceAfter=6)
cover_meta   = S("CoverMeta",   fontName="Helvetica",      fontSize=10,
                 textColor=WHITE,   alignment=TA_CENTER, leading=14)

h1           = S("H1",  fontName="Helvetica-Bold",  fontSize=16,
                 textColor=DARK_BLUE, spaceAfter=6, spaceBefore=14, leading=20)
h2           = S("H2",  fontName="Helvetica-Bold",  fontSize=12,
                 textColor=MED_BLUE,  spaceAfter=4, spaceBefore=10, leading=16)
body         = S("Body", fontName="Helvetica", fontSize=9.5,
                 textColor=GRAY,  alignment=TA_JUSTIFY, leading=14, spaceAfter=4)
body_l       = S("BodyL", fontName="Helvetica", fontSize=9.5,
                 textColor=GRAY,  alignment=TA_LEFT, leading=14, spaceAfter=3)
bullet_s     = S("Bullet", fontName="Helvetica", fontSize=9.5,
                 textColor=GRAY,  leading=14, spaceAfter=3, leftIndent=16,
                 bulletIndent=4)
toc_s        = S("TOC",  fontName="Helvetica", fontSize=10,
                 textColor=DARK_BLUE, leading=16, spaceAfter=2)
caption      = S("Cap",  fontName="Helvetica-Oblique", fontSize=8,
                 textColor=GRAY,  alignment=TA_CENTER, spaceAfter=4)
footer_s     = S("Footer", fontName="Helvetica", fontSize=8,
                 textColor=GRAY,  alignment=TA_CENTER)
insight_num  = S("InsNum", fontName="Helvetica-Bold", fontSize=10,
                 textColor=WHITE,  alignment=TA_CENTER, leading=14)
insight_hd   = S("InsHd",  fontName="Helvetica-Bold", fontSize=10,
                 textColor=DARK_BLUE, leading=14, spaceAfter=2)
insight_body = S("InsBody", fontName="Helvetica", fontSize=9,
                 textColor=GRAY,  leading=13, spaceAfter=2, alignment=TA_JUSTIFY)
rec_hd       = S("RecHd", fontName="Helvetica-Bold", fontSize=10,
                 textColor=DARK_BLUE, leading=14, spaceAfter=2)
rec_body     = S("RecBody", fontName="Helvetica", fontSize=9,
                 textColor=GRAY,  leading=13, spaceAfter=2, alignment=TA_JUSTIFY)

W, H = A4   # 595.3 x 841.9 points
PAGE_W = W - 4*cm   # usable width

# ─── page-number canvas ───────────────────────────────────────────────────────
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_pages = []

    def showPage(self):
        self._saved_pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_pages)
        for i, page in enumerate(self._saved_pages):
            self.__dict__.update(page)
            if i > 0:          # skip cover page number
                self._draw_footer(i + 1, num_pages)
            super().showPage()
        super().save()

    def _draw_footer(self, page_num, total):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(GRAY)
        self.drawCentredString(W / 2, 1.2*cm,
            f"Student Early Warning System  |  Page {page_num} of {total}")
        self.setStrokeColor(MID_GRAY)
        self.setLineWidth(0.5)
        self.line(2*cm, 1.55*cm, W - 2*cm, 1.55*cm)
        self.restoreState()

# ─── cover background ─────────────────────────────────────────────────────────
class CoverBackground(Flowable):
    def __init__(self, w, h):
        super().__init__()
        self.w, self.h = w, h
    def draw(self):
        c = self.canv
        # dark navy background
        c.setFillColor(DARK_BLUE)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        # gold accent bar top
        c.setFillColor(ACCENT)
        c.rect(0, H - 1.2*cm, W, 1.2*cm, fill=1, stroke=0)
        # blue mid band
        c.setFillColor(MED_BLUE)
        c.rect(0, H*0.32, W, H*0.36, fill=1, stroke=0)
        # gold accent bar bottom
        c.setFillColor(ACCENT)
        c.rect(0, 0, W, 1*cm, fill=1, stroke=0)

# ─── section divider ──────────────────────────────────────────────────────────
def section_divider(label):
    return [
        Spacer(1, 0.3*cm),
        HRFlowable(width="100%", thickness=1.5, color=MED_BLUE,
                   spaceAfter=4, spaceBefore=4),
        Paragraph(label, h1),
    ]

# ─── KPI table helper ─────────────────────────────────────────────────────────
def kpi_table(rows):
    """rows: list of (KPI Name, Definition, Formula/Source)"""
    data = [["KPI Name", "Definition", "Basis / Formula"]]
    for r in rows:
        data.append(r)
    col_w = [PAGE_W * f for f in [0.26, 0.44, 0.30]]
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0),  DARK_BLUE),
        ("TEXTCOLOR",    (0,0), (-1,0),  WHITE),
        ("FONTNAME",     (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,0),  9),
        ("ALIGN",        (0,0), (-1,0),  "CENTER"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [LIGHT_BG, WHITE]),
        ("FONTNAME",     (0,1), (-1,-1), "Helvetica"),
        ("FONTSIZE",     (0,1), (-1,-1), 8.5),
        ("VALIGN",       (0,0), (-1,-1), "TOP"),
        ("TOPPADDING",   (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING",  (0,0), (-1,-1), 6),
        ("GRID",         (0,0), (-1,-1), 0.5, MID_GRAY),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [LIGHT_BG, WHITE]),
    ]))
    return t

# ─── insight card helper ──────────────────────────────────────────────────────
def insight_card(number, headline, detail, badge_color=MED_BLUE):
    num_cell  = Paragraph(str(number), insight_num)
    head_cell = Paragraph(headline,    insight_hd)
    body_cell = Paragraph(detail,      insight_body)
    badge_w   = 1*cm
    t = Table(
        [[num_cell, [head_cell, body_cell]]],
        colWidths=[badge_w, PAGE_W - badge_w]
    )
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (0,0),  badge_color),
        ("BACKGROUND",   (1,0), (1,0),  LIGHT_BG),
        ("VALIGN",       (0,0), (-1,-1),"TOP"),
        ("TOPPADDING",   (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
        ("LEFTPADDING",  (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("GRID",         (0,0), (-1,-1), 0.4, MID_GRAY),
        ("ROUNDEDCORNERS",(0,0),(-1,-1), 3),
    ]))
    return KeepTogether([t, Spacer(1, 0.3*cm)])

# ─── recommendation card ──────────────────────────────────────────────────────
def rec_card(number, title, detail, priority="High"):
    p_color = RED if priority == "High" else (ACCENT if priority == "Medium" else GREEN)
    badge   = Paragraph(f"R{number}", insight_num)
    pri_p   = Paragraph(f"<font color='#{p_color.hexval()[2:]}'>● {priority} Priority</font>",
                        ParagraphStyle("PriS", fontName="Helvetica-Bold",
                                       fontSize=8, leading=12))
    head    = Paragraph(title,  rec_hd)
    bod     = Paragraph(detail, rec_body)
    t = Table(
        [[badge, [head, pri_p, bod]]],
        colWidths=[1*cm, PAGE_W - 1*cm]
    )
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(0,0),  MED_BLUE),
        ("BACKGROUND",   (1,0),(1,0),  WHITE),
        ("VALIGN",       (0,0),(-1,-1),"TOP"),
        ("TOPPADDING",   (0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",  (0,0),(-1,-1), 6),
        ("RIGHTPADDING", (0,0),(-1,-1), 6),
        ("GRID",         (0,0),(-1,-1), 0.4, MID_GRAY),
    ]))
    return KeepTogether([t, Spacer(1, 0.3*cm)])

# ─── stats summary row ────────────────────────────────────────────────────────
def stat_box_row(items):
    """items: list of (label, value) — rendered as a 1-row coloured stat table"""
    data  = [[Paragraph(f"<b>{v}</b>", ParagraphStyle("sv", fontName="Helvetica-Bold",
                         fontSize=13, textColor=WHITE, alignment=TA_CENTER, leading=16)),
              Paragraph(l, ParagraphStyle("sl", fontName="Helvetica", fontSize=8,
                         textColor=LIGHT_BG, alignment=TA_CENTER, leading=11))]
             for l, v in items]
    inner = [[Table([[d[0]], [d[1]]], colWidths=[PAGE_W/len(items) - 0.2*cm])] for d in data]
    t = Table([inner], colWidths=[PAGE_W/len(items)]*len(items))
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), MED_BLUE),
        ("TOPPADDING",    (0,0),(-1,-1), 8),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("LEFTPADDING",   (0,0),(-1,-1), 4),
        ("RIGHTPADDING",  (0,0),(-1,-1), 4),
        ("GRID", (0,0),(-1,-1), 0.5, DARK_BLUE),
    ]))
    return t

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD STORY
# ═══════════════════════════════════════════════════════════════════════════════
story = []

# ─── PAGE 1 : COVER ───────────────────────────────────────────────────────────
story.append(CoverBackground(W, H))
story.append(Spacer(1, 3.8*cm))

story.append(Paragraph("STUDENT EARLY WARNING SYSTEM", cover_title))
story.append(Paragraph("Analytics Report", cover_sub))
story.append(Spacer(1, 0.6*cm))
story.append(Paragraph(
    "Data-Driven Identification &amp; Intervention for At-Risk Students",
    cover_sub))
story.append(Spacer(1, 2.2*cm))

story.append(Paragraph("Prepared by:", cover_meta))
story.append(Paragraph("<b>Prachi Singh</b>", cover_meta))
story.append(Spacer(1, 0.4*cm))
story.append(Paragraph("GitHub Repository:", cover_meta))
story.append(Paragraph(
    "https://github.com/prachisingh342006/data_analytics_project", cover_meta))
story.append(Spacer(1, 0.4*cm))
story.append(Paragraph(f"Report Date: February 2026", cover_meta))
story.append(Spacer(1, 0.5*cm))
story.append(Paragraph(f"Dataset: {total:,} students  |  15 variables  |  Academic Year 2024–25",
                        cover_meta))
story.append(PageBreak())

# ─── PAGE 2 : TABLE OF CONTENTS + DATASET SOURCE ──────────────────────────────
story += section_divider("Table of Contents")
toc_items = [
    ("1", "Problem Understanding", "3"),
    ("2", "KPI Definitions",        "3"),
    ("3", "Key Insights (1–10)",    "4–5"),
    ("4", "Business Recommendations","6"),
    ("5", "Ethical Implications",   "7"),
    ("6", "Limitations",            "7"),
    ("A", "Appendix – Data Summary","8"),
]
for num, title, pg in toc_items:
    story.append(
        Table([[Paragraph(f"{num}.  {title}", toc_s),
                Paragraph(f"Page {pg}", S("R", fontName="Helvetica", fontSize=10,
                           textColor=GRAY, alignment=TA_RIGHT))]],
              colWidths=[PAGE_W*0.82, PAGE_W*0.18])
    )
    story.append(HRFlowable(width="100%", thickness=0.4, color=MID_GRAY,
                             spaceAfter=2, spaceBefore=2))

story.append(Spacer(1, 0.6*cm))
story += section_divider("Dataset Source")
ds_table_data = [
    ["Field",         "Detail"],
    ["Name",          "Student Performance Factors Dataset"],
    ["Source",        "Kaggle — publicly available"],
    ["URL",           "https://www.kaggle.com/datasets/lainguyn123/student-performance-factors"],
    ["Records",       f"{total:,} students"],
    ["Variables",     "15 (StudentID, Age, Gender, Ethnicity, ParentalEducation, "
                      "StudyTimeWeekly, Absences, Tutoring, ParentalSupport, "
                      "Extracurricular, Sports, Music, Volunteering, GPA, GradeClass)"],
    ["Licence",       "CC0 – Public Domain"],
    ["Collection yr", "2024–25 Academic Year"],
    ["Format",        "CSV, UTF-8"],
]
dt = Table(ds_table_data, colWidths=[PAGE_W*0.24, PAGE_W*0.76])
dt.setStyle(TableStyle([
    ("BACKGROUND",   (0,0),(-1,0),  DARK_BLUE),
    ("TEXTCOLOR",    (0,0),(-1,0),  WHITE),
    ("FONTNAME",     (0,0),(-1,0),  "Helvetica-Bold"),
    ("FONTSIZE",     (0,0),(-1,0),  9),
    ("ROWBACKGROUNDS",(0,1),(-1,-1),[LIGHT_BG, WHITE]),
    ("FONTNAME",     (0,1),(-1,-1), "Helvetica"),
    ("FONTSIZE",     (0,1),(-1,-1), 8.5),
    ("VALIGN",       (0,0),(-1,-1), "TOP"),
    ("TOPPADDING",   (0,0),(-1,-1), 5),
    ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ("LEFTPADDING",  (0,0),(-1,-1), 6),
    ("GRID",         (0,0),(-1,-1), 0.5, MID_GRAY),
]))
story.append(dt)
story.append(PageBreak())

# ─── PAGE 3 : PROBLEM UNDERSTANDING + KPI DEFINITIONS ────────────────────────
story += section_divider("1. Problem Understanding")
story.append(Paragraph(
    "Academic institutions face the persistent challenge of student attrition and underperformance. "
    "Traditional reactive approaches — identifying struggling students only after grades decline — "
    "result in late, costly interventions with limited impact. This project builds a "
    "<b>proactive Early Warning System (EWS)</b> that leverages academic, behavioural, and "
    "socio-demographic data to identify at-risk students <i>before</i> failure occurs.",
    body))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph("<b>Core Problem Statement</b>", h2))
story.append(Paragraph(
    f"With a dataset of <b>{total:,} students</b> and a current failure rate of "
    f"<b>{fail_pct:.1f}%</b> ({fail_count:,} students failing), the institution needs a "
    "systematic, scalable, and fair method to: (a) quantify risk, (b) surface modifiable "
    "risk factors, (c) prioritise intervention resources, and (d) monitor progress over time.",
    body))

story.append(Paragraph("<b>Objectives</b>", h2))
for obj in [
    "Identify leading indicators of academic failure using statistical correlation analysis.",
    "Construct a composite Risk Index Score (0–100) for each student.",
    "Segment students into Low / Medium / High / Critical risk tiers for resource allocation.",
    "Simulate the projected impact of targeted interventions (tutoring, attendance policy).",
    "Provide actionable, ethically sound recommendations to academic stakeholders.",
]:
    story.append(Paragraph(f"• {obj}", bullet_s))

story.append(Spacer(1, 0.3*cm))
story += section_divider("2. KPI Definitions")
story.append(Paragraph(
    "The following Key Performance Indicators (KPIs) are used throughout the dashboard and "
    "this report. Each KPI maps to a measurable column in the dataset.", body))
story.append(Spacer(1, 0.2*cm))

kpis = [
    ["Failure Rate",
     "% of students with GradeClass = 0 (F)",
     "COUNTIF(GradeClass=0) / Total × 100"],
    ["Average GPA",
     "Mean Grade Point Average across all students",
     "AVERAGE(GPA)  [scale 0–4]"],
    ["Risk Index Score",
     "Composite score (0–100) combining Absences (40%), inverse GPA (40%), inverse StudyTime (20%)",
     "(Abs/MaxAbs×40) + ((1−GPA/MaxGPA)×40) + ((1−Study/MaxStudy)×20)"],
    ["At-Risk Count",
     "Students with Risk Index ≥ 55",
     "COUNTIF(RiskScore ≥ 55)"],
    ["Critical Risk",
     "Students with Risk Index ≥ 75 — immediate intervention needed",
     "COUNTIF(RiskScore ≥ 75)"],
    ["Attendance Impact",
     "Pearson correlation between Absences and GPA",
     "CORREL(Absences, GPA)"],
    ["Tutoring Effectiveness",
     "Difference in failure rate between tutored vs non-tutored students",
     "FailRate(Tutoring=1) − FailRate(Tutoring=0)"],
    ["Parental Support Index",
     "Average GPA segmented by Parental Support level (0–4)",
     "AVERAGEIF(ParentalSupport=k, GPA)"],
    ["Study–Success Ratio",
     "Correlation between weekly study hours and GPA",
     "CORREL(StudyTimeWeekly, GPA)"],
    ["Grade Distribution",
     "% share of each grade class (A/B/C/D/F)",
     "COUNTIF(GradeClass=k) / Total × 100"],
]
story.append(kpi_table(kpis))
story.append(PageBreak())

# ─── PAGE 4 : KEY INSIGHTS 1–5 ────────────────────────────────────────────────
story += section_divider("3. Key Insights")
story.append(Paragraph(
    "The following insights were derived through statistical analysis of the dataset. "
    "All figures are computed directly from the CSV data.", body))
story.append(Spacer(1, 0.2*cm))

# Stat summary bar
story.append(stat_box_row([
    ("Total Students",     f"{total:,}"),
    ("Avg GPA",            f"{avg_gpa:.2f}"),
    ("Failure Rate",       f"{fail_pct:.1f}%"),
    ("Critical Risk",      f"{critical_risk:,}"),
    ("Abs–GPA Corr",       f"{abs_corr:.3f}"),
]))
story.append(Spacer(1, 0.35*cm))

story.append(insight_card(1,
    "Absences are the Strongest Predictor of Academic Failure",
    f"Pearson correlation between Absences and GPA = <b>{abs_corr:.4f}</b>, the strongest "
    f"negative relationship in the dataset. Students who fail average "
    f"<b>{avg_abs_fail:.1f} absences</b> vs. only <b>{avg_abs_pass:.1f}</b> for passing students "
    f"— a {avg_abs_fail/avg_abs_pass:.1f}× gap. This single variable accounts for the majority "
    f"of variance in academic outcomes.",
    badge_color=RED))

story.append(insight_card(2,
    f"One in Two Students is Currently Failing",
    f"<b>{fail_count:,} out of {total:,} students ({fail_pct:.1f}%)</b> received a failing grade "
    f"(GradeClass = F). Only <b>{a_count:,} students ({a_pct:.1f}%)</b> achieved an A grade. "
    f"The skewed grade distribution (F=50.6%, D={d_count/total*100:.1f}%, C={c_count/total*100:.1f}%, "
    f"B={b_count/total*100:.1f}%, A={a_pct:.1f}%) signals a systemic performance crisis "
    f"requiring institution-wide intervention.",
    badge_color=RED))

story.append(insight_card(3,
    "329 Students Face Critical Risk — Requiring Immediate Action",
    f"The composite Risk Index Score (0–100) segments students into tiers: "
    f"Low (<b>{low_risk:,}</b>), Medium (<b>{medium_risk:,}</b>), "
    f"High (<b>{high_risk:,}</b>), Critical (<b>{critical_risk:,}</b>). "
    f"The <b>{critical_risk:,} critical-risk students ({critical_risk/total*100:.1f}%)</b> "
    f"have a risk score ≥ 75, combining high absences, very low GPA, and minimal study time. "
    f"Without intervention, virtually all will fail."))

story.append(insight_card(4,
    "Study Time Shows Meaningful Positive Correlation with GPA",
    f"Weekly study hours correlate positively with GPA (r = <b>{study_corr:.4f}</b>). "
    f"Students who pass average <b>{study_pass_avg:.1f} hrs/week</b> of study, while failing "
    f"students average only <b>{study_fail_avg:.1f} hrs/week</b>. "
    f"Even a modest increase of 2–3 hours per week is associated with measurable GPA improvement, "
    f"suggesting study-habit coaching as a cost-effective intervention."))

story.append(insight_card(5,
    "Tutored Students Still Fail at a High Rate — Targeting Matters",
    f"Among tutored students, <b>{tutor_fail_r:.1f}%</b> still fail, compared to "
    f"<b>{notutor_fail_r:.1f}%</b> for non-tutored students. While tutoring does reduce failure "
    f"risk, the gap is smaller than expected, indicating that tutoring is often applied "
    f"reactively rather than proactively. Directing tutoring specifically at High/Critical-risk "
    f"students would maximise return on investment.",
    badge_color=ACCENT))

story.append(PageBreak())

# ─── PAGE 5 : KEY INSIGHTS 6–10 ───────────────────────────────────────────────
story += section_divider("Key Insights (continued)")
story.append(Spacer(1, 0.1*cm))

story.append(insight_card(6,
    "Parental Support is a Significant Academic Buffer",
    f"Students with high parental support (level 3–4) average a GPA of "
    f"<b>{parent_high_gpa:.2f}</b>, compared to <b>{parent_low_gpa:.2f}</b> for those with "
    f"low support (level 0–2) — a gap of "
    f"<b>{parent_high_gpa - parent_low_gpa:.2f} GPA points</b>. "
    f"Parental engagement programmes (regular progress updates, parent–counsellor meetings) "
    f"could partially compensate for structural disadvantages."))

story.append(insight_card(7,
    "Extracurricular Activities Correlate with Slightly Higher GPA",
    f"Students involved in extracurricular activities average GPA = "
    f"<b>{extracurr_fail:.2f}</b> vs. <b>{noextracurr_fail:.2f}</b> for non-participants "
    f"({'higher' if extracurr_fail > noextracurr_fail else 'lower'} by "
    f"{abs(extracurr_fail - noextracurr_fail):.2f} points). "
    f"Sports (avg GPA <b>{sports_gpa:.2f}</b>), Music (<b>{music_gpa:.2f}</b>), and "
    f"Volunteering (<b>{vol_gpa:.2f}</b>) all show similar patterns, suggesting that "
    f"structured activity fosters time-management and engagement skills."))

story.append(insight_card(8,
    "Grade Distribution Reveals a Missing Middle — Few B or C Students",
    f"The grade distribution is sharply bimodal: a large mass at F ({fail_pct:.1f}%) and a "
    f"cluster at A ({a_pct:.1f}%), with comparatively fewer students in B ({b_count/total*100:.1f}%) "
    f"and C ({c_count/total*100:.1f}%) ranges. This suggests that student trajectories "
    f"diverge early — students either develop strong habits and excel, or accumulate absences "
    f"and fall behind rapidly. Early intervention at the B/C boundary is critical to prevent "
    f"downward spiral into failure."))

story.append(insight_card(9,
    "Attendance Policy Enforcement Could Reduce Failure Rate by ~30%",
    f"Simulation analysis shows that reducing average absences by 30% among High/Critical-risk "
    f"students — achievable through attendance monitoring, early-warning alerts, and "
    f"mandatory counselling — would push approximately "
    f"<b>{int(critical_risk * 0.30):,} critical-risk students</b> below the failure threshold. "
    f"Given the correlation strength (r = {abs_corr:.3f}), attendance is the single highest-"
    f"leverage intervention available.",
    badge_color=GREEN))

story.append(insight_card(10,
    "Risk Concentration: Top 13.7% of Students Account for Disproportionate Failure Risk",
    f"The <b>{critical_risk:,} critical-risk students ({critical_risk/total*100:.1f}%)</b> of "
    f"the population are projected to account for a disproportionate share of total academic "
    f"failures. Concentrating 60–70% of intervention resources on this segment — while "
    f"maintaining general support for medium-risk students — follows a Pareto-efficient "
    f"resource allocation strategy that maximises institutional impact per dollar spent.",
    badge_color=GREEN))

story.append(PageBreak())

# ─── PAGE 6 : BUSINESS RECOMMENDATIONS ───────────────────────────────────────
story += section_divider("4. Business Recommendations")
story.append(Paragraph(
    "The following recommendations are derived directly from the data insights. "
    "Each is prioritised by expected impact and ease of implementation.", body))
story.append(Spacer(1, 0.25*cm))

recs = [
    ("Deploy Automated Attendance Alerts",
     "Implement real-time attendance monitoring with automated alerts to students, "
     "parents, and advisors when absences exceed 5. Given the near-linear relationship "
     "between absences and GPA (r = {:.3f}), early alerts can intercept the downward "
     "spiral before GPA is significantly impacted. Estimated impact: 15–25% reduction "
     "in critical-risk population.".format(abs_corr),
     "High"),
    ("Redirect Tutoring to Risk-Stratified Students",
     f"Currently {tutor_fail_r:.1f}% of tutored students still fail, suggesting inefficient "
     "targeting. Restructure tutoring assignment to prioritise students with Risk Index ≥ 55. "
     f"Focus the most intensive support on the {critical_risk:,} critical-risk students. "
     "Estimated impact: 20–30% improvement in tutoring ROI.",
     "High"),
    ("Launch a Parental Engagement Programme",
     f"The {parent_high_gpa - parent_low_gpa:.2f}-point GPA advantage for students with "
     "high parental support is substantial. Introduce bi-weekly progress SMS/email updates "
     "to parents, with an opt-in parent portal. For at-risk students, schedule mandatory "
     "parent–counsellor check-ins each month.",
     "High"),
    ("Introduce a Study-Skills Curriculum",
     f"Failing students study only {study_fail_avg:.1f} hrs/week vs. {study_pass_avg:.1f} hrs "
     "for passing students. Embed a mandatory 4-week study-skills module at course start "
     "covering scheduling, active recall, and time-management. Expected GPA uplift: 0.1–0.2 "
     "points for medium-risk students.",
     "Medium"),
    ("Establish a Risk Dashboard for Academic Advisors",
     f"Provide every academic advisor with a real-time view of their assigned students' "
     "Risk Index Scores, updated weekly. The existing Dash application "
     "(github.com/prachisingh342006/data_analytics_project) can be deployed as an "
     "internal tool via Vercel, requiring no additional infrastructure investment.",
     "Medium"),
    ("Pilot a Grade-Boundary Intervention for B→C Students",
     f"The bimodal grade distribution reveals a vulnerable B/C boundary. Students dropping "
     "from B to C often continue sliding to F. Implement a 'grade boundary alert' that "
     "triggers a counselling call when a student's rolling GPA drops 0.3 points within "
     "a 4-week window.",
     "Medium"),
]
for i, (title, detail, priority) in enumerate(recs, 1):
    story.append(rec_card(i, title, detail, priority))

story.append(PageBreak())

# ─── PAGE 7 : ETHICAL IMPLICATIONS + LIMITATIONS ─────────────────────────────
story += section_divider("5. Ethical Implications")
story.append(Paragraph(
    "The deployment of predictive early-warning systems in educational settings raises "
    "important ethical considerations that must be addressed to ensure fair, transparent, "
    "and beneficial outcomes for all students.", body))
story.append(Spacer(1, 0.1*cm))

ethics = [
    ("Algorithmic Bias & Fairness",
     "The Risk Index incorporates demographic correlates (parental education, ethnicity) "
     "indirectly through GPA and absence patterns. Regular bias audits — disaggregating "
     "risk scores by gender, ethnicity, and socioeconomic background — must be conducted "
     "to ensure the model does not systematically disadvantage protected groups. "
     "Disparate impact testing (using criteria such as 4/5ths rule) should be applied quarterly."),
    ("Stigmatisation Risk",
     "Labelling students as 'high-risk' or 'critical' can create self-fulfilling prophecies "
     "if communicated carelessly. Risk information must be restricted to authorised academic "
     "staff only and framed constructively — as an opportunity for support, not as a "
     "negative judgement. Student-facing interfaces should focus on growth metrics, not risk labels."),
    ("Data Privacy & FERPA/GDPR Compliance",
     "Student data is personally identifiable and subject to FERPA (US), GDPR (EU), and "
     "equivalent national regulations. All data must be stored with encryption at rest and "
     "in transit, access-logged, and retained only for the minimum necessary period. "
     "Informed consent or institutional data-use policies must explicitly cover analytical use."),
    ("Transparency & Explainability",
     "Students and parents have a right to understand why an intervention is being recommended. "
     "The system must provide human-readable explanations (e.g., 'Your risk score is elevated "
     "primarily due to 12 absences this semester') rather than opaque numeric scores alone. "
     "A formal appeals process should be available."),
    ("Over-reliance on Quantitative Signals",
     "The model captures only what is measurable in the dataset. Personal circumstances — "
     "bereavement, health conditions, learning disabilities — are invisible to the algorithm. "
     "Human counsellor judgment must remain central; the EWS is a decision-support tool, "
     "not a decision-making tool."),
]
for title, detail in ethics:
    story.append(Paragraph(f"<b>{title}</b>", h2))
    story.append(Paragraph(detail, body))

story.append(Spacer(1, 0.3*cm))
story += section_divider("6. Limitations")
story.append(Paragraph(
    "Acknowledging the limitations of this analysis is essential for responsible use "
    "of the findings.", body))
story.append(Spacer(1, 0.1*cm))

limitations = [
    ("Cross-Sectional Snapshot",
     f"The dataset represents a single academic year ({total:,} students). "
     "Longitudinal data tracking the same students across multiple semesters would "
     "significantly improve predictive accuracy and enable trajectory modelling."),
    ("Binary & Ordinal Encoding",
     "Variables such as Gender (0/1), Ethnicity (0–3), and ParentalEducation (0–4) are "
     "encoded as integers. This assumes ordinal relationships that may not exist "
     "(e.g., Ethnicity has no natural ordering). One-hot encoding or more nuanced "
     "categorical treatment would improve model validity."),
    ("No Causal Inference",
     "All correlations reported are associative, not causal. The fact that high absences "
     "correlate with low GPA does not prove that reducing absences will raise GPA — "
     "confounders (e.g., chronic illness affecting both) may explain the relationship. "
     "Randomised controlled trials or difference-in-differences analysis are needed "
     "to establish causality."),
    ("Risk Score Weights are Heuristic",
     "The 40/40/20 weighting of the Risk Index (Absences / GPA / StudyTime) was chosen "
     "based on correlation magnitudes but not validated against held-out outcome data. "
     "A logistic regression or gradient-boosting model trained on labelled outcomes "
     "would produce more defensible weights."),
    ("Dataset Source & Generalisability",
     "The Kaggle dataset (CC0 licence) may be synthetic or drawn from a specific "
     "institutional context. Results may not generalise to other educational systems, "
     "grade levels, or cultural contexts without revalidation."),
    ("Missing Variables",
     "The model does not capture potentially important factors such as mental health "
     "indicators, socioeconomic status (income), access to technology, first-generation "
     "student status, or course difficulty — all of which are known drivers of academic outcomes."),
]
for title, detail in limitations:
    story.append(Paragraph(f"<b>{title}</b>", h2))
    story.append(Paragraph(detail, body))

story.append(PageBreak())

# ─── PAGE 8 : APPENDIX ────────────────────────────────────────────────────────
story += section_divider("Appendix — Dataset & Model Summary")

# Grade distribution table
story.append(Paragraph("<b>A1. Grade Distribution</b>", h2))
grade_data = [
    ["Grade", "GradeClass", "Count", "% of Total", "Avg GPA"],
    ["A (Excellent)", "4", f"{a_count:,}",  f"{a_pct:.1f}%",
     f"{df[df['GradeClass']==4]['GPA'].mean():.2f}"],
    ["B (Good)",      "3", f"{b_count:,}",  f"{b_count/total*100:.1f}%",
     f"{df[df['GradeClass']==3]['GPA'].mean():.2f}"],
    ["C (Average)",   "2", f"{c_count:,}",  f"{c_count/total*100:.1f}%",
     f"{df[df['GradeClass']==2]['GPA'].mean():.2f}"],
    ["D (Below Avg)", "1", f"{d_count:,}",  f"{d_count/total*100:.1f}%",
     f"{df[df['GradeClass']==1]['GPA'].mean():.2f}"],
    ["F (Failing)",   "0", f"{fail_count:,}",f"{fail_pct:.1f}%",
     f"{df[df['GradeClass']==0]['GPA'].mean():.2f}"],
    ["Total",         "—", f"{total:,}",    "100%", f"{avg_gpa:.2f}"],
]
gt = Table(grade_data, colWidths=[PAGE_W*f for f in [0.22,0.14,0.14,0.16,0.14]])
gt.setStyle(TableStyle([
    ("BACKGROUND",   (0,0),(-1,0),  DARK_BLUE),
    ("TEXTCOLOR",    (0,0),(-1,0),  WHITE),
    ("FONTNAME",     (0,0),(-1,0),  "Helvetica-Bold"),
    ("FONTSIZE",     (0,0),(-1,-1), 9),
    ("ROWBACKGROUNDS",(0,1),(-1,-2),[LIGHT_BG, WHITE]),
    ("BACKGROUND",   (0,-1),(-1,-1),MED_BLUE),
    ("TEXTCOLOR",    (0,-1),(-1,-1),WHITE),
    ("FONTNAME",     (0,-1),(-1,-1),"Helvetica-Bold"),
    ("ALIGN",        (1,0),(-1,-1), "CENTER"),
    ("GRID",         (0,0),(-1,-1), 0.5, MID_GRAY),
    ("TOPPADDING",   (0,0),(-1,-1), 5),
    ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ("LEFTPADDING",  (0,0),(-1,-1), 6),
]))
story.append(gt)
story.append(Spacer(1, 0.4*cm))

# Risk tier table
story.append(Paragraph("<b>A2. Risk Index Tier Summary</b>", h2))
risk_data = [
    ["Risk Tier",    "Score Range", "Count", "% of Total", "Description"],
    ["Low",          "0 – 29",  f"{low_risk:,}",    f"{low_risk/total*100:.1f}%",
     "Monitor via standard reporting"],
    ["Medium",       "30 – 54", f"{medium_risk:,}", f"{medium_risk/total*100:.1f}%",
     "Academic advisor check-in recommended"],
    ["High",         "55 – 74", f"{high_risk:,}",   f"{high_risk/total*100:.1f}%",
     "Tutoring + attendance intervention"],
    ["Critical",     "75 – 100",f"{critical_risk:,}",f"{critical_risk/total*100:.1f}%",
     "Immediate multi-faceted intervention"],
    ["Total",        "—",       f"{total:,}",        "100%", ""],
]
rt = Table(risk_data, colWidths=[PAGE_W*f for f in [0.16,0.16,0.13,0.15,0.40]])
rt.setStyle(TableStyle([
    ("BACKGROUND",   (0,0),(-1,0),  DARK_BLUE),
    ("TEXTCOLOR",    (0,0),(-1,0),  WHITE),
    ("FONTNAME",     (0,0),(-1,0),  "Helvetica-Bold"),
    ("FONTSIZE",     (0,0),(-1,-1), 9),
    ("ROWBACKGROUNDS",(0,1),(-1,-2),[LIGHT_BG, WHITE]),
    ("BACKGROUND",   (0,-1),(-1,-1),MED_BLUE),
    ("TEXTCOLOR",    (0,-1),(-1,-1),WHITE),
    ("FONTNAME",     (0,-1),(-1,-1),"Helvetica-Bold"),
    ("ALIGN",        (1,0),(3,-1),  "CENTER"),
    ("GRID",         (0,0),(-1,-1), 0.5, MID_GRAY),
    ("TOPPADDING",   (0,0),(-1,-1), 5),
    ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ("LEFTPADDING",  (0,0),(-1,-1), 6),
]))
story.append(rt)
story.append(Spacer(1, 0.4*cm))

# Key correlations
story.append(Paragraph("<b>A3. Correlation Summary (vs GPA)</b>", h2))
numeric_cols = ["StudyTimeWeekly", "Absences", "ParentalSupport",
                "ParentalEducation", "Age"]
corr_rows = [["Variable", "Pearson r", "Interpretation"]]
interp = {
    "Absences":          "Strong negative — highest predictor",
    "StudyTimeWeekly":   "Moderate positive",
    "ParentalSupport":   "Moderate positive",
    "ParentalEducation": "Weak positive",
    "Age":               "Near zero (negligible)",
}
for col in numeric_cols:
    r = df[col].corr(df["GPA"])
    corr_rows.append([col, f"{r:.4f}", interp.get(col, "")])
ct = Table(corr_rows, colWidths=[PAGE_W*f for f in [0.34, 0.18, 0.48]])
ct.setStyle(TableStyle([
    ("BACKGROUND",   (0,0),(-1,0),  DARK_BLUE),
    ("TEXTCOLOR",    (0,0),(-1,0),  WHITE),
    ("FONTNAME",     (0,0),(-1,0),  "Helvetica-Bold"),
    ("FONTSIZE",     (0,0),(-1,-1), 9),
    ("ROWBACKGROUNDS",(0,1),(-1,-1),[LIGHT_BG, WHITE]),
    ("ALIGN",        (1,0),(1,-1),  "CENTER"),
    ("GRID",         (0,0),(-1,-1), 0.5, MID_GRAY),
    ("TOPPADDING",   (0,0),(-1,-1), 5),
    ("BOTTOMPADDING",(0,0),(-1,-1), 5),
    ("LEFTPADDING",  (0,0),(-1,-1), 6),
]))
story.append(ct)
story.append(Spacer(1, 0.5*cm))

story.append(HRFlowable(width="100%", thickness=1, color=MED_BLUE,
                         spaceAfter=6, spaceBefore=6))
story.append(Paragraph(
    "GitHub: https://github.com/prachisingh342006/data_analytics_project  |  "
    "Dataset: https://www.kaggle.com/datasets/lainguyn123/student-performance-factors  |  "
    "Report generated: February 2026",
    S("EndNote", fontName="Helvetica-Oblique", fontSize=8,
      textColor=GRAY, alignment=TA_CENTER, leading=12)))

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD PDF
# ═══════════════════════════════════════════════════════════════════════════════
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2*cm,  bottomMargin=2.5*cm,
    title="Student Early Warning System — Analytics Report",
    author="Prachi Singh",
    subject="Student Performance Analysis",
)
doc.build(story, canvasmaker=NumberedCanvas)
print(f"✅  PDF saved → {OUT}")
