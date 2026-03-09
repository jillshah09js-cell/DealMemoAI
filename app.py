"""
DealMemo AI — v2.0
Professional IC Memo Generator for Indian Startup Ecosystem
Built for: Angel Syndicates, Category 1/2 AIFs, Family Offices, Micro VC Funds
"""

import os
import json
import time
import re
import streamlit as st
from groq import Groq
import PyPDF2
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_TABLE_ALIGNMENT
from io import BytesIO
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from dotenv import load_dotenv

load_dotenv()

# ─── API CONFIGURATION ──────────────────────────────────────────────────────
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
client = Groq(api_key=GROQ_API_KEY)

# Primary model — best quality on Groq free tier
PRIMARY_MODEL   = "llama-3.3-70b-versatile"
# Fallback model — in case primary hits rate limits
FALLBACK_MODEL  = "llama3-70b-8192"

# ─── COLOUR PALETTE (used in charts + Word doc) ──────────────────────────────
NAVY        = "1F3864"   # Section headings
DARK_BLUE   = "1F4E79"   # Table headers
MID_BLUE    = "2E75B6"   # Accent
LIGHT_BLUE  = "D6E4F0"   # Alternating rows
GREEN       = "1E8449"   # Positive / Strong
AMBER       = "D68910"   # Medium / Acceptable
RED         = "C0392B"   # Negative / Weak / High Risk
GREY        = "7F8C8D"   # Muted text
WHITE       = "FFFFFF"

# ─── SAFE AI CALL — NEVER CRASHES ───────────────────────────────────────────
def safe_ai_call(messages: list, expect_json: bool = False,
                 max_tokens: int = 1200) -> str:
    """
    Wraps every Groq API call with:
    - Primary model attempt
    - Automatic fallback to secondary model on rate limit
    - JSON extraction and validation if expect_json=True
    - Graceful error string on total failure (never raises)
    """
    for model in [PRIMARY_MODEL, FALLBACK_MODEL]:
        try:
            response = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.25 if expect_json else 0.4,
                max_tokens=max_tokens,
            )
            raw = response.choices[0].message.content.strip()

            if expect_json:
                # Strip markdown code fences if present
                raw = re.sub(r"```(?:json)?", "", raw).strip()
                # Validate — if invalid, return empty structure
                try:
                    json.loads(raw)
                    return raw
                except json.JSONDecodeError:
                    # Try to extract JSON substring
                    match = re.search(r"\{[\s\S]*\}", raw)
                    if match:
                        candidate = match.group(0)
                        json.loads(candidate)   # validate
                        return candidate
                    return "{}"
            return raw

        except Exception as e:
            err = str(e)
            if "rate_limit" in err.lower() or "429" in err:
                time.sleep(8)
                continue          # try fallback model
            # Non-rate-limit error — return graceful placeholder
            return (
                '{"error": "Could not generate this section. '
                'Please retry."}' if expect_json
                else "[Section could not be generated — please retry.]"
            )
    return (
        '{"error": "Rate limit reached. Please wait 60s and retry."}'
        if expect_json
        else "[Rate limit reached — please wait 60 seconds and retry.]"
    )


# ─── PDF TEXT EXTRACTION ─────────────────────────────────────────────────────
def extract_pdf_text(pdf_file) -> str:
    try:
        reader = PyPDF2.PdfReader(pdf_file)
        pages = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pages.append(text.strip())
        return "\n\n".join(pages)[:12000]
    except Exception:
        return ""


# ════════════════════════════════════════════════════════════════════════════
#  PROMPT LIBRARY — Each prompt is engineered for maximum output quality
# ════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are a Senior Investment Analyst with 9 years of 
experience at a top-tier Indian VC fund (think Sequoia India, Accel India, 
Lightspeed India). You have personally reviewed and written IC memos for 
300+ deals across fintech, edtech, D2C, SaaS, healthtech, agritech, and 
consumer internet in India.

Your writing is:
- Precise, direct, and analytical — never fluffy or promotional
- Deeply calibrated to Indian market realities (INR figures, Indian 
  regulatory context, Indian consumer behaviour, Indian competitive 
  landscape)
- Honest about weaknesses — you flag red flags clearly
- Referenced against real Indian startup benchmarks (Zepto, Groww, CRED, 
  Razorpay, Meesho, PhonePe, Swiggy, Zomato, Ola, Byju's, Nykaa, Mamaearth)
- Formatted for an Investment Committee — not a blog post

When information is missing from the deck you write exactly:
[DATA UNAVAILABLE — Recommend requesting from management]

You NEVER make up financial figures. You NEVER use marketing language.
You ALWAYS flag founder/team red flags if present."""


def prompt_executive_summary(deck: str, company: str, fund_type: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Write the EXECUTIVE SUMMARY & INVESTMENT THESIS section of a 
professional Investment Committee memo.

COMPANY: {company}
MEMO TYPE: {fund_type}

PITCH DECK CONTENT:
---
{deck[:4000]}
---

STRUCTURE YOUR RESPONSE AS FOLLOWS (use these exact subheadings in bold):

**The Opportunity**
2-3 sentences. What is the company, what problem does it solve, and why 
does this problem exist at scale in India today.

**Investment Thesis**
3-4 sentences. The core reason to invest. What has to be true for this 
to be a great outcome. What is the bull case.

**Key Metrics at a Glance**
List the 4-5 most important numbers from the deck (revenue, growth rate, 
GMV, users, retention — whatever is most relevant). If not available, 
flag as [DATA UNAVAILABLE].

**Why Now**
2-3 sentences. Why is this the right time for this company to exist and 
to raise capital. Market timing, regulatory tailwinds, technology shifts.

**Preliminary Recommendation**
One clear sentence: Proceed to full diligence / Pass / Conditional proceed.
Give a confidence level: High / Medium / Low.

Length: 280-320 words total. No bullet points except in Key Metrics.
"""


def prompt_company_overview(deck: str, company: str, fund_type: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Write the COMPANY OVERVIEW & BUSINESS MODEL section of a professional 
IC memo.

COMPANY: {company}
MEMO TYPE: {fund_type}

PITCH DECK CONTENT:
---
{deck[:4000]}
---

STRUCTURE YOUR RESPONSE AS FOLLOWS:

**Business Description**
What the company does in plain language. Product or service. How it works 
end-to-end. Target customer segment.

**Revenue Model**
How the company makes money. All revenue streams. Pricing model. Take rate 
if marketplace. Is the model proven or hypothetical.

**Stage & Traction**
Current stage (idea / MVP / early revenue / scaling). Key traction metrics. 
What has been validated vs what is still assumption.

**Operational Model**
Key operational levers. Unit of production. What scales and what doesn't. 
Geographic focus and expansion plan.

**Regulatory Considerations**
Any SEBI, RBI, DPIIT, FSSAI, or other Indian regulatory factors that 
affect this business model. Flag any compliance risks.

Length: 260-300 words. Short paragraphs only. No bullet points.
"""


def prompt_market_opportunity(deck: str, company: str, fund_type: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Write the MARKET OPPORTUNITY section of a professional IC memo.

COMPANY: {company}
MEMO TYPE: {fund_type}

PITCH DECK CONTENT:
---
{deck[:4000]}
---

STRUCTURE YOUR RESPONSE AS FOLLOWS:

**Market Context**
2-3 sentences on why this market is interesting in India right now. 
Macro tailwinds, demographic shifts, digital penetration trends.

**TAM / SAM / SOM Analysis**
Provide your own analyst estimate of TAM, SAM, and SOM using a 
bottom-up approach where possible. Use INR figures. Cross-check 
any figures from the deck against your own analysis. If the 
company's TAM claim appears inflated, say so explicitly.

**Market Structure**
Is this a fragmented or consolidated market. Who are the incumbent 
players. Is there a dominant leader or is the market still open.

**Analyst Assessment**
Your honest view: Is this a large enough market for a venture-scale 
outcome. What share does the company need to capture to return the fund.

Length: 260-300 words. TAM/SAM/SOM can use numbers inline.
"""


def prompt_investment_recommendation(deck: str, company: str,
                                     fund_type: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Write the INVESTMENT RECOMMENDATION section — the final and most 
critical section of the IC memo.

COMPANY: {company}
MEMO TYPE: {fund_type}

PITCH DECK CONTENT:
---
{deck[:4000]}
---

STRUCTURE YOUR RESPONSE AS FOLLOWS:

**Recommendation**
State clearly: PROCEED TO FULL DILIGENCE / CONDITIONAL PROCEED / PASS
For PROCEED: State the proposed check size and ownership target.
For CONDITIONAL: State exactly what conditions must be met.
For PASS: State the primary reason.

**Bull Case**
If everything goes right — what does this look like in 5 years. 
Comparable outcome (e.g. "If this follows Zepto's trajectory..."). 
Expected return multiple.

**Bear Case**
The most realistic failure scenario. What kills this company. 
Probability of bear case in your assessment.

**Key Diligence Questions**
The 5 most important questions the IC must answer before committing. 
Be specific — not generic questions but questions specific to this deal.

**Conditions & Covenants**
Any protective provisions, information rights, board seat requirements, 
or milestone-based tranches you would recommend.

Length: 300-340 words. This section must read like a decisive analyst 
recommendation — not a "on one hand / on the other hand" hedge.
"""


def prompt_competitive_table(deck: str, company: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Generate a competitive landscape analysis for {company}.

PITCH DECK CONTENT:
---
{deck[:3500]}
---

Return ONLY a valid JSON object. No text before or after. No markdown.
No code fences. Just the raw JSON.

The JSON must follow this exact structure:

{{
  "rows": [
    {{
      "company": "Competitor company name",
      "founded": "Year founded",
      "funding": "Total funding raised (INR or USD)",
      "product_focus": "Core product in 5 words",
      "business_model": "Revenue model in 4 words",
      "india_presence": "Yes / No / Limited",
      "key_strength": "Their strongest competitive advantage",
      "key_weakness": "Their most exploitable weakness",
      "threat_level": "High / Medium / Low"
    }}
  ]
}}

Rules:
- Include 4-5 real competitors. Research your training knowledge carefully.
- The LAST row must be {company} itself for self-comparison.
- Mark {company}'s row with company name exactly as "{company} (Subject)"
- Prioritise India-based or India-focused competitors.
- threat_level = High means this competitor directly threatens the subject.
- Be honest and analytical. Do not favour the subject company.
- All values must be strings. No null values — use "N/A" if unknown.
"""


def prompt_team_table(deck: str, company: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Assess the founding and leadership team of {company}.

PITCH DECK CONTENT:
---
{deck[:3500]}
---

Return ONLY a valid JSON object. No text before or after. No markdown.

{{
  "rows": [
    {{
      "name": "Full name",
      "title": "Job title",
      "prior_company": "Most impressive prior employer",
      "prior_role": "Prior role title",
      "relevant_experience": "Why this experience is directly relevant",
      "education": "Highest relevant degree and institution",
      "flag": "Green / Yellow / Red",
      "flag_reason": "One sentence explaining the flag"
    }}
  ],
  "team_summary": "2 sentence overall team assessment. Be honest.",
  "missing_roles": "Key C-suite or functional roles not yet hired"
}}

Flag guidance:
- Green: Strong pedigree, directly relevant experience, proven operator
- Yellow: Adequate but gaps exist — flag the gap specifically  
- Red: Concerning — no relevant experience, first-time founder in 
  complex regulated space, or key info not disclosed

If a team member is not mentioned in the deck, add a row:
name = "[Not Disclosed]", flag = "Red", 
flag_reason = "Key role not mentioned in deck"

All values must be strings. No null values.
"""


def prompt_financials_table(deck: str, company: str, fund_type: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Extract and assess unit economics and financial metrics for {company}.

PITCH DECK CONTENT:
---
{deck[:3500]}
---

Return ONLY a valid JSON object. No text before or after. No markdown.

{{
  "metrics": [
    {{
      "metric": "Metric name",
      "reported_value": "Value from deck or [Not Disclosed]",
      "analyst_note": "Your assessment of this figure",
      "india_benchmark": "Typical range for this metric in Indian startups",
      "assessment": "Strong / Acceptable / Weak / Not Disclosed"
    }}
  ],
  "financial_summary": "3 sentence analyst assessment of overall financial health and quality of metrics disclosed.",
  "key_concern": "The single most important financial concern for the IC."
}}

Always include these metrics in this order (use [Not Disclosed] if absent):
1. Monthly Revenue Run Rate (INR)
2. Month-on-Month Revenue Growth
3. Gross Margin %
4. Customer Acquisition Cost (CAC)
5. Customer Lifetime Value (LTV)
6. LTV / CAC Ratio
7. Monthly Burn Rate (INR)
8. Runway (months)
9. Monthly Active Users / Customers
10. Net Revenue Retention or Churn Rate

All values must be strings. No null values.
"""


def prompt_risk_table(deck: str, company: str, fund_type: str) -> str:
    return f"""
{SYSTEM_PROMPT}

TASK: Identify and assess key investment risks for {company}.

PITCH DECK CONTENT:
---
{deck[:3500]}
---

Return ONLY a valid JSON object. No text before or after. No markdown.

{{
  "rows": [
    {{
      "risk_category": "Category (Market/Execution/Team/Regulatory/Financial/Technology/Competition)",
      "risk_title": "Short title (5 words max)",
      "risk_description": "Specific risk description for this company (not generic)",
      "probability": "High / Medium / Low",
      "impact": "High / Medium / Low",
      "time_horizon": "Near-term (0-12m) / Medium-term (1-3y) / Long-term (3y+)",
      "mitigant": "Specific mitigating factor or action",
      "residual_risk": "High / Medium / Low (after mitigation)"
    }}
  ],
  "risk_summary": "2 sentence overall risk assessment. Most critical risk and why."
}}

Include exactly 6 risks covering:
1. One market risk (TAM, timing, macro)
2. One execution risk (ops, scaling, product)
3. One team/founder risk (be honest — even if team looks strong)
4. One regulatory/compliance risk (India-specific)
5. One competitive risk (specific named competitor threat)
6. One financial risk (burn, runway, fundraising)

Be specific to this company — generic risks are not useful.
All values must be strings. No null values.
"""


# ════════════════════════════════════════════════════════════════════════════
#  CHART GENERATION — Embedded in Word document
# ════════════════════════════════════════════════════════════════════════════

def make_tam_chart(tam_data: dict, company: str) -> BytesIO:
    """
    Creates a professional TAM/SAM/SOM funnel bar chart.
    Returns PNG image as BytesIO.
    """
    fig, ax = plt.subplots(figsize=(7, 3.5))
    fig.patch.set_facecolor("#F8FAFC")
    ax.set_facecolor("#F8FAFC")

    labels = ["TAM\n(Total Addressable Market)",
              "SAM\n(Serviceable Addressable Market)",
              "SOM\n(Serviceable Obtainable Market)"]
    values = [
        tam_data.get("tam_inr", 50000),
        tam_data.get("sam_inr", 15000),
        tam_data.get("som_inr", 3000),
    ]
    colours = ["#1F4E79", "#2E75B6", "#9DC3E6"]

    bars = ax.barh(labels, values, color=colours, height=0.5, edgecolor="white")

    # Value labels
    for bar, val in zip(bars, values):
        unit = "Cr" if val >= 100 else "L"
        display = f"₹{val:,.0f} {unit}"
        ax.text(bar.get_width() + max(values) * 0.01,
                bar.get_y() + bar.get_height() / 2,
                display, va="center", fontsize=9,
                color="#1F3864", fontweight="bold")

    ax.set_xlabel("INR (Crores)", fontsize=9, color="#7F8C8D")
    ax.set_title(f"{company} — Market Sizing",
                 fontsize=11, fontweight="bold", color="#1F3864", pad=10)
    ax.tick_params(colors="#5D6D7E", labelsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#BDC3C7")
    ax.spines["bottom"].set_color("#BDC3C7")
    ax.set_xlim(0, max(values) * 1.25)

    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf


def make_risk_heatmap(risk_rows: list) -> BytesIO:
    """
    Creates a professional risk heatmap (probability vs impact matrix).
    Returns PNG image as BytesIO.
    """
    fig, ax = plt.subplots(figsize=(6, 4))
    fig.patch.set_facecolor("#F8FAFC")
    ax.set_facecolor("#FDFEFE")

    # Background quadrant colours
    quadrants = [
        (0, 0, 1, 1, "#FDFEFE"),    # Low/Low  — white
        (1, 0, 1, 1, "#FEF9E7"),    # High/Low — light yellow
        (0, 1, 1, 1, "#FEF9E7"),    # Low/High — light yellow
        (1, 1, 1, 1, "#FDEDEC"),    # High/High — light red
    ]
    for x, y, w, h, c in quadrants:
        ax.add_patch(mpatches.Rectangle((x, y), w, h, color=c, zorder=0))

    level_map = {"Low": 0.5, "Medium": 1.5, "High": 2.5}
    color_map = {"High": "#C0392B", "Medium": "#D68910", "Low": "#1E8449"}

    plotted = []
    for i, row in enumerate(risk_rows[:6]):
        prob  = row.get("probability", "Medium")
        impact = row.get("impact", "Medium")
        px = level_map.get(prob, 1.5)
        py = level_map.get(impact, 1.5)

        # Slight offset if overlapping
        offset = 0.08 * sum(1 for p in plotted if abs(p[0]-px) < 0.2
                            and abs(p[1]-py) < 0.2)
        px += offset
        py += offset
        plotted.append((px, py))

        col = color_map.get(prob, "#7F8C8D")
        ax.scatter(px, py, s=180, color=col, zorder=5, edgecolors="white",
                   linewidths=1.5)
        ax.annotate(
            row.get("risk_title", f"Risk {i+1}"),
            (px, py), textcoords="offset points", xytext=(6, 4),
            fontsize=7, color="#1F3864", fontweight="bold",
            bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                      edgecolor="#BDC3C7", alpha=0.85)
        )

    # Axis config
    ax.set_xlim(0, 3)
    ax.set_ylim(0, 3)
    ax.set_xticks([0.5, 1.5, 2.5])
    ax.set_yticks([0.5, 1.5, 2.5])
    ax.set_xticklabels(["Low", "Medium", "High"], fontsize=9, color="#5D6D7E")
    ax.set_yticklabels(["Low", "Medium", "High"], fontsize=9, color="#5D6D7E")
    ax.set_xlabel("Probability", fontsize=9, color="#7F8C8D", labelpad=8)
    ax.set_ylabel("Impact", fontsize=9, color="#7F8C8D", labelpad=8)
    ax.set_title("Risk Heatmap — Probability vs Impact",
                 fontsize=10, fontweight="bold", color="#1F3864", pad=10)
    ax.grid(True, linestyle="--", alpha=0.3, color="#BDC3C7")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf


def make_financial_bar(metrics: list) -> BytesIO:
    """
    Creates a visual assessment bar chart for financial metrics.
    Shows Strong / Acceptable / Weak / Not Disclosed status.
    """
    assessment_color = {
        "Strong":       "#1E8449",
        "Acceptable":   "#D68910",
        "Weak":         "#C0392B",
        "Not Disclosed":"#7F8C8D",
    }

    names  = [m.get("metric", f"Metric {i}")[:28]
              for i, m in enumerate(metrics[:10])]
    colors = [assessment_color.get(m.get("assessment", "Not Disclosed"),
                                   "#7F8C8D") for m in metrics[:10]]
    values = [1] * len(names)

    fig, ax = plt.subplots(figsize=(7, len(names) * 0.52 + 1))
    fig.patch.set_facecolor("#F8FAFC")
    ax.set_facecolor("#F8FAFC")

    bars = ax.barh(names, values, color=colors, height=0.6,
                   edgecolor="white", linewidth=0.5)

    # Labels inside bars
    for bar, m in zip(bars, metrics[:10]):
        assessment = m.get("assessment", "Not Disclosed")
        val        = m.get("reported_value", "N/A")
        display    = f"  {val}  ({assessment})"
        ax.text(0.02, bar.get_y() + bar.get_height() / 2,
                display, va="center", fontsize=8,
                color="white", fontweight="bold")

    ax.set_xlim(0, 1)
    ax.set_xticks([])
    ax.set_title("Unit Economics Assessment",
                 fontsize=11, fontweight="bold", color="#1F3864", pad=10)
    ax.tick_params(colors="#5D6D7E", labelsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_visible(False)

    # Legend
    legend_patches = [
        mpatches.Patch(color="#1E8449", label="Strong"),
        mpatches.Patch(color="#D68910", label="Acceptable"),
        mpatches.Patch(color="#C0392B", label="Weak"),
        mpatches.Patch(color="#7F8C8D", label="Not Disclosed"),
    ]
    ax.legend(handles=legend_patches, loc="lower right",
              fontsize=7, framealpha=0.8)

    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════════════════════════
#  WORD DOCUMENT BUILDER — Professional IC Memo formatting
# ════════════════════════════════════════════════════════════════════════════

def hex_to_rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color.lstrip("#"))
    tcPr.append(shd)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom),
                      ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def add_section_heading(doc, text: str, level: int = 1):
    """Add a styled section heading with a coloured bottom border."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(14)
    para.paragraph_format.space_after  = Pt(4)
    run  = para.add_run(text.upper())
    run.bold = True
    run.font.size  = Pt(11) if level == 1 else Pt(10)
    run.font.color.rgb = hex_to_rgb(NAVY)
    run.font.name  = "Calibri"

    # Bottom border line under heading
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), MID_BLUE)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return para


def add_body_text(doc, text: str):
    """Add a styled body paragraph."""
    para = doc.add_paragraph(text)
    para.paragraph_format.space_after  = Pt(6)
    para.paragraph_format.space_before = Pt(2)
    for run in para.runs:
        run.font.size = Pt(10)
        run.font.name = "Calibri"
    return para


def add_styled_table(doc, headers: list, rows: list,
                     col_widths: list = None,
                     stripe_color: str = LIGHT_BLUE):
    """
    Add a fully styled table with:
    - Dark blue header row (white text)
    - Alternating row shading
    - Proper cell padding
    - Consistent font
    """
    n_cols = len(headers)
    table  = doc.add_table(rows=1, cols=n_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths (in DXA — 1440 DXA = 1 inch)
    if col_widths:
        for i, width in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = Pt(width) * 20  # approximate

    # Header row
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = hdr_cells[i]
        set_cell_bg(cell, DARK_BLUE)
        set_cell_margins(cell)
        para = cell.paragraphs[0]
        para.clear()
        run = para.add_run(header)
        run.bold = True
        run.font.size  = Pt(9)
        run.font.color.rgb = hex_to_rgb(WHITE)
        run.font.name  = "Calibri"
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Data rows
    for row_idx, row_data in enumerate(rows):
        row = table.add_row()
        bg  = stripe_color if row_idx % 2 == 0 else WHITE
        for col_idx, value in enumerate(row_data):
            cell = row.cells[col_idx]
            # Override bg for color-coded cells
            if isinstance(value, tuple):
                text, override_color = value
                set_cell_bg(cell, override_color)
                set_cell_margins(cell)
                para = cell.paragraphs[0]
                para.clear()
                run  = para.add_run(str(text))
                run.bold = True
                run.font.size  = Pt(9)
                run.font.color.rgb = hex_to_rgb(WHITE)
                run.font.name  = "Calibri"
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                set_cell_bg(cell, bg)
                set_cell_margins(cell)
                para = cell.paragraphs[0]
                para.clear()
                run  = para.add_run(str(value))
                run.font.size = Pt(9)
                run.font.name = "Calibri"

    doc.add_paragraph()
    return table


def color_for_level(level: str, invert: bool = False) -> str:
    """Return hex color for High/Medium/Low risk levels."""
    mapping = {"High": RED, "Medium": AMBER, "Low": GREEN}
    return mapping.get(level, GREY)


def embed_chart(doc, chart_buf: BytesIO, width_inches: float = 6.0):
    """Embed a matplotlib chart PNG into the Word document."""
    doc.add_picture(chart_buf, width=Inches(width_inches))
    last_para = doc.paragraphs[-1]
    last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()


def build_word_memo(text_sections: dict, table_data: dict,
                    company: str, fund_type: str) -> BytesIO:
    doc = Document()

    # ── Page setup ───────────────────────────────────────────────────────────
    section = doc.sections[0]
    section.top_margin    = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin   = Inches(1.0)
    section.right_margin  = Inches(1.0)

    # ── Cover header strip ───────────────────────────────────────────────────
    # Confidential banner
    conf = doc.add_paragraph()
    conf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    conf.paragraph_format.space_before = Pt(0)
    conf.paragraph_format.space_after  = Pt(4)
    conf_run = conf.add_run("STRICTLY CONFIDENTIAL — FOR IC USE ONLY")
    conf_run.font.size  = Pt(8)
    conf_run.font.name  = "Calibri"
    conf_run.font.color.rgb = hex_to_rgb(GREY)

    # Main title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("INVESTMENT COMMITTEE MEMORANDUM")
    title_run.bold = True
    title_run.font.size  = Pt(18)
    title_run.font.name  = "Calibri"
    title_run.font.color.rgb = hex_to_rgb(NAVY)

    # Company name
    co_para = doc.add_paragraph()
    co_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    co_run = co_para.add_run(company.upper())
    co_run.bold = True
    co_run.font.size  = Pt(15)
    co_run.font.name  = "Calibri"
    co_run.font.color.rgb = hex_to_rgb(MID_BLUE)

    doc.add_paragraph()

    # Metadata table
    meta_rows = [
        ["Memo Type",       fund_type],
        ["Classification",  "Confidential — Investment Committee Only"],
        ["Prepared By",     "DealMemo AI  |  Senior Analyst Review"],
        ["Status",          "First-Pass Analysis — Full Diligence Pending"],
    ]
    meta_table = doc.add_table(rows=len(meta_rows), cols=2)
    meta_table.style = "Table Grid"
    for i, (label, value) in enumerate(meta_rows):
        meta_table.cell(i, 0).text = label
        meta_table.cell(i, 1).text = value
        set_cell_bg(meta_table.cell(i, 0), LIGHT_BLUE)
        set_cell_margins(meta_table.cell(i, 0))
        set_cell_margins(meta_table.cell(i, 1))
        for run in meta_table.cell(i, 0).paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(9)
            run.font.name = "Calibri"
        for run in meta_table.cell(i, 1).paragraphs[0].runs:
            run.font.size = Pt(9)
            run.font.name = "Calibri"

    doc.add_paragraph()
    doc.add_page_break()

    # ── SECTION 1: Executive Summary ─────────────────────────────────────────
    add_section_heading(doc, "1. Executive Summary & Investment Thesis")
    add_body_text(doc, text_sections.get("executive_summary", ""))
    doc.add_paragraph()

    # ── SECTION 2: Company Overview ──────────────────────────────────────────
    add_section_heading(doc, "2. Company Overview & Business Model")
    add_body_text(doc, text_sections.get("company_overview", ""))
    doc.add_paragraph()

    # ── SECTION 3: Market Opportunity ────────────────────────────────────────
    add_section_heading(doc, "3. Market Opportunity")
    add_body_text(doc, text_sections.get("market_opportunity", ""))
    doc.add_paragraph()

    # TAM Chart
    if "tam_chart" in table_data:
        embed_chart(doc, table_data["tam_chart"], width_inches=5.5)

    # ── SECTION 4: Competitive Landscape ─────────────────────────────────────
    add_section_heading(doc, "4. Competitive Landscape")
    comp_data = table_data.get("competitive", {})
    comp_rows_raw = comp_data.get("rows", [])

    if comp_rows_raw:
        headers = ["Company", "Founded", "Funding",
                   "Product Focus", "Business Model",
                   "India Presence", "Key Strength",
                   "Key Weakness", "Threat"]
        rows = []
        for r in comp_rows_raw:
            threat = r.get("threat_level", "Medium")
            rows.append([
                r.get("company", ""),
                r.get("founded", "N/A"),
                r.get("funding", "N/A"),
                r.get("product_focus", ""),
                r.get("business_model", ""),
                r.get("india_presence", ""),
                r.get("key_strength", ""),
                r.get("key_weakness", ""),
                (threat, color_for_level(threat)),
            ])
        add_styled_table(doc, headers, rows)

    # ── SECTION 5: Team Assessment ───────────────────────────────────────────
    add_section_heading(doc, "5. Team Assessment")
    team_data = table_data.get("team", {})
    team_rows_raw = team_data.get("rows", [])
    team_summary  = team_data.get("team_summary", "")
    missing_roles = team_data.get("missing_roles", "")

    if team_rows_raw:
        headers = ["Name", "Title", "Prior Company",
                   "Prior Role", "Relevant Experience",
                   "Education", "Assessment"]
        rows = []
        for r in team_rows_raw:
            flag = r.get("flag", "Yellow")
            rows.append([
                r.get("name", ""),
                r.get("title", ""),
                r.get("prior_company", ""),
                r.get("prior_role", ""),
                r.get("relevant_experience", ""),
                r.get("education", ""),
                (flag, color_for_level(
                    "Low" if flag == "Green"
                    else "Medium" if flag == "Yellow"
                    else "High"
                )),
            ])
        add_styled_table(doc, headers, rows)

    if team_summary:
        add_body_text(doc, f"Team Summary: {team_summary}")
    if missing_roles:
        add_body_text(doc, f"Missing Roles: {missing_roles}")

    # ── SECTION 6: Unit Economics & Financials ───────────────────────────────
    add_section_heading(doc, "6. Unit Economics & Financial Analysis")
    fin_data    = table_data.get("financials", {})
    metrics_raw = fin_data.get("metrics", [])
    fin_summary = fin_data.get("financial_summary", "")
    key_concern = fin_data.get("key_concern", "")

    if metrics_raw:
        headers = ["Metric", "Reported Value",
                   "Analyst Note", "India Benchmark", "Assessment"]
        rows = []
        for m in metrics_raw:
            assessment = m.get("assessment", "Not Disclosed")
            color_map  = {
                "Strong":       GREEN,
                "Acceptable":   AMBER,
                "Weak":         RED,
                "Not Disclosed": GREY,
            }
            rows.append([
                m.get("metric", ""),
                m.get("reported_value", "N/A"),
                m.get("analyst_note", ""),
                m.get("india_benchmark", ""),
                (assessment, color_map.get(assessment, GREY)),
            ])
        add_styled_table(doc, headers, rows)

    # Financial assessment chart
    if metrics_raw and "fin_chart" in table_data:
        embed_chart(doc, table_data["fin_chart"], width_inches=6.0)

    if fin_summary:
        add_body_text(doc, f"Financial Assessment: {fin_summary}")
    if key_concern:
        add_body_text(doc, f"⚠  Key Concern: {key_concern}")

    # ── SECTION 7: Key Risks ──────────────────────────────────────────────────
    add_section_heading(doc, "7. Key Risks & Mitigants")
    risk_data    = table_data.get("risks", {})
    risk_rows_raw = risk_data.get("rows", [])
    risk_summary  = risk_data.get("risk_summary", "")

    if risk_rows_raw:
        headers = ["Category", "Risk", "Description",
                   "Probability", "Impact",
                   "Time Horizon", "Mitigant", "Residual Risk"]
        rows = []
        for r in risk_rows_raw:
            prob    = r.get("probability", "Medium")
            impact  = r.get("impact", "Medium")
            residual = r.get("residual_risk", "Medium")
            rows.append([
                r.get("risk_category", ""),
                r.get("risk_title", ""),
                r.get("risk_description", ""),
                (prob,     color_for_level(prob)),
                (impact,   color_for_level(impact)),
                r.get("time_horizon", ""),
                r.get("mitigant", ""),
                (residual, color_for_level(residual)),
            ])
        add_styled_table(doc, headers, rows)

    # Risk heatmap chart
    if risk_rows_raw and "risk_chart" in table_data:
        embed_chart(doc, table_data["risk_chart"], width_inches=5.0)

    if risk_summary:
        add_body_text(doc, f"Risk Summary: {risk_summary}")

    # ── SECTION 8: Investment Recommendation ──────────────────────────────────
    doc.add_page_break()
    add_section_heading(doc, "8. Investment Recommendation")
    add_body_text(doc, text_sections.get("recommendation", ""))

    # ── Footer disclaimer ─────────────────────────────────────────────────────
    doc.add_paragraph()
    disc_para = doc.add_paragraph()
    disc_para.paragraph_format.space_before = Pt(12)
    line = disc_para.add_run("─" * 90)
    line.font.color.rgb = hex_to_rgb(GREY)
    line.font.size = Pt(8)

    disclaimer = doc.add_paragraph()
    d_run = disclaimer.add_run(
        "DISCLAIMER: This memorandum was prepared using DealMemo AI as a "
        "first-draft analytical tool. All figures, assessments, and "
        "recommendations must be independently verified by a qualified "
        "investment professional before any capital commitment. This "
        "document does not constitute financial advice or a solicitation "
        "to invest. Past performance of comparable companies is not "
        "indicative of future results."
    )
    d_run.font.size  = Pt(8)
    d_run.font.color.rgb = hex_to_rgb(GREY)
    d_run.font.name  = "Calibri"

    # ── Save ──────────────────────────────────────────────────────────────────
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════════════════════════
#  TAM DATA EXTRACTION — Used to power the market chart
# ════════════════════════════════════════════════════════════════════════════

def extract_tam_data(deck: str, company: str) -> dict:
    """Extract TAM/SAM/SOM figures from deck text for chart generation."""
    result = safe_ai_call(
        messages=[
            {"role": "system",
             "content": "You are a JSON data extractor. Return only valid JSON."},
            {"role": "user", "content": f"""
Extract or estimate TAM, SAM, SOM figures for {company} from this deck.
Return ONLY this JSON (no text, no markdown):

{{
  "tam_inr": <number in INR Crores, integer only>,
  "sam_inr": <number in INR Crores, integer only>,
  "som_inr": <number in INR Crores, integer only>
}}

If not explicitly stated, use your analyst knowledge to estimate.
SAM must be less than TAM. SOM must be less than SAM.

Deck content:
{deck[:2000]}
"""}
        ],
        expect_json=True,
        max_tokens=120
    )
    try:
        data = json.loads(result)
        # Sanity check
        if data.get("tam_inr", 0) > 0:
            return data
    except Exception:
        pass
    return {"tam_inr": 50000, "sam_inr": 12000, "som_inr": 2500}


# ════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="DealMemo AI",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for cleaner UI
st.markdown("""
<style>
    .main { max-width: 1100px; }
    .stButton button {
        background-color: #1F4E79;
        color: white;
        font-weight: 600;
        border-radius: 4px;
        border: none;
        padding: 0.6rem 1.2rem;
    }
    .stButton button:hover { background-color: #2E75B6; }
    .metric-box {
        background: #F0F4F8;
        border-left: 4px solid #1F4E79;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 8px 0;
    }
    h1 { color: #1F3864 !important; }
    h2 { color: #1F4E79 !important; }
</style>
""", unsafe_allow_html=True)

# Header
st.title("📋 DealMemo AI")
st.subheader("Professional IC Memo Generator — Indian Startup Ecosystem")
st.markdown(
    "Upload a pitch deck and get a **full Investment Committee memo** — "
    "with tables, charts, and structured analysis — calibrated for "
    "Indian funds."
)

st.divider()

# Input area
col1, col2 = st.columns([1, 1])

with col1:
    company = st.text_input(
        "Company Name",
        placeholder="e.g. Zomato, Groww, YourStartup",
        help="Enter the exact company name as it appears in the deck."
    )
    fund_type = st.selectbox(
        "Memo Format",
        ["Angel Syndicate",
         "Category 1 AIF",
         "Category 2 AIF",
         "Family Office",
         "Micro VC Fund",
         "PE / Growth Equity"]
    )

with col2:
    pdf_file = st.file_uploader(
        "Upload Pitch Deck (PDF)",
        type=["pdf"],
        help="Upload the startup's pitch deck as a PDF file."
    )
    if pdf_file:
        st.success(f"✅ Uploaded: {pdf_file.name}")
    st.caption(
        "🔒 Your deck is never stored or shared. "
        "It is processed in memory only."
    )

st.divider()

# What gets generated info
with st.expander("📄 What this memo includes", expanded=False):
    st.markdown("""
    **8 Sections:**
    1. Executive Summary & Investment Thesis
    2. Company Overview & Business Model
    3. Market Opportunity (with TAM/SAM/SOM chart)
    4. Competitive Landscape (colour-coded comparison matrix)
    5. Team Assessment (with Green/Yellow/Red flags)
    6. Unit Economics & Financial Analysis (with visual assessment)
    7. Key Risks & Mitigants (with risk heatmap)
    8. Investment Recommendation (with bull/bear case)

    **Output:** Professional Word document (.docx) ready for IC presentation.
    """)

# Generate button
generate = st.button(
    "🚀 Generate IC Memo",
    type="primary",
    use_container_width=True
)

if generate:
    if not company:
        st.error("⚠️ Please enter the company name.")
    elif not pdf_file:
        st.error("⚠️ Please upload a pitch deck PDF.")
    elif not GROQ_API_KEY:
        st.error("⚠️ API key not configured. Contact support.")
    else:
        # Extract PDF
        deck_text = extract_pdf_text(pdf_file)

        if not deck_text.strip():
            st.error(
                "❌ Could not extract text from this PDF. "
                "Please ensure it is not a scanned image-only PDF. "
                "Try copy-pasting text from the PDF to verify."
            )
        else:
            st.info(
                f"⚡ Generating IC memo for **{company}**... "
                "Takes about 3-4 minutes. Do not close this tab."
            )

            # Progress tracking
            progress  = st.progress(0)
            status    = st.empty()
            TOTAL_STEPS = 11  # 4 text + 4 tables + 3 charts

            text_sections = {}
            table_data    = {}
            step = 0

            def tick(msg):
                global step
                step += 1
                progress.progress(min(step / TOTAL_STEPS, 1.0))
                status.text(msg)

            # ── Text Sections ─────────────────────────────────────────────
            tick("✍️  Writing Executive Summary...")
            text_sections["executive_summary"] = safe_ai_call(
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",
                     "content": prompt_executive_summary(
                         deck_text, company, fund_type)}
                ],
                max_tokens=1000
            )
            time.sleep(3)

            tick("✍️  Writing Company Overview...")
            text_sections["company_overview"] = safe_ai_call(
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",
                     "content": prompt_company_overview(
                         deck_text, company, fund_type)}
                ],
                max_tokens=900
            )
            time.sleep(3)

            tick("✍️  Writing Market Opportunity...")
            text_sections["market_opportunity"] = safe_ai_call(
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",
                     "content": prompt_market_opportunity(
                         deck_text, company, fund_type)}
                ],
                max_tokens=900
            )
            time.sleep(3)

            tick("✍️  Writing Investment Recommendation...")
            text_sections["recommendation"] = safe_ai_call(
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",
                     "content": prompt_investment_recommendation(
                         deck_text, company, fund_type)}
                ],
                max_tokens=1000
            )
            time.sleep(3)

            # ── Table Sections ────────────────────────────────────────────
            tick("📊 Building Competitive Matrix...")
            raw = safe_ai_call(
                messages=[
                    {"role": "system",
                     "content": "You output only valid JSON. No text."},
                    {"role": "user",
                     "content": prompt_competitive_table(deck_text, company)}
                ],
                expect_json=True,
                max_tokens=1200
            )
            try:
                table_data["competitive"] = json.loads(raw)
            except Exception:
                table_data["competitive"] = {"rows": []}
            time.sleep(3)

            tick("👥 Building Team Assessment Table...")
            raw = safe_ai_call(
                messages=[
                    {"role": "system",
                     "content": "You output only valid JSON. No text."},
                    {"role": "user",
                     "content": prompt_team_table(deck_text, company)}
                ],
                expect_json=True,
                max_tokens=1200
            )
            try:
                table_data["team"] = json.loads(raw)
            except Exception:
                table_data["team"] = {"rows": []}
            time.sleep(3)

            tick("💰 Building Financial Analysis Table...")
            raw = safe_ai_call(
                messages=[
                    {"role": "system",
                     "content": "You output only valid JSON. No text."},
                    {"role": "user",
                     "content": prompt_financials_table(
                         deck_text, company, fund_type)}
                ],
                expect_json=True,
                max_tokens=1400
            )
            try:
                table_data["financials"] = json.loads(raw)
            except Exception:
                table_data["financials"] = {"metrics": []}
            time.sleep(3)

            tick("⚠️  Building Risk Matrix...")
            raw = safe_ai_call(
                messages=[
                    {"role": "system",
                     "content": "You output only valid JSON. No text."},
                    {"role": "user",
                     "content": prompt_risk_table(
                         deck_text, company, fund_type)}
                ],
                expect_json=True,
                max_tokens=1400
            )
            try:
                table_data["risks"] = json.loads(raw)
            except Exception:
                table_data["risks"] = {"rows": []}
            time.sleep(2)

            # ── Charts ────────────────────────────────────────────────────
            tick("📈 Generating TAM/SAM/SOM Chart...")
            tam_data = extract_tam_data(deck_text, company)
            table_data["tam_chart"] = make_tam_chart(tam_data, company)

            tick("📊 Generating Financial Assessment Chart...")
            metrics = table_data.get("financials", {}).get("metrics", [])
            if metrics:
                table_data["fin_chart"] = make_financial_bar(metrics)

            tick("🗺️  Generating Risk Heatmap...")
            risk_rows = table_data.get("risks", {}).get("rows", [])
            if risk_rows:
                table_data["risk_chart"] = make_risk_heatmap(risk_rows)

            # ── Done ──────────────────────────────────────────────────────
            progress.progress(1.0)
            status.text("✅ IC Memo Complete!")
            st.success(f"🎉 IC Memo for **{company}** is ready!")
            st.divider()

            # ── Preview ───────────────────────────────────────────────────
            st.subheader(f"Preview — {company} IC Memo")

            tab1, tab2, tab3, tab4 = st.tabs([
                "📝 Text Sections",
                "📊 Tables",
                "📈 Charts",
                "⚠️  Risks"
            ])

            with tab1:
                for title, key in [
                    ("Executive Summary & Investment Thesis",
                     "executive_summary"),
                    ("Company Overview & Business Model",
                     "company_overview"),
                    ("Market Opportunity",
                     "market_opportunity"),
                    ("Investment Recommendation",
                     "recommendation"),
                ]:
                    with st.expander(f"📌 {title}", expanded=False):
                        st.write(text_sections.get(key, ""))

            with tab2:
                st.markdown("**Competitive Landscape**")
                comp_rows = table_data.get(
                    "competitive", {}).get("rows", [])
                if comp_rows:
                    st.dataframe(comp_rows, use_container_width=True)

                st.markdown("**Team Assessment**")
                team_rows = table_data.get("team", {}).get("rows", [])
                if team_rows:
                    st.dataframe(team_rows, use_container_width=True)

                st.markdown("**Unit Economics**")
                fin_metrics = table_data.get(
                    "financials", {}).get("metrics", [])
                if fin_metrics:
                    st.dataframe(fin_metrics, use_container_width=True)

            with tab3:
                if "tam_chart" in table_data:
                    st.image(table_data["tam_chart"],
                             caption="TAM / SAM / SOM",
                             use_container_width=True)
                if "fin_chart" in table_data:
                    table_data["fin_chart"].seek(0)
                    st.image(table_data["fin_chart"],
                             caption="Unit Economics Assessment",
                             use_container_width=True)

            with tab4:
                risk_rows = table_data.get("risks", {}).get("rows", [])
                if risk_rows:
                    st.dataframe(risk_rows, use_container_width=True)
                if "risk_chart" in table_data:
                    table_data["risk_chart"].seek(0)
                    st.image(table_data["risk_chart"],
                             caption="Risk Heatmap",
                             use_container_width=True)

            st.divider()

            # ── Download ──────────────────────────────────────────────────
            # Reset chart buffers for Word doc
            for key in ["tam_chart", "fin_chart", "risk_chart"]:
                if key in table_data:
                    table_data[key].seek(0)

            word_file = build_word_memo(
                text_sections, table_data, company, fund_type
            )

            st.download_button(
                label="📥 Download Full IC Memo — Word Document",
                data=word_file,
                file_name=f"IC_Memo_{company.replace(' ', '_')}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                ),
                use_container_width=True
            )

            st.caption(
                "💡 **Next step:** Review the memo, edit any sections "
                "where you have additional context, and present to IC. "
                "Always verify figures independently before investment decisions."
            )
