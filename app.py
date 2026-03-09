import os
from dotenv import load_dotenv
load_dotenv()
import json
import time
import streamlit as st
from groq import Groq
import PyPDF2
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use("Agg")

# ─── API CONFIGURATION ──────────────────────────────────────
import os
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
client = Groq(api_key=GROQ_API_KEY)
MODEL = "llama-3.3-70b-versatile"

# ─── SECTION TYPES ──────────────────────────────────────────
# "text" sections → written paragraphs
# "table" sections → structured table output
TEXT_SECTIONS = [
    "Executive Summary & Investment Thesis",
    "Company Overview & Business Model",
    "Market Opportunity & TAM/SAM/SOM",
    "Investment Recommendation"
]

TABLE_SECTIONS = [
    "Competitive Landscape",
    "Team Assessment",
    "Unit Economics & Financial Projections",
    "Key Risks & Mitigants"
]

# ─── EXTRACT TEXT FROM PDF ──────────────────────────────────
def extract_pdf_text(pdf_file):
    reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    return text[:8000]

# ─── GENERATE TEXT SECTION ──────────────────────────────────
def generate_text_section(deck_text, section, company, fund_type):
    prompt = f"""
    You are a Senior Investment Analyst with 8 years of experience
    writing Investment Committee memos for Indian VC and angel funds.
    You have reviewed 200+ startup deals.

    Write the '{section}' section of a professional IC memo.

    Company: {company}
    Memo Format: {fund_type}

    Pitch Deck Content:
    ---
    {deck_text[:3000]}
    ---

    Requirements:
    - Indian market context and INR figures where relevant
    - Reference comparable Indian startups where relevant
    - Direct, analytical, professional tone — no fluff
    - Flag missing info as: [Not available — recommend requesting]
    - Length: 200-250 words
    - Short paragraphs only. No bullet points.
    - Do not write the section title
    - Do not add preamble like "Here is the section"
    """

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a Senior Investment Analyst at a top Indian "
                    "VC fund. You write precise, analytical IC memos. "
                    "Never use fluff. Always flag missing information."
                )
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.3,
        max_tokens=600
    )
    return response.choices[0].message.content

# ─── GENERATE COMPETITIVE LANDSCAPE TABLE ───────────────────
def generate_competitive_table(deck_text, company, fund_type):
    prompt = f"""
    You are a Senior Investment Analyst writing an IC memo for {company}.

    Pitch Deck Content:
    ---
    {deck_text[:3000]}
    ---

    Create a competitive landscape analysis.
    Return ONLY a JSON object in this exact format, nothing else:

    {{
      "rows": [
        {{
          "company": "Company name",
          "product_focus": "What they do",
          "pricing": "Pricing model",
          "india_presence": "Yes or No",
          "key_weakness": "Main weakness",
          "threat_level": "High / Medium / Low"
        }}
      ]
    }}

    Rules:
    - Include {company} as the LAST row, label it clearly
    - Include 4-5 real competitors, India-focused where possible
    - Be honest and analytical
    - Return valid JSON only, no extra text
    """

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are a JSON data generator. Return only valid JSON, nothing else."
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.1,
        max_tokens=800
    )

    raw = response.choices[0].message.content.strip()
    # Clean up in case model adds markdown
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ─── GENERATE TEAM TABLE ────────────────────────────────────
def generate_team_table(deck_text, company):
    prompt = f"""
    You are a Senior Investment Analyst writing an IC memo for {company}.

    Pitch Deck Content:
    ---
    {deck_text[:3000]}
    ---

    Extract the founding team and key hires from this pitch deck.
    Return ONLY a JSON object in this exact format, nothing else:

    {{
      "rows": [
        {{
          "name": "Full name",
          "role": "Job title",
          "background": "Previous company or experience",
          "relevant_experience": "Why they are right for this role",
          "flag": "Green / Yellow / Red"
        }}
      ]
    }}

    Rules:
    - Flag meaning: Green = strong fit, Yellow = some gaps, Red = concern
    - If a team member is missing from the deck write the name as
      [Not disclosed] and flag as Red
    - Be honest about gaps
    - Return valid JSON only, no extra text
    """

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are a JSON data generator. Return only valid JSON, nothing else."
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.1,
        max_tokens=800
    )

    raw = response.choices[0].message.content.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ─── GENERATE RISK MATRIX TABLE ─────────────────────────────
def generate_risk_table(deck_text, company, fund_type):
    prompt = f"""
    You are a Senior Investment Analyst writing an IC memo for {company}.

    Pitch Deck Content:
    ---
    {deck_text[:3000]}
    ---

    Identify the key investment risks for this deal.
    Return ONLY a JSON object in this exact format, nothing else:

    {{
      "rows": [
        {{
          "risk": "Risk title",
          "description": "One sentence description",
          "probability": "High / Medium / Low",
          "impact": "High / Medium / Low",
          "mitigant": "How this risk can be mitigated"
        }}
      ]
    }}

    Rules:
    - Include 5-6 most important risks
    - Be specific to this company and sector
    - Include at least one market risk, one execution risk,
      one regulatory risk, and one team or founder risk
    - Return valid JSON only, no extra text
    """

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are a JSON data generator. Return only valid JSON, nothing else."
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.1,
        max_tokens=800
    )

    raw = response.choices[0].message.content.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ─── GENERATE UNIT ECONOMICS TABLE ──────────────────────────
def generate_financials_table(deck_text, company, fund_type):
    prompt = f"""
    You are a Senior Investment Analyst writing an IC memo for {company}.

    Pitch Deck Content:
    ---
    {deck_text[:3000]}
    ---

    Extract or estimate key unit economics and financial metrics.
    Return ONLY a JSON object in this exact format, nothing else:

    {{
      "metrics": [
        {{
          "metric": "Metric name",
          "value": "Current value or estimate",
          "benchmark": "Industry benchmark for India",
          "assessment": "Strong / Acceptable / Weak / Not Disclosed"
        }}
      ]
    }}

    Rules:
    - Include these metrics if available: CAC, LTV, LTV/CAC ratio,
      Gross Margin, Monthly Burn Rate, Runway, MoM Growth Rate,
      Revenue Run Rate, Payback Period, NPS or Retention Rate
    - Use INR figures where possible
    - If a metric is not in the deck mark value as [Not Disclosed]
    - Return valid JSON only, no extra text
    """

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are a JSON data generator. Return only valid JSON, nothing else."
            },
            {"role": "user", "content": prompt}
        ],
        temperature=0.1,
        max_tokens=800
    )

    raw = response.choices[0].message.content.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ─── WORD TABLE HELPER FUNCTIONS ────────────────────────────
def set_cell_background(cell, hex_color):
    """Set background color of a Word table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def add_header_row(table, headers, header_color="1F4E79"):
    """Add a styled header row to a Word table."""
    header_row = table.rows[0]
    for i, header_text in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header_text
        set_cell_background(cell, header_color)
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)

def style_data_cell(cell, font_size=9, bold=False, color=None):
    """Style a data cell in a Word table."""
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.size = Pt(font_size)
            if bold:
                run.bold = True
            if color:
                run.font.color.rgb = color

def get_risk_color(level):
    """Return hex color based on risk level."""
    colors = {
        "High": "FF4444",
        "Medium": "FFA500",
        "Low": "00AA00"
    }
    return colors.get(level, "888888")

def get_flag_color(flag):
    """Return hex color based on team flag."""
    colors = {
        "Green": "00AA00",
        "Yellow": "FFA500",
        "Red": "FF4444"
    }
    return colors.get(flag, "888888")

# ─── ADD COMPETITIVE TABLE TO DOC ───────────────────────────
def add_competitive_table(doc, data):
    headers = [
        "Company", "Product Focus", "Pricing",
        "India Presence", "Key Weakness", "Threat Level"
    ]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    add_header_row(table, headers)

    for row_data in data["rows"]:
        row = table.add_row()
        values = [
            row_data.get("company", ""),
            row_data.get("product_focus", ""),
            row_data.get("pricing", ""),
            row_data.get("india_presence", ""),
            row_data.get("key_weakness", ""),
            row_data.get("threat_level", "")
        ]
        for i, value in enumerate(values):
            row.cells[i].text = value
            style_data_cell(row.cells[i])

        # Color the threat level cell
        threat = row_data.get("threat_level", "")
        threat_cell = row.cells[5]
        color = get_risk_color(threat)
        set_cell_background(threat_cell, color)
        for para in threat_cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

# ─── ADD TEAM TABLE TO DOC ──────────────────────────────────
def add_team_table(doc, data):
    headers = [
        "Name", "Role", "Background",
        "Relevant Experience", "Assessment"
    ]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    add_header_row(table, headers)

    for row_data in data["rows"]:
        row = table.add_row()
        values = [
            row_data.get("name", ""),
            row_data.get("role", ""),
            row_data.get("background", ""),
            row_data.get("relevant_experience", ""),
            row_data.get("flag", "")
        ]
        for i, value in enumerate(values):
            row.cells[i].text = value
            style_data_cell(row.cells[i])

        # Color the flag cell
        flag = row_data.get("flag", "")
        flag_cell = row.cells[4]
        color = get_flag_color(flag)
        set_cell_background(flag_cell, color)
        for para in flag_cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

# ─── ADD RISK TABLE TO DOC ──────────────────────────────────
def add_risk_table(doc, data):
    headers = [
        "Risk", "Description",
        "Probability", "Impact", "Mitigant"
    ]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    add_header_row(table, headers)

    for row_data in data["rows"]:
        row = table.add_row()
        values = [
            row_data.get("risk", ""),
            row_data.get("description", ""),
            row_data.get("probability", ""),
            row_data.get("impact", ""),
            row_data.get("mitigant", "")
        ]
        for i, value in enumerate(values):
            row.cells[i].text = value
            style_data_cell(row.cells[i])

        # Color probability and impact cells
        prob = row_data.get("probability", "")
        impact = row_data.get("impact", "")

        prob_cell = row.cells[2]
        set_cell_background(prob_cell, get_risk_color(prob))
        for para in prob_cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

        impact_cell = row.cells[3]
        set_cell_background(impact_cell, get_risk_color(impact))
        for para in impact_cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

# ─── ADD FINANCIALS TABLE TO DOC ────────────────────────────
def add_financials_table(doc, data):
    headers = [
        "Metric", "Current Value",
        "India Benchmark", "Assessment"
    ]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    add_header_row(table, headers)

    assessment_colors = {
        "Strong": "00AA00",
        "Acceptable": "FFA500",
        "Weak": "FF4444",
        "Not Disclosed": "888888"
    }

    for row_data in data["metrics"]:
        row = table.add_row()
        assessment = row_data.get("assessment", "")
        values = [
            row_data.get("metric", ""),
            row_data.get("value", ""),
            row_data.get("benchmark", ""),
            assessment
        ]
        for i, value in enumerate(values):
            row.cells[i].text = value
            style_data_cell(row.cells[i])

        # Bold the metric name
        row.cells[0].paragraphs[0].runs[0].bold = True

        # Color the assessment cell
        color = assessment_colors.get(assessment, "888888")
        set_cell_background(row.cells[3], color)
        for para in row.cells[3].paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

# ─── BUILD FULL WORD DOCUMENT ───────────────────────────────
def build_word_memo(text_sections, table_data, company, fund_type):
    doc = Document()

    # Page margins
    section_props = doc.sections[0]
    section_props.top_margin = Pt(72)
    section_props.bottom_margin = Pt(72)
    section_props.left_margin = Pt(80)
    section_props.right_margin = Pt(80)

    # Main title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run("INVESTMENT COMMITTEE MEMORANDUM")
    title_run.bold = True
    title_run.font.size = Pt(16)

    # Company name
    company_para = doc.add_paragraph()
    company_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    company_run = company_para.add_run(company.upper())
    company_run.bold = True
    company_run.font.size = Pt(14)
    company_run.font.color.rgb = RGBColor(31, 78, 121)

    doc.add_paragraph()

    # Metadata table
    meta_table = doc.add_table(rows=3, cols=2)
    meta_table.style = "Table Grid"
    meta = [
        ("Memo Format", fund_type),
        ("Classification", "CONFIDENTIAL — For IC Use Only"),
        ("Prepared By", "DealMemo AI")
    ]
    for i, (label, value) in enumerate(meta):
        meta_table.cell(i, 0).text = label
        meta_table.cell(i, 1).text = value
        meta_table.cell(i, 0).paragraphs[0].runs[0].bold = True

    doc.add_paragraph()

    # ── Text sections first ──
    for section_title in TEXT_SECTIONS:
        if section_title in text_sections:
            doc.add_paragraph()
            heading = doc.add_paragraph()
            heading_run = heading.add_run(section_title.upper())
            heading_run.bold = True
            heading_run.font.size = Pt(11)
            heading_run.font.color.rgb = RGBColor(31, 78, 121)

            doc.add_paragraph(text_sections[section_title])
            doc.add_paragraph()

    # ── Table sections ──

    # Competitive Landscape
    if "competitive" in table_data:
        doc.add_paragraph()
        heading = doc.add_paragraph()
        heading.add_run("COMPETITIVE LANDSCAPE").bold = True
        heading.runs[0].font.size = Pt(11)
        heading.runs[0].font.color.rgb = RGBColor(31, 78, 121)
        doc.add_paragraph()
        add_competitive_table(doc, table_data["competitive"])
        doc.add_paragraph()

    # Team Assessment
    if "team" in table_data:
        doc.add_paragraph()
        heading = doc.add_paragraph()
        heading.add_run("TEAM ASSESSMENT").bold = True
        heading.runs[0].font.size = Pt(11)
        heading.runs[0].font.color.rgb = RGBColor(31, 78, 121)
        doc.add_paragraph()
        add_team_table(doc, table_data["team"])
        doc.add_paragraph()

    # Unit Economics
    if "financials" in table_data:
        doc.add_paragraph()
        heading = doc.add_paragraph()
        heading.add_run(
            "UNIT ECONOMICS & FINANCIAL PROJECTIONS"
        ).bold = True
        heading.runs[0].font.size = Pt(11)
        heading.runs[0].font.color.rgb = RGBColor(31, 78, 121)
        doc.add_paragraph()
        add_financials_table(doc, table_data["financials"])
        doc.add_paragraph()

    # Risk Matrix
    if "risks" in table_data:
        doc.add_paragraph()
        heading = doc.add_paragraph()
        heading.add_run("KEY RISKS & MITIGANTS").bold = True
        heading.runs[0].font.size = Pt(11)
        heading.runs[0].font.color.rgb = RGBColor(31, 78, 121)
        doc.add_paragraph()
        add_risk_table(doc, table_data["risks"])
        doc.add_paragraph()

    # Investment Recommendation — always last
    if "Investment Recommendation" in text_sections:
        doc.add_paragraph()
        heading = doc.add_paragraph()
        heading.add_run("INVESTMENT RECOMMENDATION").bold = True
        heading.runs[0].font.size = Pt(11)
        heading.runs[0].font.color.rgb = RGBColor(31, 78, 121)
        doc.add_paragraph(
            text_sections["Investment Recommendation"]
        )

    # Footer
    doc.add_paragraph()
    doc.add_paragraph("─" * 80)
    disclaimer = doc.add_paragraph(
        "DISCLAIMER: This memo was prepared using DealMemo AI as a "
        "first-draft research tool. All analysis must be independently "
        "verified before any investment decision is made. This document "
        "does not constitute financial advice."
    )
    disclaimer.runs[0].font.size = Pt(9)
    disclaimer.runs[0].font.color.rgb = RGBColor(128, 128, 128)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ─── STREAMLIT UI ────────────────────────────────────────────
st.set_page_config(
    page_title="DealMemo AI",
    page_icon="📋",
    layout="wide"
)

st.title("📋 DealMemo AI")
st.subheader("Investment Committee Memo Generator for Indian Funds")
st.markdown(
    "Upload any pitch deck and get a **full IC memo with tables "
    "and structured analysis** in under 5 minutes."
)
st.divider()

col1, col2 = st.columns(2)

with col1:
    company = st.text_input(
        "Company Name",
        placeholder="e.g. Zomato, Groww, YourStartup"
    )
    fund_type = st.selectbox(
        "Memo Format",
        [
            "Angel Syndicate",
            "Category 1 AIF",
            "Category 2 AIF",
            "Family Office",
            "Micro VC Fund"
        ]
    )

with col2:
    pdf_file = st.file_uploader(
        "Upload Pitch Deck (PDF only)",
        type="pdf"
    )
    st.caption(
        "Your pitch deck is never stored. "
        "Used only to generate this memo."
    )

st.divider()

if st.button(
    "🚀 Generate IC Memo", type="primary", use_container_width=True
):
    if not company:
        st.error("Please enter the company name.")
    elif not pdf_file:
        st.error("Please upload a pitch deck PDF.")
    else:
        deck_text = extract_pdf_text(pdf_file)

        if not deck_text.strip():
            st.error(
                "Could not extract text from this PDF. "
                "Please ensure it is not a scanned image."
            )
        else:
            st.info(
                "Generating your IC memo with tables... "
                "Takes about 3 minutes. Do not close this tab."
            )

            progress_bar = st.progress(0)
            status_text = st.empty()
            total_steps = len(TEXT_SECTIONS) + len(TABLE_SECTIONS)
            step = 0

            # ── Generate text sections ──
            text_sections = {}
            for section in TEXT_SECTIONS:
                status_text.text(f"⚡ Writing: {section}...")
                text_sections[section] = generate_text_section(
                    deck_text, section, company, fund_type
                )
                step += 1
                progress_bar.progress(step / total_steps)
                time.sleep(2)

            # ── Generate table sections ──
            table_data = {}

            status_text.text("📊 Building Competitive Matrix...")
            table_data["competitive"] = generate_competitive_table(
                deck_text, company, fund_type
            )
            step += 1
            progress_bar.progress(step / total_steps)
            time.sleep(2)

            status_text.text("👥 Building Team Assessment Table...")
            table_data["team"] = generate_team_table(
                deck_text, company
            )
            step += 1
            progress_bar.progress(step / total_steps)
            time.sleep(2)

            status_text.text(
                "💰 Building Unit Economics Table..."
            )
            table_data["financials"] = generate_financials_table(
                deck_text, company, fund_type
            )
            step += 1
            progress_bar.progress(step / total_steps)
            time.sleep(2)

            status_text.text("⚠️ Building Risk Matrix...")
            table_data["risks"] = generate_risk_table(
                deck_text, company, fund_type
            )
            step += 1
            progress_bar.progress(step / total_steps)
            time.sleep(2)

            status_text.text("✅ All sections complete!")
            st.success("Your IC Memo is ready!")
            st.divider()

            # ── Preview on screen ──
            st.subheader(f"IC Memo Preview — {company}")

            for section_title in TEXT_SECTIONS:
                with st.expander(
                    f"📌 {section_title}", expanded=True
                ):
                    st.write(text_sections.get(section_title, ""))

            with st.expander(
                "📊 Competitive Landscape", expanded=True
            ):
                rows = table_data["competitive"].get("rows", [])
                if rows:
                    st.table(rows)

            with st.expander(
                "👥 Team Assessment", expanded=True
            ):
                rows = table_data["team"].get("rows", [])
                if rows:
                    st.table(rows)

            with st.expander(
                "💰 Unit Economics & Financials", expanded=True
            ):
                metrics = table_data["financials"].get(
                    "metrics", []
                )
                if metrics:
                    st.table(metrics)

            with st.expander(
                "⚠️ Risk Matrix", expanded=True
            ):
                rows = table_data["risks"].get("rows", [])
                if rows:
                    st.table(rows)

            st.divider()

            # ── Download button ──
            word_file = build_word_memo(
                text_sections, table_data, company, fund_type
            )
            st.download_button(
                label="📥 Download IC Memo as Word Document",
                data=word_file,
                file_name=f"IC_Memo_{company}.docx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".wordprocessingml.document",
                use_container_width=True
            )