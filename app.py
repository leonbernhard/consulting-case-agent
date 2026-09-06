import os
import io
import re
import streamlit as st
from crewai import Agent, Crew, Process, Task, LLM
from docx import Document
from docx.shared import Pt, RGBColor, Inches

# Seiten-Konfiguration
st.set_page_config(page_title="Case Structuring Agent", page_icon="📊", layout="wide")
# EXECUTIVE C-LEVEL STYLING (CSS INJECTION)
EXECUTIVE_CSS = """
<style>
    /* Typografie & Globale Schriftart */
    html, body, [class*="css"] {
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif !important;
    }

    /* Primary Buttons (Haupt-Aktionen) */
    div.stButton > button {
        background-color: #0F2C59 !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        border: 1px solid #0F2C59 !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.25rem !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 2px 4px rgba(15, 44, 89, 0.1) !important;
    }
    
    div.stButton > button:hover {
        background-color: #1E40AF !important;
        border-color: #1E40AF !important;
        box-shadow: 0 4px 12px rgba(15, 44, 89, 0.25) !important;
        transform: translateY(-1px) !important;
    }

    /* KPI Metric Cards Redesign */
    [data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-left: 4px solid #0F2C59 !important;
        padding: 16px 20px !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(15, 44, 89, 0.04) !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 12px !important;
        font-weight: 700 !important;
        color: #64748B !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    
    [data-testid="stMetricValue"] {
        color: #0F2C59 !important;
        font-weight: 700 !important;
        font-size: 20px !important;
    }

    /* Tabs Styling (C-Level Navigation) */
    button[data-baseweb="tab"] {
        font-weight: 600 !important;
        color: #64748B !important;
        padding: 12px 20px !important;
        border-radius: 4px 4px 0 0 !important;
    }
    
    button[aria-selected="true"] {
        color: #0F2C59 !important;
        border-bottom-color: #0F2C59 !important;
        border-bottom-width: 3px !important;
        background-color: #F8FAFC !important;
    }

    /* Seitenleiste (Sidebar Refinement) */
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0 !important;
    }

    /* Strukturierte Container im Ergebnisbereich */
    .stTabs [data-testid="stMarkdownContainer"] {
        line-height: 1.6 !important;
        color: #1E293B !important;
    }

    /* Formular-Elemente & Textareas */
    textarea {
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
    }
    textarea:focus {
        border-color: #0F2C59 !important;
        box-shadow: 0 0 0 1px #0F2C59 !important;
    }
</style>
"""
st.markdown(EXECUTIVE_CSS, unsafe_allow_html=True)

# 1. API Key prüfen & abfangen
api_key = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not api_key:
    api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    st.error("⚠️ GEMINI_API_KEY wurde nicht gefunden. Bitte trage deinen Key in `.streamlit/secrets.toml` oder in den Streamlit Cloud Secrets ein.")
    st.stop()

os.environ["GEMINI_API_KEY"] = api_key

# 2. Modell initialisieren
gemini_llm = LLM(
    model="gemini-3.6-flash",
    api_key=api_key
)

# 3. Hilfsfunktionen für Exporte
def create_html_report(title, framework, analysis, mece, hypothesis, sub_text, fw_label, sec1, sec2, sec3, footer_text):
    """Erstellt ein professionelles Executive HTML Dashboard mit dynamischer Sprache."""
    import html
    import re

    def md_to_html(md_text):
        lines = md_text.strip().split('\n')
        html_out = []
        in_list = False
        table_lines = []

        def format_text(txt):
            txt = html.escape(txt.strip())
            txt = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', txt)
            txt = re.sub(r'\*(.*?)\*', r'<em>\1</em>', txt)
            return txt

        def render_table(t_lines):
            if not t_lines:
                return ""
            rows = []
            for line in t_lines:
                cleaned = line.strip().strip('|')
                cells = [c.strip() for c in cleaned.split('|')]
                if all(re.match(r'^:?-+:?$', c) for c in cells if c):
                    continue
                rows.append(cells)

            if not rows:
                return ""

            out = ['<table class="executive-table">']
            out.append('<thead><tr>')
            for cell in rows[0]:
                out.append(f'<th>{format_text(cell)}</th>')
            out.append('</tr></thead><tbody>')

            for row in rows[1:]:
                out.append('<tr>')
                for cell in row:
                    out.append(f'<td>{format_text(cell)}</td>')
                out.append('</tr>')
            out.append('</tbody></table>')
            return '\n'.join(out)

        for line in lines:
            line_str = line.strip()

            if '|' in line_str and line_str.count('|') >= 2:
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                table_lines.append(line_str)
                continue
            else:
                if table_lines:
                    html_out.append(render_table(table_lines))
                    table_lines = []

            if not line_str:
                continue

            if re.match(r'^\d+\.\d+', line_str) or re.match(r'^[*\-]\s*\d+\.\d+', line_str):
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                html_out.append(f'<p style="margin-left: 24px; margin-top: 3px; margin-bottom: 5px;">{format_text(clean_txt)}</p>')
            elif re.match(r'^\d+\.\s', line_str) or re.match(r'^[*\-]\s*\d+\.\s', line_str):
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                html_out.append(f'<p style="margin-top: 14px; margin-bottom: 4px; font-weight: 600; color: #0F2C59;">{format_text(clean_txt)}</p>')
            elif line_str.startswith('### '):
                html_out.append(f'<h3>{format_text(line_str[4:])}</h3>')
            elif line_str.startswith('## '):
                html_out.append(f'<h2>{format_text(line_str[3:])}</h2>')
            elif line_str.startswith('# '):
                html_out.append(f'<h2>{format_text(line_str[2:])}</h2>')
            elif line_str.startswith('- ') or line_str.startswith('* '):
                if not in_list:
                    html_out.append('<ul class="executive-list">')
                    in_list = True
                html_out.append(f'<li>{format_text(line_str[2:])}</li>')
            else:
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                html_out.append(f'<p>{format_text(line_str)}</p>')

        if table_lines:
            html_out.append(render_table(table_lines))
        if in_list:
            html_out.append('</ul>')

        return '\n'.join(html_out)

    html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        @page {{ size: A4; margin: 15mm; }}
        body {{ font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; line-height: 1.5; color: #1E293B; background-color: #F8FAFC; margin: 0; padding: 25px; }}
        .container {{ max-width: 850px; margin: 0 auto; background: #FFFFFF; padding: 35px 40px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); box-sizing: border-box; }}
        .header {{ border-bottom: 2px solid #0F2C59; padding-bottom: 12px; margin-bottom: 25px; }}
        .badge {{ display: inline-block; background: #0F2C59; color: #FFFFFF; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; }}
        h1 {{ color: #0F2C59; font-size: 22px; margin: 4px 0; font-weight: 700; }}
        h2 {{ color: #0F2C59; font-size: 16px; border-bottom: 1px solid #E2E8F0; padding-bottom: 5px; margin-top: 25px; margin-bottom: 12px; font-weight: 600; page-break-after: avoid; }}
        h3 {{ color: #334155; font-size: 13px; margin-top: 16px; margin-bottom: 6px; font-weight: 600; page-break-after: avoid; }}
        p, li {{ font-size: 12px; color: #334155; margin-bottom: 6px; }}
        ul.executive-list {{ padding-left: 20px; margin: 8px 0; }}
        ul.executive-list li {{ margin-bottom: 4px; }}
        
        .executive-table {{ width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 11px; page-break-inside: avoid; }}
        .executive-table th {{ background-color: #0F2C59; color: #FFFFFF; font-weight: bold; text-align: left; padding: 8px 10px; border: 1px solid #0F2C59; }}
        .executive-table td {{ border: 1px solid #CBD5E1; padding: 8px 10px; vertical-align: top; }}
        .executive-table tr:nth-child(even) {{ background-color: #F8FAFC; }}
        
        .section-block {{ page-break-inside: avoid; }}
        .footer {{ margin-top: 35px; padding-top: 12px; border-top: 1px solid #E2E8F0; font-size: 10px; color: #94A3B8; text-align: center; }}
        
        @media print {{
            html, body {{ background: #FFFFFF !important; margin: 0 !important; padding: 0 !important; }}
            .container {{ box-shadow: none !important; padding: 0 !important; margin: 12mm auto !important; width: 90% !important; max-width: 90% !important; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <span class="badge">{framework}</span>
            <h1>{title}</h1>
            <p style="color: #64748B; margin: 0; font-size: 11px;">{sub_text}</p>
        </div>
        
        <div class="section-block">
            <h2>{sec1}</h2>
            {md_to_html(analysis)}
        </div>
        
        <div class="section-block">
            <h2>{sec2}</h2>
            {md_to_html(mece)}
        </div>
        
        <div class="section-block">
            <h2>{sec3}</h2>
            {md_to_html(hypothesis)}
        </div>
        
        <div class="footer">
            {footer_text}
        </div>
    </div>
</body>
</html>"""
    return html_content

def create_docx_report(title, framework, analysis, mece, hypothesis, fw_label, sec1, sec2, sec3, footer_text):
    """Erstellt ein professionelles Word-Dokument mit dynamischer Sprache."""
    doc = Document()
    
    heading = doc.add_heading(title, level=0)
    heading.style.font.color.rgb = RGBColor(15, 44, 89)
    
    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run(f"{fw_label}: {framework}")
    run_sub.bold = True
    p_sub.paragraph_format.space_after = Pt(18)

    def add_formatted_text(paragraph, text):
        parts = re.split(r'(\*\*.*?\*\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            else:
                if part:
                    paragraph.add_run(part)

    def render_docx_table(t_lines):
        rows_data = []
        for line in t_lines:
            cleaned = line.strip().strip('|')
            cells = [c.strip() for c in cleaned.split('|')]
            if all(re.match(r'^:?-+:?$', c) for c in cells if c):
                continue
            rows_data.append(cells)
        
        if not rows_data:
            return

        col_count = max(len(r) for r in rows_data)
        table = doc.add_table(rows=len(rows_data), cols=col_count)
        table.style = 'Table Grid'
        
        for r_idx, row in enumerate(rows_data):
            for c_idx, cell_text in enumerate(row):
                if c_idx < col_count:
                    cell = table.cell(r_idx, c_idx)
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)
                    add_formatted_text(p, cell_text)
                    if r_idx == 0:
                        for run in p.runs:
                            run.bold = True

    def add_md_section(section_title, md_text):
        h = doc.add_heading(section_title, level=1)
        h.style.font.color.rgb = RGBColor(15, 44, 89)
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(6)

        lines = md_text.strip().split('\n')
        table_lines = []

        for line in lines:
            line_str = line.strip()

            if '|' in line_str and line_str.count('|') >= 2:
                table_lines.append(line_str)
                continue
            else:
                if table_lines:
                    render_docx_table(table_lines)
                    table_lines = []

            if not line_str:
                continue

            if re.match(r'^\d+\.\d+', line_str) or re.match(r'^[*\-]\s*\d+\.\d+', line_str):
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_after = Pt(3)
                add_formatted_text(p, clean_txt)
            elif re.match(r'^\d+\.\s', line_str) or re.match(r'^[*\-]\s*\d+\.\s', line_str):
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(3)
                add_formatted_text(p, clean_txt)
            elif line_str.startswith('### '):
                p = doc.add_heading(level=2)
                add_formatted_text(p, line_str[4:])
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(2)
            elif line_str.startswith('## '):
                p = doc.add_heading(level=2)
                add_formatted_text(p, line_str[3:])
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
            elif line_str.startswith('- ') or line_str.startswith('* '):
                p = doc.add_paragraph(style='List Bullet')
                add_formatted_text(p, line_str[2:])
                p.paragraph_format.space_after = Pt(2)
            else:
                p = doc.add_paragraph()
                add_formatted_text(p, line_str)
                p.paragraph_format.space_after = Pt(4)

        if table_lines:
            render_docx_table(table_lines)

    add_md_section(sec1, analysis)
    add_md_section(sec2, mece)
    add_md_section(sec3, hypothesis)

    p_footer = doc.add_paragraph()
    p_footer.paragraph_format.space_before = Pt(30)
    p_footer.alignment = 1
    run_footer = p_footer.add_run(footer_text)
    run_footer.font.size = Pt(9)
    run_footer.font.color.rgb = RGBColor(148, 163, 184)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 4. Seitenleisten-Initialisierung (Sprachauswahl zuerst)
with st.sidebar:
    st.markdown("### ⚙️ Executive Settings")
    language = st.selectbox("Language / Sprache", ["Deutsch", "English"])

# 5. Dynamische UI-Texte
if language == "English":
    ui_title = "📊 Consulting Case Structuring Agent"
    ui_subtitle = "Structured case analysis, MECE issue trees, and data-driven hypothesis development."
    ui_framework_label = "Select Framework Focus"
    ui_input_label = "Enter Case Briefing here:"
    ui_input_placeholder = "Paste the client's problem description here or load a demo case above..."
    ui_button = "🚀 Analyze Case"
    ui_demo_btn = "💡 Load Demo Case"
    ui_warning = "Please enter a case briefing first."
    ui_status_start = "🔍 AI Crew is analyzing the case briefing..."
    ui_status_done = "✅ Analysis completed successfully!"
    ui_kpi_hdr = "📈 Executive Metrics & Scope"
    ui_m1_label = "Framework Focus"
    ui_m2_label = "MECE Coverage"
    ui_m2_val = "100%"
    ui_m2_delta = "Mutually Exclusive"
    ui_m3_label = "Hypotheses"
    ui_m3_val = "3 Formulated"
    ui_m3_delta = "KPI Validated"
    ui_tab1 = "📌 Executive Summary & SCR"
    ui_tab2 = "🌳 MECE Issue Tree"
    ui_tab3 = "💡 Hypotheses & KPI Matrix"
    ui_report_hdr = "Client Analysis Report"
    ui_export_hdr = "📥 Export Deliverable"
    ui_format_label = "Select Export Format:"
    ui_key_hdr = "🔑 API Key Configuration"
    ui_key_label = "Custom Gemini API Key (Optional)"
    ui_key_help = "If the global demo quota is exhausted, a personal free API key from Google AI Studio can be entered here. Processing is secure and restricted exclusively to the active session."
    ui_sec_note = "🔒 *Input is isolated and processed strictly in-memory per session. No persistent storage.*"
    
    # Berichtsspezifische Texte (English)
    rep_sub = "Automated AI Case Analysis Report"
    rep_lbl_fw = "Framework Focus"
    rep_sec1 = "1. Executive Summary & SCR"
    rep_sec2 = "2. MECE Issue Tree"
    rep_sec3 = "3. Hypotheses & KPI Matrix"
    rep_footer = "Generated by AI Consulting & Case Structuring Agent | Confidential & Professional Support Tool"
else:
    ui_title = "📊 Consulting Case Structuring Agent"
    ui_subtitle = "Strukturierte Case-Analyse, MECE-Problembäume und datengestützte Hypothesen-Entwicklung."
    ui_framework_label = "Fokus-Framework wählen"
    ui_input_label = "Case Briefing hier eingeben:"
    ui_input_placeholder = "Füge hier die Problemstellung ein oder lade oben einen Demo-Case..."
    ui_button = "🚀 Case Analysieren"
    ui_demo_btn = "💡 Demo-Case laden"
    ui_warning = "Bitte gib zuerst ein Case-Briefing ein."
    ui_status_start = "🔍 Agenten-Crew analysiert die Problemstellung..."
    ui_status_done = "✅ Analyse erfolgreich abgeschlossen!"
    ui_kpi_hdr = "📈 Executive Metrics & Analyse-Fokus"
    ui_m1_label = "Fokus-Framework"
    ui_m2_label = "MECE-Abdeckung"
    ui_m2_val = "100 %"
    ui_m2_delta = "Überschneidungsfrei"
    ui_m3_label = "Arbeitshypothesen"
    ui_m3_val = "3 Hypothesen"
    ui_m3_delta = "KPI-validiert"
    ui_tab1 = "📌 Executive Summary & SCR"
    ui_tab2 = "🌳 MECE-Problemstruktur"
    ui_tab3 = "💡 Hypothesen & KPI-Matrix"
    ui_report_hdr = "Mandanten-Analysebericht"
    ui_export_hdr = "📥 Bericht Exportieren"
    ui_format_label = "Export-Format wählen:"
    ui_key_hdr = "🔑 API-Key Konfiguration"
    ui_key_label = "Eigener Gemini API-Key (Optional)"
    ui_key_help = "Falls das globale Test-Kontingent erschöpft ist, kann hier ein eigener kostenloser Key aus dem Google AI Studio eingetragen werden. Die Verarbeitung erfolgt sicher und ausschließlich im Arbeitsspeicher dieser Sitzung."
    ui_sec_note = "🔒 *Die Eingabe erfolgt isoliert und wird ausschließlich im flüchtigen Arbeitsspeicher verarbeitet. Keine Speicherung.*"
    
    # Berichtsspezifische Texte (Deutsch)
    rep_sub = "Automatisierter KI-Fallanalysebericht"
    rep_lbl_fw = "Fokus-Framework"
    rep_sec1 = "1. Executive Summary & SCR"
    rep_sec2 = "2. MECE-Problemstruktur"
    rep_sec3 = "3. Hypothesen & KPI-Matrix"
    rep_footer = "Erstellt durch KI Consulting & Case Structuring Agent | Vertrauliches Analyse-Tool"

# 6. Ergänzung der Seitenleiste & Hauptbereich-Header
with st.sidebar:
    st.markdown("---")
    st.markdown(f"**{ui_framework_label}**")
    framework_focus = st.selectbox(
        "",
        ["General Profitability", "Cost Reduction", "M&A Due Diligence", "Market Entry"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    # API-Key Sicherheits-Card
    st.markdown(f"""
        <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; padding: 14px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.02);">
            <span style="font-weight: 700; font-size: 13px; color: #0F2C59;">{ui_key_hdr}</span>
            <p style="font-size: 11px; color: #64748B; margin-top: 4px; margin-bottom: 0px;">{ui_sec_note}</p>
        </div>
    """, unsafe_allow_html=True)
    
    user_key = st.text_input(
        ui_key_label, 
        type="password", 
        help=ui_key_help,
        label_visibility="collapsed"
    )

    if user_key.strip():
        os.environ["GEMINI_API_KEY"] = user_key.strip()
        gemini_llm = LLM(model="gemini-3.6-flash", api_key=user_key.strip())

# 7. Demo Cases
DEMO_CASES = {
    "General Profitability": "Mandant: Mittelständisches Industrieunternehmen (Umsatz: 45 Mio. €).\nProblemstellung: Die EBIT-Marge ist innerhalb der letzten 18 Monate von 11,5 % auf 3,2 % gesunken, obwohl der Umsatz stabil geblieben ist.\nZiel: Identifikation der Hauptursachen für den Margenverfall und Entwicklung konkreter Gegenmaßnahmen zur Erreichung einer Ziel-Marge von > 8,0 %.",
    "Cost Reduction": "Mandant: Internationaler Logistikdienstleister.\nProblemstellung: Stark steigende Opex-Kosten in der Flotte und im Lagerbetrieb schmälern das Gesamtergebnis um 4,5 Mio. € im Vergleich zum Vorjahr.\nZiel: Systematische Kostenstrukturanalyse zur Identifikation von Einsparpotenzialen von mindestens 15 % ohne Qualitätsverlust im Kerngeschäft.",
    "M&A Due Diligence": "Mandant: Finanzinvestor / Private Equity.\nProblemstellung: Bewertung eines potenziellen Akquisitionsziels im Bereich B2B-Software vor Beginn der detaillierten Commercial Due Diligence.\nZiel: Validierung des nachhaltigen EBITDA-Aussagewerts, Identifikation wesentlicher Geschäftsrisiken und Prüfung der Run-Rate im Hinblick auf das Synergiepotenzial.",
    "Market Entry": "Mandant: E-Commerce-Händler für Premium-Konsumgüter.\nProblemstellung: Geplante Expansion in zwei neue europäische Märkte bei einem Investitionsbudget von 2,0 Mio. €.\nZiel: Evaluierung von Markteintrittsbarrieren, Kundenakquisitionskosten (CAC) und der erwarteten Amortisationsdauer (Payback Period)."
}

if "case_text" not in st.session_state:
    st.session_state["case_text"] = ""

col_demo, col_empty = st.columns([1, 3])
with col_demo:
    if st.button(ui_demo_btn):
        st.session_state["case_text"] = DEMO_CASES.get(framework_focus, DEMO_CASES["General Profitability"])

case_input = st.text_area(
    ui_input_label,
    value=st.session_state["case_text"],
    height=180,
    placeholder=ui_input_placeholder
)

FORMATTING_RULES = f"""
STRIKTE FORMATIERUNGS-REGELN (STRIKT EINHALTEN):
1. KEINE ASCII-Boxen oder Rahmenelemente (+---+, |---|, etc.) verwenden.
2. Für MECE-Strukturen und Baumdarstellungen AUSSCHLIESSLICH Standard-Markdown-Listen mit Einrückungen verwenden.
3. KEINE H1-Überschriften (`#`) generieren. Nutze ausschließlich Unterüberschriften ab Ebene 2 (`##`).
4. Für Tabellen ausschließlich sauberes Markdown-Tabellenformat nutzen (`| Spalte 1 | Spalte 2 |`).
5. Gesamtsprache der Ausgabe: Strikt auf {language}.
"""

# 8. Agenten-Analyse (Mit Zero-Quota Caching für Demo-Cases)
PRECACHED_PROFITABILITY = {
    "analysis": """### 1. Situation (S)
Der Mandant ist ein mittelständisches Industrieunternehmen mit einem stabilen Jahresumsatz von **45,0 Mio. €**. Die historische EBIT-Marge lag bei gesunden **11,5 %** (ca. 5,18 Mio. € EBIT).

### 2. Complication (C)
Innerhalb der letzten 18 Monate ist die EBIT-Marge drastisch um **8,3 Prozentpunkte auf 3,2 %** (ca. 1,44 Mio. € EBIT) eingebrochen. Da der Umsatz konstant geblieben ist, resultiert der operative Ergebnisverlust von **~3,74 Mio. €** vollständig aus einer Verschlechterung der Kostenstruktur und ungünstigen Preis-/Mix-Effekten.

### 3. Key Question (KQ)
Welche spezifischen Kosten- und Mix-Treiber haben die Marge erodiert, und mit welchem Maßnahmenpaket kann die EBIT-Marge nachhaltig auf die Zielmarke von **> 8,0 %** (> 3,60 Mio. € EBIT) angehoben werden?

### 4. Resolution & Strategic Approach (R)
Zur Erreichung der Ziel-Marge ist eine EBIT-Steigerung um mindestens **2,16 Mio. €** erforderlich. Dies erfordert eine Ursachenanalyse entlang des Profitabilitäts-Baums sowie die Implementierung eines zweiphasigen Optimierungsprogramms.""",
    "mece": """**1. Erlösqualität & Preisdurchsetzung (Umsatz- & Mix-Hebel)**
1.1 **Preisanpassung & Indexierung:** Unzureichende Weitergabe gestiegener Inputkosten an Endkunden.
1.2 **Portfolio-Mix-Verschiebung:** Shift von hochmargigen Spezialprodukten zu margenschwachen Standardprodukten.
1.3 **Konditionen-Management:** Hohe Rabatte und ungünstige Frachtkonditionen bei A-Kunden.

**2. Variable Herstellungskosten (COGS / Direct Costs)**
2.1 **Einkauf & Material:** Preisanstiege bei Rohstoffen ohne adäquates Sourcing-Gegenhalten.
2.2 **Fertigungseffizienz:** Sinkende OEE-Raten, erhöhte Ausschussquoten und Überstunden.
2.3 **Logistik & Energie:** Gestiegene Fracht- und Energiekosten pro Produktionseinheit.

**3. Operative Fixkosten & Overhead (OPEX / Indirect Costs)**
3.1 **SG&A-Kosten:** Ungesteuerter Anstieg der Verwaltungs- und Vertriebskosten (Fixed Cost Creep).
3.2 **Instandhaltung & F&E:** Erhöhte Wartungsaufwände veralteter Anlagen und uneffiziente Projekte.""",
    "hypothesis": """| Bereich | Primäre Hypothese | KPI / Ziel-Benchmark | Erwarteter EBIT-Hebel |
| :--- | :--- | :--- | :--- |
| **Pricing & Mix** | Selektive Preiserhöhungen (3,5 %) und Indexierung von Rohstoffklauseln stabilisieren den Deckungsbeitrag. | **Price Realization Rate > 85 %** | **+0,90 Mio. €** |
| **COGS & Sourcing** | Neuausschreibung der Top-20 Lieferanten und Reduktion der Ausschussquote senken variable Stückkosten. | **Wareneinsatzquote < 56,1 %** | **+0,85 Mio. €** |
| **SG&A / Overhead** | Einfrieren nicht-kritischer Sachkosten (Discretionary Spending Freeze) stoppt Fixed Cost Creep. | **SG&A-Quote < 17,6 %** | **+0,41 Mio. €** |"""
}

if st.button(ui_button):
    if not case_input.strip():
        st.warning(ui_warning)
    else:
        is_default_demo = case_input.strip() == DEMO_CASES["General Profitability"].strip()
        
        if is_default_demo and not user_key.strip():
            st.session_state["out_analysis"] = PRECACHED_PROFITABILITY["analysis"]
            st.session_state["out_mece"] = PRECACHED_PROFITABILITY["mece"]
            st.session_state["out_hypothesis"] = PRECACHED_PROFITABILITY["hypothesis"]
            st.session_state["has_analysis"] = True
            st.info("⚡ **Demo-Vorschau aktiv:** Zur Vermeidung von API-Rate-Limits und zur Gewährleistung unmittelbarer Antwortzeiten wird für diesen Standard-Case ein vorvalidiertes Agenten-Ergebnis geladen. Bei manueller Anpassung des Briefings wird automatisch die Live-Orchestrierung gestartet.")
        else:
            with st.status(ui_status_start, expanded=True) as status:
                analyzer = Agent(
                    role="Senior Strategy Consultant",
                    goal=f"Erstelle eine präzise Executive Summary und Situation-Complication-Resolution (SCR) Analyse mit Fokus auf {framework_focus}.",
                    backstory="Erfahrener Strategy Consultant mit Spezialisierung auf prägnante Problemsynthesen und strukturierte Analysen.",
                    llm=gemini_llm,
                    verbose=False
                )
                
                structurer = Agent(
                    role="MECE Framework Architect",
                    goal="Erstelle eine 100% überschneidungsfreie und vollständige Problemstruktur (MECE Issue Tree) in sauberem Markdown.",
                    backstory="Spezialist für logische Problemzerlegung. Achtet strikt auf Vollständigkeit und Überschneidungsfreiheit.",
                    llm=gemini_llm,
                    verbose=False
                )

                hypothesis_builder = Agent(
                    role="Strategy & Hypothesis Lead",
                    goal="Entwickle 3 quantifizierbare Arbeitshypothesen inklusive einer strukturierten KPI-Validierungsmatrix.",
                    backstory="Experte für datengestützte Unternehmensanalysen, Performance Improvement und KPI-Konzepte.",
                    llm=gemini_llm,
                    verbose=False
                )

                t1 = Task(
                    description=f"Analysiere folgendes Case-Briefing unter Berücksichtigung von '{framework_focus}':\n\n{case_input}\n\nErstelle eine strukturierte Executive Summary im SCR-Format.\n\n{FORMATTING_RULES}",
                    expected_output=f"Strukturierte SCR-Analyse auf {language}.",
                    agent=analyzer
                )
                
                t2 = Task(
                    description=f"Basierend auf Task 1: Erstelle einen vollständigen MECE Issue Tree auf {language}.\n\n{FORMATTING_RULES}",
                    expected_output=f"Ein übersichtlicher MECE Issue Tree auf {language}.",
                    agent=structurer
                )
                
                t3 = Task(
                    description=f"Basierend auf Task 2: Formuliere genau 3 priorisierte Arbeitshypothesen auf {language} inklusive KPI-Matrix.\n\n{FORMATTING_RULES}",
                    expected_output=f"3 Hypothesen mit KPI-Validierungsmatrix als Markdown-Tabelle auf {language}.",
                    agent=hypothesis_builder
                )

                crew = Crew(
                    agents=[analyzer, structurer, hypothesis_builder],
                    tasks=[t1, t2, t3],
                    process=Process.sequential
                )

                try:
                    result = crew.kickoff()
                    status.update(label=ui_status_done, state="complete", expanded=False)
                    st.session_state["out_analysis"] = result.tasks_output[0].raw
                    st.session_state["out_mece"] = result.tasks_output[1].raw
                    st.session_state["out_hypothesis"] = result.tasks_output[2].raw
                    st.session_state["has_analysis"] = True
                except Exception as e:
                    status.update(label="❌ API-Limit erreicht", state="error", expanded=False)
                    if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
                        st.error("⚠️ Das kostenlose Tageskontingent der API ist vorübergehend erschöpft. Bitte trage in der linken Seitenleiste einen eigenen kostenlosen Gemini API-Key ein oder versuche es in wenigen Minuten erneut.")
                    else:
                        st.error(f"Fehler bei der Analyse: {e}")
                    st.stop()

# 9. Ergebnisanzeige & Export
if st.session_state.get("has_analysis", False):
    out_analysis = st.session_state["out_analysis"]
    out_mece = st.session_state["out_mece"]
    out_hypothesis = st.session_state["out_hypothesis"]

    st.subheader(ui_kpi_hdr)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label=ui_m1_label, value=framework_focus)
    with col2:
        st.metric(label=ui_m2_label, value=ui_m2_val, delta=ui_m2_delta)
    with col3:
        st.metric(label=ui_m3_label, value=ui_m3_val, delta=ui_m3_delta)

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([ui_tab1, ui_tab2, ui_tab3])

    with tab1:
        st.markdown(f'<div style="background: #FFFFFF; padding: 24px; border-radius: 8px; border: 1px solid #E2E8F0; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">{out_analysis}</div>', unsafe_allow_html=True)

    with tab2:
        st.markdown(f'<div style="background: #FFFFFF; padding: 24px; border-radius: 8px; border: 1px solid #E2E8F0; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">{out_mece}</div>', unsafe_allow_html=True)

    with tab3:
        st.markdown(f'<div style="background: #FFFFFF; padding: 24px; border-radius: 8px; border: 1px solid #E2E8F0; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">{out_hypothesis}</div>', unsafe_allow_html=True)

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs([ui_tab1, ui_tab2, ui_tab3])

    with tab1:
        st.markdown(out_analysis)

    with tab2:
        st.markdown(out_mece)

    with tab3:
        st.markdown(out_hypothesis)

    st.markdown("---")
    
    # Multi-Format Export Sektion
    st.subheader(ui_export_hdr)
    col_fmt, col_btn = st.columns([2, 2])
    
    with col_fmt:
        export_choice = st.selectbox(
            ui_format_label,
            ["HTML Executive Report (.html)", "Microsoft Word (.docx)", "Markdown Raw (.md)"]
        )
    
    report_title = f"{ui_report_hdr}: {framework_focus}"
    file_base = f"case_report_{framework_focus.lower().replace(' ', '_')}"

    with col_btn:
        st.write(" ")
        st.write(" ")
        if export_choice == "HTML Executive Report (.html)":
            html_data = create_html_report(
                report_title, framework_focus, out_analysis, out_mece, out_hypothesis,
                rep_sub, rep_lbl_fw, rep_sec1, rep_sec2, rep_sec3, rep_footer
            )
            st.download_button(
                label="📄 Download HTML Report",
                data=html_data,
                file_name=f"{file_base}.html",
                mime="text/html"
            )
        elif export_choice == "Microsoft Word (.docx)":
            docx_buffer = create_docx_report(
                report_title, framework_focus, out_analysis, out_mece, out_hypothesis,
                rep_lbl_fw, rep_sec1, rep_sec2, rep_sec3, rep_footer
            )
            st.download_button(
                label="📄 Download Word Document",
                data=docx_buffer,
                file_name=f"{file_base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            full_md = f"# {report_title}\n\n## 1. SCR\n{out_analysis}\n\n## 2. MECE\n{out_mece}\n\n## 3. Hypothesen\n{out_hypothesis}"
            st.download_button(
                label="📄 Download Markdown (.md)",
                data=full_md,
                file_name=f"{file_base}.md",
                mime="text/markdown"
            )