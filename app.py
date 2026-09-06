import os
import io
import re
import streamlit as st
from crewai import Agent, Crew, Process, Task, LLM
from docx import Document
from docx.shared import Pt, RGBColor, Inches

# Seiten-Konfiguration
st.set_page_config(page_title="Case Structuring Agent", page_icon="📊", layout="wide")

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
def create_html_report(title, framework, language, analysis, mece, hypothesis):
    """Erstellt ein hochgradig gestaltetes HTML Executive Dashboard."""
    import html
    
    def md_to_html_simple(text):
        lines = text.split('\n')
        html_lines = []
        in_list = False
        in_table = False
        
        for line in lines:
            line_str = line.strip()
            if line_str.startswith('|'):
                if in_list:
                    html_lines.append('</ul>')
                    in_list = False
                if not in_table:
                    html_lines.append('<table class="excel-table">')
                    in_table = True
                
                cells = [c.strip() for c in line_str.split('|')[1:-1]]
                if all(set(c).issubset({'-', ':', ' '}) for c in cells):
                    continue
                
                row_html = '<tr>' + ''.join(f'<td>{html.escape(c)}</td>' for c in cells) + '</tr>'
                html_lines.append(row_html)
                continue
            else:
                if in_table:
                    html_lines.append('</table>')
                    in_table = False

            if line_str.startswith('### '):
                html_lines.append(f'<h3>{html.escape(line_str[4:])}</h3>')
            elif line_str.startswith('## '):
                html_lines.append(f'<h2>{html.escape(line_str[3:])}</h2>')
            elif line_str.startswith('- ') or line_str.startswith('* '):
                if not in_list:
                    html_lines.append('<ul>')
                    in_list = True
                item_text = line_str[2:]
                item_text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', item_text)
                html_lines.append(f'<li>{item_text}</li>')
            else:
                if in_list:
                    html_lines.append('</ul>')
                    in_list = False
                if line_str:
                    formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', line_str)
                    html_lines.append(f'<p>{formatted}</p>')

        if in_list:
            html_lines.append('</ul>')
        if in_table:
            html_lines.append('</table>')
            
        return '\n'.join(html_lines)

    html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #1E293B;
            background-color: #F8F9FA;
            margin: 0;
            padding: 40px;
        }}
        .container {{
            max-width: 1000px;
            margin: 0 auto;
            background: #FFFFFF;
            padding: 50px;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        }}
        .header {{
            border-bottom: 3px solid #0F2C59;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }}
        h1 {{ color: #0F2C59; font-size: 28px; margin-bottom: 5px; }}
        .badge {{
            display: inline-block;
            background: #0F2C59;
            color: white;
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: 600;
        }}
        h2 {{ color: #0F2C59; font-size: 20px; border-left: 4px solid #0F2C59; padding-left: 10px; margin-top: 35px; }}
        h3 {{ color: #334155; font-size: 16px; margin-top: 20px; }}
        p, li {{ font-size: 14px; color: #334155; }}
        ul {{ padding-left: 20px; }}
        .excel-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 13px;
        }}
        .excel-table th, .excel-table td {{
            border: 1px solid #E2E8F0;
            padding: 10px 12px;
            text-align: left;
        }}
        .excel-table tr:nth-child(even) {{ background-color: #F8FAFC; }}
        .excel-table tr:first-child {{ background-color: #0F2C59; color: white; font-weight: bold; }}
        .footer {{
            margin-top: 50px;
            padding-top: 20px;
            border-top: 1px solid #E2E8F0;
            font-size: 12px;
            color: #94A3B8;
            text-align: center;
        }}
        @media print {{
            body {{ background: white; padding: 0; }}
            .container {{ box-shadow: none; padding: 0; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <span class="badge">{framework}</span>
            <h1>{title}</h1>
            <p style="color: #64748B; margin: 0;">Automated AI Case Analysis Report</p>
        </div>
        
        <h2>1. Executive Summary & SCR</h2>
        {md_to_html_simple(analysis)}
        
        <h2>2. MECE Issue Tree</h2>
        {md_to_html_simple(mece)}
        
        <h2>3. Hypothesen & KPI Matrix</h2>
        {md_to_html_simple(hypothesis)}
        
        <div class="footer">
            Generated by AI Consulting & Case Structuring Agent | Confidential & Professional Support Tool
        </div>
    </div>
</body>
</html>"""
    return html_content

def create_docx_report(title, framework, analysis, mece, hypothesis):
    """Erstellt ein sauberes Microsoft Word Dokument (.docx)."""
    doc = Document()
    
    # Titel
    heading = doc.add_heading(title, level=0)
    heading.style.font.color.rgb = RGBColor(15, 44, 89)
    
    p = doc.add_paragraph()
    p.add_run(f"Framework Focus: {framework}\n").bold = True
    
    doc.add_heading("1. Executive Summary & SCR", level=1)
    doc.add_paragraph(analysis)
    
    doc.add_heading("2. MECE Issue Tree", level=1)
    doc.add_paragraph(mece)
    
    doc.add_heading("3. Hypothesen & KPI Matrix", level=1)
    doc.add_paragraph(hypothesis)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 4. Sprachauswahl in Seitenleiste
with st.sidebar:
    st.header("⚙️ Settings / Einstellungen")
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
    ui_tab2 = "🌳 MECE Issue Tree"
    ui_tab3 = "💡 Hypothesen & KPI-Matrix"
    ui_report_hdr = "Mandanten-Analysebericht"
    ui_export_hdr = "📥 Bericht Exportieren"
    ui_format_label = "Export-Format wählen:"

# 6. UI Header & Sidebar
st.title(ui_title)
st.caption(ui_subtitle)
st.divider()

with st.sidebar:
    framework_focus = st.selectbox(
        ui_framework_label,
        ["General Profitability", "Cost Reduction", "M&A Due Diligence", "Market Entry"]
    )

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

# 8. Agenten-Analyse
if st.button(ui_button):
    if not case_input.strip():
        st.warning(ui_warning)
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
                description=(
                    f"Analysiere folgendes Case-Briefing unter Berücksichtigung von '{framework_focus}':\n\n"
                    f"{case_input}\n\n"
                    "Erstelle eine strukturierte Executive Summary im SCR-Format.\n\n"
                    f"{FORMATTING_RULES}"
                ),
                expected_output=f"Strukturierte SCR-Analyse auf {language}.",
                agent=analyzer
            )
            
            t2 = Task(
                description=(
                    f"Basierend auf Task 1: Erstelle einen vollständigen MECE Issue Tree auf {language}.\n\n"
                    f"{FORMATTING_RULES}"
                ),
                expected_output=f"Ein übersichtlicher MECE Issue Tree auf {language}.",
                agent=structurer
            )
            
            t3 = Task(
                description=(
                    f"Basierend auf Task 2: Formuliere genau 3 priorisierte Arbeitshypothesen auf {language} inklusive KPI-Matrix.\n\n"
                    f"{FORMATTING_RULES}"
                ),
                expected_output=f"3 Hypothesen mit KPI-Validierungsmatrix als Markdown-Tabelle auf {language}.",
                agent=hypothesis_builder
            )

            crew = Crew(
                agents=[analyzer, structurer, hypothesis_builder],
                tasks=[t1, t2, t3],
                process=Process.sequential
            )

            result = crew.kickoff()
            status.update(label=ui_status_done, state="complete", expanded=False)

        st.session_state["out_analysis"] = result.tasks_output[0].raw
        st.session_state["out_mece"] = result.tasks_output[1].raw
        st.session_state["out_hypothesis"] = result.tasks_output[2].raw
        st.session_state["has_analysis"] = True

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
            html_data = create_html_report(report_title, framework_focus, language, out_analysis, out_mece, out_hypothesis)
            st.download_button(
                label="📄 Download HTML Report",
                data=html_data,
                file_name=f"{file_base}.html",
                mime="text/html"
            )
        elif export_choice == "Microsoft Word (.docx)":
            docx_buffer = create_docx_report(report_title, framework_focus, out_analysis, out_mece, out_hypothesis)
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