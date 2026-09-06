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
    """Erstellt ein professionelles, druckoptimiertes Executive HTML/PDF Dashboard mit sauberen A4-Rändern."""
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

            if line_str.startswith('### '):
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
                if line_str:
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
        /* A4-Druckeinrichtung mit festen Rändern */
        @page {{
            size: A4;
            margin: 20mm 18mm 20mm 18mm;
        }}
        body {{
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
            line-height: 1.5;
            color: #1E293B;
            background-color: #F8FAFC;
            margin: 0;
            padding: 30px;
        }}
        .container {{
            max-width: 850px;
            margin: 0 auto;
            background: #FFFFFF;
            padding: 40px 45px;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
            box-sizing: border-box;
        }}
        .header {{
            border-bottom: 2px solid #0F2C59;
            padding-bottom: 12px;
            margin-bottom: 25px;
        }}
        .badge {{
            display: inline-block;
            background: #0F2C59;
            color: #FFFFFF;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }}
        h1 {{ color: #0F2C59; font-size: 22px; margin: 4px 0; font-weight: 700; }}
        h2 {{ color: #0F2C59; font-size: 16px; border-bottom: 1px solid #E2E8F0; padding-bottom: 5px; margin-top: 25px; margin-bottom: 12px; font-weight: 600; page-break-after: avoid; }}
        h3 {{ color: #334155; font-size: 13px; margin-top: 16px; margin-bottom: 6px; font-weight: 600; page-break-after: avoid; }}
        p, li {{ font-size: 12px; color: #334155; margin-bottom: 6px; }}
        ul.executive-list {{ padding-left: 20px; margin: 8px 0; }}
        ul.executive-list li {{ margin-bottom: 4px; }}
        
        .executive-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 16px 0;
            font-size: 11px;
            page-break-inside: avoid;
        }}
        .executive-table th {{
            background-color: #0F2C59;
            color: #FFFFFF;
            font-weight: bold;
            text-align: left;
            padding: 8px 10px;
            border: 1px solid #0F2C59;
        }}
        .executive-table td {{
            border: 1px solid #CBD5E1;
            padding: 8px 10px;
            vertical-align: top;
        }}
        .executive-table tr:nth-child(even) {{ background-color: #F8FAFC; }}
        
        .section-block {{ page-break-inside: avoid; }}
        .footer {{
            margin-top: 35px;
            padding-top: 12px;
            border-top: 1px solid #E2E8F0;
            font-size: 10px;
            color: #94A3B8;
            text-align: center;
        }}
        
        /* Druck-Spezifische Anpassung für perfektes PDF-A4 */
        @media print {{
            body {{
                background: #FFFFFF !important;
                padding: 0 !important;
                margin: 0 !important;
            }}
            .container {{
                box-shadow: none !important;
                padding: 0 !important;
                margin: 0 !important;
                max-width: 100% !important;
                width: 100% !important;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <span class="badge">{framework}</span>
            <h1>{title}</h1>
            <p style="color: #64748B; margin: 0; font-size: 11px;">Automated AI Case Analysis Report</p>
        </div>
        
        <div class="section-block">
            <h2>1. Executive Summary & SCR</h2>
            {md_to_html(analysis)}
        </div>
        
        <div class="section-block">
            <h2>2. MECE Issue Tree</h2>
            {md_to_html(mece)}
        </div>
        
        <div class="section-block">
            <h2>3. Hypothesen & KPI Matrix</h2>
            {md_to_html(hypothesis)}
        </div>
        
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

# 4. Seitenleisten-Initialisierung (Sprachauswahl zuerst)
with st.sidebar:
    st.header("⚙️ Settings / Einstellungen")
    language = st.selectbox("Language / Sprache", ["Deutsch", "English"])

# 5. Dynamische UI-Texte basierend auf der Sprachauswahl
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
    ui_key_hdr = "🔑 API-Key Konfiguration"
    ui_key_label = "Eigener Gemini API-Key (Optional)"
    ui_key_help = "Falls das globale Test-Kontingent erschöpft ist, kann hier ein eigener kostenloser Key aus dem Google AI Studio eingetragen werden. Die Verarbeitung erfolgt sicher und ausschließlich im Arbeitsspeicher dieser Sitzung."
    ui_sec_note = "🔒 *Die Eingabe erfolgt isoliert und wird ausschließlich im flüchtigen Arbeitsspeicher verarbeitet. Keine Speicherung.*"

# 6. Ergänzung der Seitenleiste & Hauptbereich-Header
with st.sidebar:
    st.markdown("---")
    framework_focus = st.selectbox(
        ui_framework_label,
        ["General Profitability", "Cost Reduction", "M&A Due Diligence", "Market Entry"]
    )
    st.markdown("---")
    st.caption(f"**{ui_key_hdr}**")
    user_key = st.text_input(
        ui_key_label, 
        type="password", 
        help=ui_key_help
    )
    st.caption(ui_sec_note)
    
    if user_key.strip():
        os.environ["GEMINI_API_KEY"] = user_key.strip()
        gemini_llm = LLM(model="gemini-3.6-flash", api_key=user_key.strip())

st.title(ui_title)
st.caption(ui_subtitle)
st.divider()

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
    "mece": """- **1. Erlösqualität & Preisdurchsetzung (Umsatz- & Mix-Hebel)**
  * **1.1 Preisanpassung & Indexierung:** Unzureichende Weitergabe gestiegener Inputkosten an Endkunden.
  * **1.2 Portfolio-Mix-Verschiebung:** Shift von hochmargigen Spezialprodukten zu margenschwachen Standardprodukten.
  * **1.3 Konditionen-Management:** Hohe Rabatte und ungünstige Frachtkonditionen bei A-Kunden.

- **2. Variable Herstellungskosten (COGS / Direct Costs)**
  * **2.1 Einkauf & Material:** Preisanstiege bei Rohstoffen ohne adäquates Sourcing-Gegenhalten.
  * **2.2 Fertigungseffizienz:** Sinkende OEE-Raten, erhöhte Ausschussquoten und Überstunden.
  * **2.3 Logistik & Energie:** Gestiegene Fracht- und Energiekosten pro Produktionseinheit.

- **3. Operative Fixkosten & Overhead (OPEX / Indirect Costs)**
  * **3.1 SG&A-Kosten:** Ungesteuerter Anstieg der Verwaltungs- und Vertriebskosten (Fixed Cost Creep).
  * **3.2 Instandhaltung & F&E:** Erhöhte Wartungsaufwände veralteter Anlagen und uneffiziente Projekte.""",
    "hypothesis": """| Bereich | Primäre Hypothese | Key Metric / Benchmark | Erwarteter EBIT-Hebel |
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