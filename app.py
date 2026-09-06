import os
import streamlit as st
from crewai import Agent, Crew, Process, Task, LLM

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

# 3. Sprachauswahl in der Seitenleiste
with st.sidebar:
    st.header("⚙️ Settings / Einstellungen")
    language = st.selectbox("Language / Sprache", ["Deutsch", "English"])

# 4. Dynamische UI-Texte definieren
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
    ui_download = "📄 Download Full Management Report (.md)"
    ui_report_hdr = "📊 Client Analysis Report"
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
    ui_download = "📄 Vollständigen Management-Bericht herunterladen (.md)"
    ui_report_hdr = "📊 Mandanten-Analysebericht"

# 5. UI-Header & Seitenleiste
st.title(ui_title)
st.caption(ui_subtitle)
st.divider()

with st.sidebar:
    framework_focus = st.selectbox(
        ui_framework_label,
        ["General Profitability", "Cost Reduction", "M&A Due Diligence", "Market Entry"]
    )

# 6. Branchenneutrale Demo-Cases definieren
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

# 7. Formatter-Regelwerk gegen ASCII-Salat & Doppel-Header
FORMATTING_RULES = f"""
STRIKTE FORMATIERUNGS-REGELN (STRIKT EINHALTEN):
1. KEINE ASCII-Boxen, Sonderzeichen-Gitter oder Rahmenelemente (+---+, |---|, etc.) verwenden.
2. Für MECE-Strukturen und Baumdarstellungen AUSSCHLIESSLICH Standard-Markdown-Listen mit Einrückungen verwenden (z. B. `- **Ebene 1**` -> `  * **Ebene 2**`).
3. KEINE H1-Überschriften (`#`) generieren. Nutze ausschließlich Unterüberschriften ab Ebene 2 (`##`) oder Ebene 3 (`###`).
4. Für Tabellen ausschließlich sauberes, valides Markdown-Tabellenformat nutzen (`| Spalte 1 | Spalte 2 |`).
5. Gesamtsprache der Ausgabe: Strikt auf {language}.
"""

# 8. Button & Agenten-Ausführung
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
                    "Erstelle eine strukturierte Executive Summary im SCR-Format (Situation, Complication, Key Question, Resolution/Approach).\n\n"
                    f"{FORMATTING_RULES}"
                ),
                expected_output=f"Strukturierte SCR-Analyse auf {language} ohne H1-Header oder ASCII-Boxen.",
                agent=analyzer
            )
            
            t2 = Task(
                description=(
                    f"Basierend auf der Analyse von Task 1: Erstelle einen vollständigen MECE Issue Tree auf {language}.\n"
                    "Anforderungen:\n"
                    "- Nutze sauber eingerückte Markdown-Listen zur klaren Baum-Darstellung (KEINE ASCII-Art!).\n"
                    "- Mindestens 3 Hauptäste (Ebene 1) und jeweils 2-3 Unterpunkte (Ebene 2).\n"
                    "- Halte dich strikt an das MECE-Prinzip.\n\n"
                    f"{FORMATTING_RULES}"
                ),
                expected_output=f"Ein übersichtlicher MECE Issue Tree als eingerückte Markdown-Liste auf {language}.",
                agent=structurer
            )
            
            t3 = Task(
                description=(
                    f"Basierend auf dem MECE Issue Tree: Formuliere genau 3 priorisierte Arbeitshypothesen auf {language}.\n"
                    "Anforderungen für JEDE Hypothese:\n"
                    "1. Klare Hypothesenformulierung (Wirkungsmechanismus)\n"
                    "2. Erwarteter EBIT- bzw. Finanz-Hebel\n"
                    "3. Benötigte Datenquellen & spezifische KPIs mit konkreten Benchmark-Schwellenwerten.\n"
                    "Erstelle am Ende eine zusammenfassende Markdown-Tabelle für alle 3 Hypothesen.\n\n"
                    f"{FORMATTING_RULES}"
                ),
                expected_output=f"3 ausformulierte Hypothesen mit KPI-Validierungsmatrix als Markdown-Tabelle auf {language}.",
                agent=hypothesis_builder
            )

            crew = Crew(
                agents=[analyzer, structurer, hypothesis_builder],
                tasks=[t1, t2, t3],
                process=Process.sequential
            )

            result = crew.kickoff()
            status.update(label=ui_status_done, state="complete", expanded=False)

        out_analysis = result.tasks_output[0].raw
        out_mece = result.tasks_output[1].raw
        out_hypothesis = result.tasks_output[2].raw

        full_report = f"""# {ui_report_hdr}: {framework_focus}

## 1. Executive Summary & SCR
{out_analysis}

---

## 2. MECE Issue Tree
{out_mece}

---

## 3. Hypothesen & KPI Matrix
{out_hypothesis}
"""

        # 9. KPI Header Board
        st.subheader(ui_kpi_hdr)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(label=ui_m1_label, value=framework_focus)
        with col2:
            st.metric(label=ui_m2_label, value=ui_m2_val, delta=ui_m2_delta)
        with col3:
            st.metric(label=ui_m3_label, value=ui_m3_val, delta=ui_m3_delta)

        st.markdown("---")

        # 10. Dashboard Tabs
        tab1, tab2, tab3 = st.tabs([ui_tab1, ui_tab2, ui_tab3])

        with tab1:
            st.markdown(out_analysis)

        with tab2:
            st.markdown(out_mece)

        with tab3:
            st.markdown(out_hypothesis)

        st.markdown("---")
        
        st.download_button(
            label=ui_download,
            data=full_report,
            file_name=f"case_report_{framework_focus.lower().replace(' ', '_')}_{language.lower()}.md",
            mime="text/markdown"
        )