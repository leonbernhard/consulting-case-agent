import os
import streamlit as st
from crewai import Agent, Crew, Process, Task, LLM

# Seiten-Konfiguration
st.set_page_config(page_title="Case Structuring Agent", page_icon="📊", layout="wide")

# 1. API Key setzen
if "GEMINI_API_KEY" in st.secrets:
    os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]

# 2. Modell initialisieren
gemini_llm = LLM(
    model="gemini/gemini-2.0-flash",
    api_key=os.environ["GEMINI_API_KEY"]
)

# 3. Sprachauswahl in der Seitenleiste
with st.sidebar:
    st.header("⚙️ Settings / Einstellungen")
    language = st.selectbox("Language / Sprache", ["Deutsch", "English"])

# 4. Dynamische UI-Texte definieren (bodenständig & professionell)
if language == "English":
    ui_title = "📊 Consulting Case Structuring Agent"
    ui_subtitle = "Structured case analysis, MECE issue trees, and data-driven hypothesis development."
    ui_framework_label = "Select Framework Focus"
    ui_input_label = "Enter Case Briefing here:"
    ui_input_placeholder = "Paste the client's problem description here..."
    ui_button = "🚀 Analyze Case"
    ui_warning = "Please enter a case briefing first."
    ui_spinner = "The AI agents are analyzing the problem..."
    ui_success = "Analysis completed successfully!"
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
    ui_input_placeholder = "Füge hier die Problemstellung des Mandanten ein..."
    ui_button = "🚀 Case Analysieren"
    ui_warning = "Bitte gib zuerst ein Case-Briefing ein."
    ui_spinner = "Die KI-Agenten analysieren das Problem..."
    ui_success = "Analyse erfolgreich abgeschlossen!"
    ui_tab1 = "📌 Executive Summary & SCR"
    ui_tab2 = "🌳 MECE Issue Tree"
    ui_tab3 = "💡 Hypothesen & KPI-Matrix"
    ui_download = "📄 Vollständigen Management-Bericht herunterladen (.md)"
    ui_report_hdr = "📊 Mandanten-Analysebericht"

# 5. UI-Komponenten darstellen
st.title(ui_title)
st.caption(ui_subtitle)
st.divider()

with st.sidebar:
    framework_focus = st.selectbox(
        ui_framework_label,
        ["General Profitability", "Market Entry", "Cost Reduction", "M&A Due Diligence"]
    )

case_input = st.text_area(
    ui_input_label,
    height=180,
    placeholder=ui_input_placeholder
)

# 6. Button & Analyse-Logik
if st.button(ui_button):
    if not case_input.strip():
        st.warning(ui_warning)
    else:
        with st.spinner(ui_spinner):
            
            # Agenten-Definitionen
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

            # Task-Definitionen
            t1 = Task(
                description=(
                    f"Analysiere folgendes Case-Briefing unter Berücksichtigung von '{framework_focus}':\n\n"
                    f"{case_input}\n\n"
                    f"WICHTIG: Erstelle die gesamte Ausgabe strikt auf {language}.\n"
                    "Erstelle eine strukturierte Executive Summary im SCR-Format (Situation, Complication, Key Question, Resolution/Approach)."
                ),
                expected_output=f"Gliederung nach SCR auf {language}.",
                agent=analyzer
            )
            
            t2 = Task(
                description=(
                    f"Basierend auf der Analyse von Task 1: Erstelle einen vollständigen MECE Issue Tree auf {language}.\n"
                    "Anforderungen:\n"
                    "- Nutze eingerückte Markdown-Listen oder Code-Blöcke zur klaren Baum-Darstellung.\n"
                    "- Mindestens 3 Hauptäste (Ebene 1) und jeweils 2-3 Unterpunkte (Ebene 2).\n"
                    "- Halte dich strikt an das MECE-Prinzip."
                ),
                expected_output=f"Ein übersichtlicher MECE Issue Tree auf {language}.",
                agent=structurer
            )
            
            t3 = Task(
                description=(
                    f"Basierend auf dem MECE Issue Tree: Formuliere genau 3 priorisierte Arbeitshypothesen auf {language}.\n"
                    "Anforderungen für JEDE Hypothese:\n"
                    "1. Klare Hypothesenformulierung (Wirkungsmechanismus)\n"
                    "2. Erwarteter EBIT- bzw. Finanz-Hebel\n"
                    "3. Benötigte Datenquellen & spezifische KPIs mit konkreten Benchmark-Schwellenwerten.\n"
                    "Erstelle am Ende eine zusammenfassende Markdown-Tabelle für alle 3 Hypothesen."
                ),
                expected_output=f"3 ausformulierte Hypothesen mit KPI-Validierungsmatrix auf {language}.",
                agent=hypothesis_builder
            )

            # Crew ausführen
            crew = Crew(
                agents=[analyzer, structurer, hypothesis_builder],
                tasks=[t1, t2, t3],
                process=Process.sequential
            )

            result = crew.kickoff()

            # Einzelergebnisse extrahieren
            out_analysis = result.tasks_output[0].raw
            out_mece = result.tasks_output[1].raw
            out_hypothesis = result.tasks_output[2].raw

            # Gesamtbericht aufbereiten
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

            st.success(ui_success)
            
            # Tabs zur Anzeige
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
                file_name=f"case_report_{language.lower()}.md",
                mime="text/markdown"
            )