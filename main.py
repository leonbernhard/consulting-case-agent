import os
from crewai import Agent, Crew, Process, Task, LLM

# 1. Füge hier deinen API-Key aus Google AI Studio zwischen den Anführungszeichen ein
os.environ["GEMINI_API_KEY"] = "DEIN_GEMINI_API_KEY_HIER"

# 2. Kostenloses Modell über Gemini Flash einbinden
gemini_llm = LLM(
    model="gemini/gemini-3.6-flash",
    api_key=os.environ["GEMINI_API_KEY"]
)

# 3. Agenten (Rollen) definieren
analyzer = Agent(
    role="Senior Case Analyst",
    goal="Analysiere das Briefing und bestimme die Kernfrage sowie die Case-Kategorie.",
    backstory="Erfahrener Unternehmensberater, der Unklarheiten rasch auf den Punkt bringt.",
    llm=gemini_llm,
    verbose=True
)

structurer = Agent(
    role="MECE Framework Architect",
    goal="Erstelle eine überschneidungsfreie und vollständige (MECE) Problemstruktur (Issue Tree).",
    backstory="Spezialist für Logikbäume und systematische Zerlegung von Geschäftsproblemen.",
    llm=gemini_llm,
    verbose=True
)

hypothesis_builder = Agent(
    role="Strategy & Hypothesis Lead",
    goal="Formuliere 3 Arbeitshypothesen und definiere die benötigten KPIs.",
    backstory="Experte für datengestützte Validierung und strategische Entscheidungsfindung.",
    llm=gemini_llm,
    verbose=True
)

# 4. Aufgaben (Tasks) definieren
task_analyze = Task(
    description="Analysiere folgendes Briefing:\n{case_description}\nBestimme das primäre Ziel des Kunden, die Case-Kategorie und wichtige Rahmenbedingungen.",
    expected_output="Prägnante Zusammenfassung der Problemstellung und Einordnung.",
    agent=analyzer
)

task_structure = Task(
    description="Erstelle basierend auf der Analyse einen strukturierten MECE Issue Tree mit 2-3 Hauptästen und Unterfragen.",
    expected_output="Strukturierter Issue Tree im Markdown-Format.",
    agent=structurer
)

task_hypotheses = Task(
    description="Entwickle 3 klare Arbeitshypothesen sowie eine Liste der konkret benötigten Daten/KPIs zu deren Prüfung.",
    expected_output="Hypothesenpapier mit Datenanforderungsprofil.",
    agent=hypothesis_builder
)

# 5. Prozess und Team festlegen
consulting_crew = Crew(
    agents=[analyzer, structurer, hypothesis_builder],
    tasks=[task_analyze, task_structure, task_hypotheses],
    process=Process.sequential
)

# Beispiel-Case zum Testen
sample_case = """
Ein führender deutscher Hersteller von Präzisionswerkzeugen verzeichnet seit 18 Monaten 
einen kontinuierlichen Rückgang der operativen Marge von 14% auf 6%. Der Gesamtmarkt wächst leicht, 
aber günstigere Mitbewerber aus Osteuropa gewinnen Marktanteile. Der Vorstand fragt sich, 
ob die Preise gesenkt werden müssen oder ob das Kostenproblem in der eigenen Produktion liegt.
"""

if __name__ == "__main__":
    print("Starte die Agenten-Analyse...\n")
    result = consulting_crew.kickoff(inputs={"case_description": sample_case})
    print("\n================ ERGEBNIS ================\n")
    print(result)