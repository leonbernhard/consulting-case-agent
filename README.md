# 📊 AI Consulting & Case Structuring Agent

An open-source, multi-agent AI framework designed for strategy consultants, corporate finance analysts, and business teams to automate initial case structuring, MECE issue tree generation, and hypothesis validation.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.38+-red.svg)
![CrewAI](https://img.shields.io/badge/CrewAI-Multi--Agent-green.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## 🎯 Key Features & Framework Focus
Designed to streamline strategic problem decomposition across multiple business domains:
- **SCR Framework Analysis:** Synthesizes complex client briefings into structured Situation, Complication, Key Question, and Resolution.
- **MECE Issue Trees:** Decomposes operational and financial value drivers into 100% mutually exclusive and collectively exhaustive logic trees.
- **Hypothesis & KPI Matrix:** Formulates 3 prioritized, quantitative hypotheses backed by concrete financial benchmarks and metrics.
- **Multi-Framework Flexibility:** Pre-built logic for Profitability, Market Entry, Cost Reduction, and M&A Due Diligence.

## 🏗️ Architecture & Tech Stack (100% Free & Open Source)
- **Frontend:** Streamlit
- **Multi-Agent Orchestration:** CrewAI
- **LLM Engine:** Google Gemini API (gemini-3.6-flash)
- **Privacy First:** No client data persistent storage.

## 🚀 Local Installation & Quickstart

1. Clone the repository:
   git clone https://github.com/leonbernhard/consulting-case-agent.git
   cd consulting-case-agent

2. Install dependencies:
   pip install -r requirements.txt

3. Configure your API Key:
   Create .streamlit/secrets.toml in the project root:
   GEMINI_API_KEY = "your_free_gemini_api_key_here"

4. Launch the application:
   streamlit run app.py

---
*Disclaimer: This tool is an automated decision-support assistant designed for initial case structuring. All AI-generated analyses should be validated by a qualified domain expert.*