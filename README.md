# 📊 AI Consulting Case Structuring Agent

An interactive, multi-agent AI system built with **CrewAI**, **Streamlit**, and **Google Gemini** to automate early-stage case analysis, construct MECE issue trees, and generate data-driven hypothesis matrices for strategic consulting engagements.

---

## 🌟 Key Features

* **Multi-Agent Collaboration:** Sequential workflow utilizing specialized AI roles (Strategy Consultant, MECE Architect, Hypothesis Lead).
* **MECE Framework Construction:** Automated generation of mutually exclusive, collectively exhaustive issue trees.
* **KPI & Validation Matrix:** Formulates quantifiable hypotheses with target benchmarks and data sources (e.g., ERP, shopfloor data).
* **Bilingual UI:** Full support for English and German case briefs and outputs.
* **Executive Export:** Instant export of full management reports in Markdown format.

---

## 🛠️ Architecture & Tech Stack

* **Frontend:** Streamlit
* **Orchestration:** CrewAI Framework
* **LLM:** Google Gemini (`gemini-3.6-flash`) via `google-genai`
* **Language:** Python 3.11+

---

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have Python installed and obtain a free API key from Google AI Studio.

### 2. Installation
Clone this repository and install required packages:
```bash
git clone [https://github.com/YOUR_USERNAME/consulting-case-agent.git](https://github.com/YOUR_USERNAME/consulting-case-agent.git)
cd consulting-case-agent
pip install -r requirements.txt