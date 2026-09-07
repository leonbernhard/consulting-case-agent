# 📊 Consulting Case Structuring Agent

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)
![Python](https://img.shields.io/badge/Python-3.10%2B-0F2C59?style=flat&logo=python&logoColor=white)
![CrewAI](https://img.shields.io/badge/Orchestration-CrewAI-1E40AF?style=flat)
![LLM](https://img.shields.io/badge/LLM-Gemini_3.6_Flash-0F2C59?style=flat)
![License](https://img.shields.io/badge/License-MIT-green?style=flat)
[![Live Demo](https://img.shields.io/badge/🚀_Live_Demo-Open_App-0F2C59?style=for-the-badge)](https://leon-bernhard-case-agent.streamlit.app)

An enterprise-grade, multi-agent AI system designed to automate strategic problem breakdown, MECE issue trees, and hypothesis-driven KPI matrix development for management consulting, private equity, and corporate finance cases.

---

## 🎯 Executive Summary & Core Value Proposition

In top-tier management consulting and M&A advisory, structuring complex client problems rapidly and logically is critical. The **Consulting Case Structuring Agent** leverages a sequential multi-agent AI crew to transform raw corporate briefings into actionable advisory deliverables within seconds.

* **Methodological Rigor:** Enforces strict consulting standards including **Situation-Complication-Resolution (SCR)** frameworks and **Mutually Exclusive, Collectively Exhaustive (MECE)** issue trees.
* **Dual-Execution Pipeline:** Features a live multi-agent execution engine via **CrewAI & Gemini 3.6-flash** alongside a zero-quota pre-cached preview mode for instant, quota-safe demonstrations.
* **C-Level Deliverable Export:** Generates client-ready deliverables formatted for executive presentation in **HTML (A4 print-optimized)**, **Microsoft Word (.docx)** with native grid tables, and **Raw Markdown (.md)**.

---

## 🏗️ System Architecture & Workflow

The platform operates on a sequential three-agent workflow coordinated through CrewAI. Context and task outputs are dynamically passed downstream to build a cohesive strategic synthesis.

```mermaid
graph TD
    User([User Briefing Input]) --> UI[Streamlit Executive Interface]
    UI --> Router{Execution Mode}
    
    Router -->|Demo Mode| Cache[Pre-Cached Zero-Quota Engine]
    Router -->|Live Analysis| Crew[CrewAI Sequential Pipeline]
    
    subgraph Multi-Agent AI Crew
        Crew --> Agent1[Agent 1: Senior Strategy Consultant]
        Agent1 -->|Task 1: SCR Executive Summary| Agent2[Agent 2: MECE Framework Architect]
        Agent2 -->|Task 2: MECE Issue Tree| Agent3[Agent 3: Strategy & Hypothesis Lead]
        Agent3 -->|Task 3: Quantified KPI Matrix| Synthesizer[Structured Synthesis Output]
    end
    
    Cache --> Output[Executive Dashboard Display]
    Synthesizer --> Output
    
    Output --> Export[Multi-Format Export Engine]
    Export --> PDF[Print-Ready HTML / PDF]
    Export --> DOCX[Microsoft Word .docx]
    Export --> MD[Raw Markdown .md]
### 🛠️ Technical Challenges & Engineering Highlights

* **Quota Resilience & Zero-Downtime:** Engineered a hybrid execution engine that dynamically switches between live LLM orchestration and pre-cached benchmark outputs, ensuring 100% uptime and instant responses during recruiter demonstrations.
* **Data Sanitization & Formatting:** Implemented custom regex-based text normalization to strip raw Markdown artifacts before generating native `openpyxl` Excel grids and structured Word tables.
* **C-Level PDF Precision:** Configured native `@page` CSS print boundary rules to enforce A4 layout constraints and prevent mid-table page splits in generated HTML/PDF deliverables.