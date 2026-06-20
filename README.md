# 📄 Resume.io AI Builder & Co-Pilot

A professional-grade resume builder and refiner web application modeled after the standards of **Resume.io**. It enables users to fill in their professional details, preview their resume in premium designs (Dublin, Toronto, Stockholm), and use a **Deep AI Agent** (powered by LangGraph, LangChain, and `deepagents`) to research and optimize their content.

The app supports downloading the resume in **styled multi-sheet Excel format**, standard **HTML**, or printing/saving directly as a **pixel-perfect PDF** via browser printing.

---

## 🛠️ Technology Stack

1. **Backend / Agentic Layer:**
   - **Python** (version >= 3.13)
   - **LangGraph** (Stateful multi-agent orchestration)
   - **LangChain** (LLM abstractions and tools)
   - **Deep Agents (`deepagents`)** (Harness framework for planning, subagents, and structured data execution)
   - **Tavily API** (Web search tools for researching role requirements and keywords)
   - **Groq API** (Llama-3.3-70b-versatile as primary reasoning LLM)
   - **OpenAI API** (GPT-4o-mini as automated fallback)
2. **Frontend:**
   - **Streamlit** (Python-based interactive web application framework)
3. **Data & Exports:**
   - **openpyxl** (Custom styling and multi-sheet generation for Excel exports)
   - **Jinja2 / CSS** (Resume.io design templates: Dublin, Toronto, Stockholm)

---

## ✨ Features

- **Interactive Form Builder:** Fill and edit contact details, professional summaries, work history, education, skills, side projects, certifications, and languages.
- **Sample Data Loader:** Load a pre-built vague/weak software developer resume with a single click to instantly test the agent's optimization capabilities.
- **AI Co-Pilot (Deep Agent):**
  - **Self-Planning:** The agent plans steps to refine the resume.
  - **Live Web Research:** Uses Tavily to query standard responsibilities, skills, and industry terms for the target role.
  - **Quantified Achievements (XYZ Formula):** Rewrites work experience bullets to focus on metrics and impact (e.g. *Accomplished [X], as measured by [Y], by doing [Z]*).
  - **Active Voice:** Replaces passive phrases with strong action verbs.
  - **Structured Pydantic Returns:** Enforces perfect structural validity conforming to the Pydantic schema.
- **Premium Resume.io Templates:**
  - **Dublin:** Standard two-column layout with sidebar and main content areas.
  - **Toronto:** Single-column layout with a sleek colored top accent bar.
  - **Stockholm:** Minimalist black & white editorial design with Playfair Display typography.
- **Multiple Exporters:**
  - **Excel Download:** Structured sheet-per-section (.xlsx) formatted with headers, autospaced columns, and text wrapping.
  - **HTML Download:** Standalone responsive file.
  - **Print / PDF Save:** Open a print window where you can print to physical paper or print-to-PDF with correct margin breaks.

---

## 🚀 Getting Started

### 📋 Prerequisites

Ensure you have [Python 3.13+](https://www.python.org/downloads/) and [uv](https://github.com/astral-sh/uv) installed.

### ⚙️ Installation

1. **Clone the project workspace:**
   ```bash
   cd C:\Roshan\ResumeAI
   ```

2. **Sync the dependencies:**
   `uv` will automatically read `pyproject.toml` and setup the virtual environment:
   ```bash
   uv sync
   ```

### 🏃 Running the Web App

Start the Streamlit development server:
```bash
uv run streamlit run app.py
```
Or use the virtual environment's executable directly:
```bash
.venv\Scripts\streamlit.exe run app.py
```

Once running, the application will open in your default web browser (typically at `http://localhost:8501`).

---

## 📂 Project Structure

- [app.py](file:///C:/Roshan/ResumeAI/app.py) - Streamlit web interface and session state manager.
- [schemas.py](file:///C:/Roshan/ResumeAI/schemas.py) - Pydantic schemas validating structured resume data.
- [models.py](file:///C:/Roshan/ResumeAI/models.py) - LLM constructor with API key authentication and fallback configurations.
- [agent.py](file:///C:/Roshan/ResumeAI/agent.py) - Deep agent setup, system instructions, and Tavily tool integration.
- [exporter.py](file:///C:/Roshan/ResumeAI/exporter.py) - Styled Excel generator and HTML resume templates.
