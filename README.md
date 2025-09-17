# ESG Report Generator

This project generates **company-specific ESG (Environmental & Social)** reports by combining three data sources:  
- DOCX company descriptions  
- XLSX insights  
- Online ESG search results (via Tavily API)  

The pipeline uses **Azure OpenAI** (LLM + embeddings) and **LangChain** to retrieve, plan, and generate evidence-based reports aligned with **ESRS standards**. Reports are exported in both **Markdown** and **PDF** formats. A **Streamlit web app** is included for interactive use.

---

## Features

- Extract company descriptions from DOCX  
- Load ESG insights from Excel  
- Retrieve additional online ESG data via Tavily API  
- Store documents in FAISS vector stores for retrieval  
- Use Azure OpenAI for:  
  - Section planning  
  - Evidence retrieval  
  - Report writing (impacts, gaps, recommendations)  
- Export results as Markdown (`.md`) and styled PDF (`.pdf`)  
- Streamlit web app with file upload + report download  
- Dockerized for reproducible deployment  

---

## Project Structure

├── app.py                 # Streamlit web app
├── version1.py            # Core pipeline functions
├── requirements.txt       # Python dependencies
├── data/                  # Input DOCX and XLSX files
├── reports/               # Generated reports (md + pdf)
├── vectorstores/          # Stored FAISS indexes
└── README.md              # Project documentation



---

## Installation

### 1. Clone repository
```bash
git clone https://github.com/atasolouki/esg-report-generator.git
cd esg-report-generator
```

### Create virtual environment
```bash
python -m venv .venv
source .venv/bin/activate
```

### Install dependencies
```
pip install -r requirements.txt
```
