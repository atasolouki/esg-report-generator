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

### Environment Setup

Before running the project, create a file named `.env` in the root directory of the project and fill in the required values:
```
AZURE_OPENAI_ENDPOINT=<your-azure-endpoint>
AZURE_OPENAI_API_KEY=<your-azure-api-key>
AZURE_OPENAI_DEPLOYMENT_LLM=o4-mini
AZURE_OPENAI_DEPLOYMENT_EMBEDDING=text-embedding-3-large
AZURE_OPENAI_API_VERSION=2025-01-01-preview
TAVILY_API_KEY=<your-tavily-api-key>
```
## Running Locally
### Streamlit web app
```
streamlit run app.py
```
Upload DOCX and XLSX files → download the generated ESG report (PDF + Markdown).

### Jupyter Notebook
A step-by-step notebook is provided to walk through the pipeline with explanations.

## Running with Docker
Build image
```
docker build -t esg-report-app .
```
Run container
```
docker run -p 8501:8501 --env-file .env \
  -v $(pwd)/data:/app/data \
  -v $(pwd)/reports:/app/reports \
  esg-report-app
```
Then open: http://localhost:8501

## Output
- report_<company>.md → Editable Markdown report
- report_<company>.pdf → Styled PDF report
