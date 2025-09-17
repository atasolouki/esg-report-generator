import streamlit as st
from pathlib import Path
import os

# import your existing functions
from main import (
    extract_company_descriptions,
    load_first_sheet,
    build_documents,
    build_or_load_vs,
    build_report,
    save_report_pdf,
    embeddings
)

st.set_page_config(page_title="ESG Report Generator", page_icon="📊", layout="wide")

st.title("📊 ESG Report Generator")
st.markdown("Upload your **DOCX** (company descriptions) and **XLSX** (insights) files to generate ESG reports in PDF format.")

uploaded_docx = st.file_uploader("Upload DOCX file", type=["docx"])
uploaded_xlsx = st.file_uploader("Upload XLSX file", type=["xlsx"])

if uploaded_docx and uploaded_xlsx:
    st.success("✅ Files uploaded successfully!")

    # Save uploaded files temporarily
    temp_docx = Path("temp_docx.docx")
    temp_xlsx = Path("temp_xlsx.xlsx")

    with open(temp_docx, "wb") as f:
        f.write(uploaded_docx.getbuffer())
    with open(temp_xlsx, "wb") as f:
        f.write(uploaded_xlsx.getbuffer())

    # Extract company data
    company_map = extract_company_descriptions(temp_docx)
    insights_df = load_first_sheet(temp_xlsx)

    company_list = list(company_map.keys())

    if company_list:
        selected_company = st.selectbox("Select a company to generate report", company_list)

        if st.button("Generate ESG Report"):
            description = company_map[selected_company]

            with st.spinner(f"Generating ESG report for {selected_company}..."):
                docs = build_documents(selected_company, description, insights_df)
                vs = build_or_load_vs(selected_company, docs, embeddings)
                report_md = build_report(selected_company, vs, industry_hint="Logistics and Supply Chain")
                pdf_path = save_report_pdf(selected_company, report_md)

            st.success(f"✅ Report for {selected_company} generated!")

            with open(pdf_path, "rb") as pdf_file:
                st.download_button(
                    label=f"⬇️ Download {selected_company} ESG Report (PDF)",
                    data=pdf_file,
                    file_name=pdf_path.name,
                    mime="application/pdf"
                )
    else:
        st.error("No companies found in the uploaded DOCX file.")
else:
    st.info("⬆️ Please upload both a DOCX and an XLSX file to continue.")
