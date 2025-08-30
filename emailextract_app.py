import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
from docx import Document

st.title("Email Extractor from Documents")

uploaded_file = st.file_uploader("Upload a file (PDF or DOCX)", type=["pdf", "docx"])

emails = []
if uploaded_file:
    file_type = uploaded_file.type
    file_bytes = uploaded_file.read()
    full_text = ""
    if file_type == "application/pdf":
        doc = fitz.open(stream=BytesIO(file_bytes), filetype="pdf")
        for page in doc:
            full_text += page.get_text("text")
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        document = Document(BytesIO(file_bytes))
        for para in document.paragraphs:
            full_text += para.text + "\n"
    else:
        st.error("Unsupported file type.")
    if full_text:
        email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
        emails = sorted(set(re.findall(email_pattern, full_text)))
        if emails:
            st.success(f"Extracted {len(emails)} emails:")
            st.dataframe(pd.DataFrame(emails, columns=["Email"]))
            csv = pd.DataFrame(emails, columns=["Email"]).to_csv(index=False).encode()
            st.download_button("Download CSV", csv, "extracted_emails.csv", "text/csv")
        else:
            st.warning("No emails found in the document.")
