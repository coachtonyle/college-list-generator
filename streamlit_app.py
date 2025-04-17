import streamlit as st
import pandas as pd
import openai
from io import BytesIO
from docx import Document
import base64
from aotomate import (
    setup_openai, read_student_data, generate_overview_content, generate_student_criteria_content,
    generate_college_list, generate_reasons_for_selections, generate_detailed_college_info,
    create_new_document, create_overview_section, create_student_criteria_section,
    create_college_list_table, create_reasons_section, create_detailed_college_info
)

st.set_page_config(page_title="College List Generator", layout="centered")
st.title("📄 College List Generator")

st.markdown("Upload a single student CSV (exported from the preferences form)")

uploaded_file = st.file_uploader("Choose CSV file", type="csv")

api_key = st.text_input("Enter your OpenAI API Key", type="password")

if uploaded_file and api_key:
    df = pd.read_csv(uploaded_file)

    if df.shape[0] > 1:
        st.warning("Only the first student will be processed in this version.")

    student_info = df.iloc[0].to_dict()

    with st.spinner("Generating document... this may take 1-2 minutes"):
        global API_KEY
        API_KEY = api_key
        client, is_new_api = setup_openai()
        openai.api_key = api_key

        overview = generate_overview_content(client, student_info, is_new_api)
        criteria = generate_student_criteria_content(client, student_info, is_new_api)
        college_list = generate_college_list(client, student_info, is_new_api)
        reasons = generate_reasons_for_selections(client, student_info, college_list, is_new_api)

        details = {
            "Reach": generate_detailed_college_info(client, student_info, college_list, "Reach", is_new_api),
            "Target": generate_detailed_college_info(client, student_info, college_list, "Target", is_new_api),
            "Safety": generate_detailed_college_info(client, student_info, college_list, "Safety", is_new_api)
        }

        doc, filename = create_new_document(
            student_info.get("Student's Name - First Name", "Student"),
            student_info.get("Student's Name - Last Name", "Name")
        )

        create_overview_section(doc, overview)
        create_student_criteria_section(doc, criteria)
        create_college_list_table(doc, college_list, student_info)
        create_reasons_section(doc, reasons)
        create_detailed_college_info(doc, details)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Document created successfully!")
        st.download_button(
            label="📥 Download College Recommendation Doc",
            data=buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("Upload a file and enter your API key to get started.")
