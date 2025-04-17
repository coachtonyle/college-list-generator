import streamlit as st
import pandas as pd
import openai
from io import BytesIO
from docx import Document
from aotomate import (
    setup_openai, generate_overview_content, generate_student_criteria_content,
    generate_college_list, generate_reasons_for_selections, generate_detailed_college_info,
    create_new_document, create_overview_section, create_student_criteria_section,
    create_college_list_table, create_reasons_section, create_detailed_college_info
)

st.set_page_config(page_title="College List Generator", layout="wide")
st.title("📄 College List Generator - Multi Student")

uploaded_file = st.file_uploader("Upload student CSV (exported from the preferences form)", type="csv")
api_key = st.text_input("Enter your OpenAI API Key", type="password")

if uploaded_file and api_key:
    df = pd.read_csv(uploaded_file)
    student_choices = df[["Student's Name - First Name", "Student's Name - Last Name", "Email Address"]].copy()
    student_choices["Student"] = student_choices["Student's Name - First Name"] + " " + student_choices["Student's Name - Last Name"]

    selected_students = st.multiselect(
        "Select student(s) to generate reports for:",
        options=student_choices["Student"].tolist(),
        default=[]
    )

    if st.button("Generate Reports") and selected_students:
        client, is_new_api = setup_openai(api_key)
        generated_docs = []

        for i, student in student_choices.iterrows():
            full_name = f"{student['Student's Name - First Name']} {student['Student's Name - Last Name']}"
            if full_name in selected_students:
                student_info = df.iloc[i].to_dict()
                overview = generate_overview_content(client, student_info, is_new_api)
                criteria = generate_student_criteria_content(client, student_info, is_new_api)
                college_list = generate_college_list(client, student_info, is_new_api)
                reasons = generate_reasons_for_selections(client, student_info, college_list, is_new_api)

                details = {
                    "Reach": generate_detailed_college_info(client, student_info, college_list, "Reach", is_new_api),
                    "Target": generate_detailed_college_info(client, student_info, college_list, "Target", is_new_api),
                    "Safety": generate_detailed_college_info(client, student_info, college_list, "Safety", is_new_api)
                }

                doc, filename = create_new_document(student_info.get("Student's Name - First Name", "Student"), student_info.get("Student's Name - Last Name", "Name"))
                create_overview_section(doc, overview)
                create_student_criteria_section(doc, criteria)
                create_college_list_table(doc, college_list, student_info)
                create_reasons_section(doc, reasons)
                create_detailed_college_info(doc, details)

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                generated_docs.append((full_name, filename, buffer))

        st.success("Reports generated!")
        for full_name, filename, buffer in generated_docs:
            st.download_button(
                label=f"📥 Download {full_name}'s Report",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    elif not selected_students:
        st.info("Select at least one student from the list to generate their report.")
else:
    st.info("Upload a CSV and enter your API key to begin.")