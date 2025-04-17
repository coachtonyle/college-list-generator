
import openai
from openai import OpenAI

def setup_openai(api_key):
    openai.api_key = api_key
    client = OpenAI(api_key=api_key)
    is_new_api = hasattr(client, 'chat')
    return client, is_new_api

def generate_overview_content(client, student_info, is_new_api):
    # Placeholder logic - Replace with real generation
    return "Generated overview content."

def generate_student_criteria_content(client, student_info, is_new_api):
    return "Student criteria summary."

def generate_college_list(client, student_info, is_new_api):
    return {"Reach": ["Stanford", "MIT"], "Target": ["UCLA", "USC"], "Safety": ["CSU Fullerton", "UC Riverside"]}

def generate_reasons_for_selections(client, student_info, college_list, is_new_api):
    return {level: {school: f"Reason why {school} is a {level} school." for school in schools}
            for level, schools in college_list.items()}

def generate_detailed_college_info(client, student_info, college_list, level, is_new_api):
    return {school: [f"{school} - Detail A", f"{school} - Detail B"] for school in college_list.get(level, [])}

def create_new_document(first_name, last_name):
    filename = f"{first_name}_{last_name}_College_Recommendations.docx"
    doc = Document()
    doc.add_heading(f"{first_name} {last_name} – College List Report", level=1)
    return doc, filename

def create_overview_section(doc, overview_text):
    doc.add_heading("Overview", level=2)
    doc.add_paragraph(overview_text)

def create_student_criteria_section(doc, criteria_text):
    doc.add_heading("Student College Criteria", level=2)
    doc.add_paragraph(criteria_text)

def create_college_list_table(doc, college_list, student_info):
    doc.add_heading("College List Summary", level=2)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'School Level'
    hdr_cells[1].text = 'List of Colleges'
    for level, schools in college_list.items():
        row_cells = table.add_row().cells
        row_cells[0].text = level
        row_cells[1].text = "\n".join(schools)

def create_reasons_section(doc, reasons_dict):
    doc.add_heading("Reasons for School Classification", level=2)
    for level, school_reasons in reasons_dict.items():
        doc.add_heading(f"{level} Schools", level=3)
        table = doc.add_table(rows=1, cols=2)
        table.rows[0].cells[0].text = "College"
        table.rows[0].cells[1].text = "Why It’s a {level}"
        for school, reason in school_reasons.items():
            row = table.add_row().cells
            row[0].text = school
            row[1].text = reason

def create_detailed_college_info(doc, detailed_info_dict):
    doc.add_heading("College Detailed Information", level=2)
    for level, school_data in detailed_info_dict.items():
        for school, bullets in school_data.items():
            doc.add_heading(school, level=3)
            for item in bullets:
                doc.add_paragraph(item, style='List Bullet')
