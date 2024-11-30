import os
from flask import Flask, request, render_template, send_file
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['OUTPUT_FOLDER'] = 'output/'

# Create necessary folders
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

if not os.path.exists(app.config['OUTPUT_FOLDER']):
    os.makedirs(app.config['OUTPUT_FOLDER'])

@app.route('/')
def upload_page():
    return render_template('upload.html', download_link=None)

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return render_template('upload.html', error="No file uploaded!", download_link=None)

    file = request.files['file']
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    output_path = process_excel(file_path)
    download_link = f"/output/{os.path.basename(output_path)}"

    return render_template('upload.html', download_link=download_link)

@app.route('/output/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename), as_attachment=True)

def process_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    metadata = extract_metadata(df)

    if metadata:
        question_start_row = find_question_start(df)  # Find survey start row
        df_survey = pd.read_excel(file_path, skiprows=question_start_row)
    else:
        df_survey = pd.read_excel(file_path)  # Read directly, assume first row is header
        metadata = {"Subject Code": "Unknown", "Subject Name": "Unknown", "Branch": "Unknown", "Year": "Unknown"}

    df_survey.drop(columns=['NAME', 'USN'], errors='ignore', inplace=True)

    course_code = metadata.get('Subject Code', 'Unknown')

    for idx, col in enumerate(df_survey.columns[1:], start=1):
        course_code_column_name = f"{course_code}.{idx}"  # Use course code + question number
        df_survey.rename(columns={col: course_code_column_name}, inplace=True)

    total_students = len(df_survey)
    summary = []

    for question in df_survey.columns[1:]:  # Skip first column (course code column)
        responses = df_survey[question].value_counts().to_dict()
        excellent = responses.get('Excellent', 0)
        very_good = responses.get('Very Good', 0)
        good = responses.get('Good', 0)
        satisfactory = responses.get('Satisfactory', 0)
        poor = responses.get('Poor', 0)
        evg = excellent + very_good + good
        percentage = (evg / total_students) * 100 if total_students > 0 else 0

        summary.append({
            "Subject Code": question,  # Subject code (e.g., 23EVS127.1)
            "Excellent": excellent,
            "Very Good": very_good,
            "Good": good,
            "Satisfactory": satisfactory,
            "Poor": poor,
            "E+V+G": evg,
            "%": round(percentage, 2),  # Add percentage of Excellent + Very Good + Good
        })

    output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{metadata.get('Subject Code', 'Unknown')}_analysis.docx")
    doc = Document()

    institute_name = "B.N.M. Institute of Technology, Bengaluru-70"
    para = doc.add_paragraph(institute_name, style="Heading 1")
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.runs[0].bold = True
    para.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    department_name = "Department of Chemistry"
    para = doc.add_paragraph(department_name, style="Heading 1")
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.runs[0].bold = True
    para.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    analysis_text = "Analysis of course exit survey"
    para = doc.add_paragraph(analysis_text)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.runs[0].bold = True
    para.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    course_info = f"Course Code: {metadata.get('Subject Code')}    Course: {metadata.get('Subject Name')}    Branch: {metadata.get('Branch')}    Year: {metadata.get('Year')}"
    para = doc.add_paragraph(course_info)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.runs[0].bold = True
    para.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    total_students_text = f"TOTAL NO OF STUDENTS TAKEN SURVEY= {total_students}"
    para = doc.add_paragraph(total_students_text)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.runs[0].bold = True
    para.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=8)  # Added one more column for Percentage
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = metadata.get('Subject Code', 'Unknown')
    hdr_cells[1].text = 'EXCELLENT'
    hdr_cells[2].text = 'VERY GOOD'
    hdr_cells[3].text = 'GOOD'
    hdr_cells[4].text = 'SATISFACTORY'
    hdr_cells[5].text = 'POOR'
    hdr_cells[6].text = 'E+V+G'
    hdr_cells[7].text = '%'

    for row in summary:
        row_cells = table.add_row().cells
        row_cells[0].text = row["Subject Code"]
        row_cells[1].text = str(row["Excellent"])
        row_cells[2].text = str(row["Very Good"])
        row_cells[3].text = str(row["Good"])
        row_cells[4].text = str(row["Satisfactory"])
        row_cells[5].text = str(row["Poor"])
        row_cells[6].text = str(row["E+V+G"])
        row_cells[7].text = str(row["%"])

    doc.save(output_path)
    return output_path

def extract_metadata(df):
    metadata = {}

    for i, row in df.iterrows():
        for val in row:
            if isinstance(val, str):
                if 'Subject Name' in val:
                    metadata['Subject Name'] = val.split(":")[-1].strip()
                elif 'Subject Code' in val:
                    metadata['Subject Code'] = val.split(":")[-1].strip()
                elif 'Branch' in val:
                    metadata['Branch'] = val.split(":")[-1].strip()
                elif 'Year' in val:
                    metadata['Year'] = val.split(":")[-1].strip()

    return metadata

def find_question_start(df):
    for i, row in df.iterrows():
        if 'Question' in str(row[0]):
            return i
    return 7  # Default start if nothing is found

if __name__ == '__main__':
    # Use the dynamic port from Render
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 5000)), debug=True)
