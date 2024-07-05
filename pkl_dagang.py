import streamlit as st
from docx import Document
from io import BytesIO
import pandas as pd

# Load the combined Excel file
file_path_combined = 'Combined_Questions.xlsx'
df_combined = pd.read_excel(file_path_combined)

# Load the new Excel file with additional data
file_path_latest = 'Elemen_dagang.xlsx'
df_latest = pd.read_excel(file_path_latest)

# Extract Question_Number as a list of integers
question_numbers = df_latest['Question_Number'].astype(int).tolist()

# Function to replace hashtags with user input
def replace_hashtags(doc_path, replacements):
    doc = Document(doc_path)
    for p in doc.paragraphs:
        for hashtag, replacement in replacements.items():
            if hashtag in p.text:
                p.text = p.text.replace(hashtag, str(replacement))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for hashtag, replacement in replacements.items():
                    if hashtag in cell.text:
                        cell.text = cell.text.replace(hashtag, str(replacement))
    return doc

# Function to convert document to BytesIO
def convert_to_bytes(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit form
st.title("Form Pengisian Data")

# Initialize session state for each question if not already present
for _, row in df_combined.iterrows():
    hashtag = row['Penanda'] if row['Source'] == "Data Diri" else row['Answer Yes']
    if hashtag not in st.session_state:
        st.session_state[hashtag] = ""

num_questions = len(df_combined)
num_pages = (num_questions // 7) + (1 if num_questions % 7 != 0 else 0)

# Create expanders for each "page" of questions
for page in range(1, num_pages + 1):
    start_index = (page - 1) * 7
    end_index = min(start_index + 7, num_questions)

    with st.expander(f"Halaman {page}"):
        for i, (index, row) in enumerate(df_combined.iloc[start_index:end_index].iterrows(), start=1):
            actual_question_number = start_index + i + 2

            # Check if we need to insert a subheader
            if actual_question_number in question_numbers:
                subheader = df_latest[df_latest['Question_Number'] == actual_question_number]['Capaian Pembelajaran'].values[0]
                st.subheader(subheader)

            question = row['Pertanyaan']
            hashtag = row['Penanda'] if row['Source'] == "Data Diri" else row['Answer Yes']
            tipe = row['Tipe']

            if tipe == 'short description':
                st.session_state[hashtag] = st.text_input(question, value=st.session_state[hashtag], key=f"input_{hashtag}")
            elif tipe == 'paragraph':
                st.session_state[hashtag] = st.text_area(question, value=st.session_state[hashtag], key=f"area_{hashtag}")
            elif tipe == 'date':
                date_value = pd.to_datetime(st.session_state[hashtag], errors='coerce')
                if pd.isna(date_value):
                    date_value = None
                st.session_state[hashtag] = st.date_input(question, value=date_value, key=f"date_{hashtag}")
            elif tipe == 'radio':
                options = ['Belum memilih', 'Sekolah', 'Iduka']
                index = options.index(st.session_state[hashtag]) if st.session_state[hashtag] in options else 0
                st.session_state[hashtag] = st.radio(question, options, index=index, key=f"radio_{hashtag}")

            # Add a separator between questions
            st.markdown("---")

# Place the submit button outside the expanders
submit_button = st.button(label="Generate Document")

if submit_button:
    all_filled = all(value != "" and value is not None for value in st.session_state.values())
    if all_filled:
        # Replace hashtags in the document
        doc_path = "Template_Document.docx"  # Ensure this file is in the same directory
        updated_doc = replace_hashtags(doc_path, st.session_state)

        # Convert document to BytesIO
        bytes_io = convert_to_bytes(updated_doc)

        # Provide download link
        st.download_button(label="Download file disini", data=bytes_io, file_name="Filled_Document.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.warning("Harap selesaikan semua pertanyaan sebelum mengunduh PDF.")
