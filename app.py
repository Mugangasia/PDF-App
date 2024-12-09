import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
import re
from PyPDF2 import PdfReader
import io
import base64

def extract_text_from_pdf(pdf_file):
    """Extract text from uploaded PDF file"""
    reader = PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def extract_numbered_items(text):
    """Extract numbered items and their descriptions"""
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    topics = []
    descriptions = []
    
    current_topic = ""
    current_description = []
    number_pattern = r'^\d+\.'
    
    for line in lines:
        if line == "Georgette Review" or "Study online" in line:
            continue
            
        if re.match(number_pattern, line):
            if current_topic and current_description:
                topics.append(current_topic)
                descriptions.append(' '.join(current_description))
            
            parts = re.split(number_pattern, line, maxsplit=1)
            if len(parts) > 1:
                current_topic = parts[1].strip()
                current_description = []
        else:
            if line and not line.endswith('/'):
                current_description.append(line)
    
    if current_topic and current_description:
        topics.append(current_topic)
        descriptions.append(' '.join(current_description))
    
    return topics, descriptions

def create_formatted_excel(topics, descriptions):
    """Create formatted Excel file in memory"""
    output = io.BytesIO()
    
    # Create DataFrame
    df = pd.DataFrame({
        'Topic': topics,
        'Description': descriptions
    })
    
    # Create Excel writer object
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    # Write DataFrame to Excel
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # Get the workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Define border style
    border = Border(bottom=Side(style='thin'))
    
    # Format columns
    for idx in range(len(topics) + 1):
        # Apply bold to first column
        cell = worksheet.cell(row=idx+1, column=1)
        cell.font = Font(bold=True)
        
        # Add horizontal lines
        for col in range(1, 3):
            cell = worksheet.cell(row=idx+1, column=col)
            cell.border = border
    
    # Adjust column widths
    worksheet.column_dimensions['A'].width = 30
    worksheet.column_dimensions['B'].width = 50
    
    # Save the file
    writer.close()
    
    # Reset pointer to start of file
    output.seek(0)
    return output

def get_download_link(excel_file, filename):
    """Generate a download link for the Excel file"""
    b64 = base64.b64encode(excel_file.getvalue()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'

# Set up the Streamlit page
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄")

# Add title and description
st.title("📄 PDF to Excel Converter")
st.markdown("""
This app converts structured PDF documents into formatted Excel files.
Upload your PDF file below to get started.
""")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    try:
        # Show processing message
        with st.spinner('Processing PDF...'):
            # Extract text from PDF
            pdf_text = extract_text_from_pdf(uploaded_file)
            
            # Extract topics and descriptions
            topics, descriptions = extract_numbered_items(pdf_text)
            
            # Create preview of the data
            df_preview = pd.DataFrame({
                'Topic': topics,
                'Description': descriptions
            })
            
            # Show preview
            st.subheader("Preview of Extracted Data")
            st.dataframe(df_preview)
            
            # Create formatted Excel file
            excel_file = create_formatted_excel(topics, descriptions)
            
            # Create download button
            st.subheader("Download Excel File")
            output_filename = uploaded_file.name.replace('.pdf', '_converted.xlsx')
            st.markdown(get_download_link(excel_file, output_filename), unsafe_allow_html=True)
            
            # Show success message
            st.success("✅ PDF processed successfully!")
            
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.markdown("Please make sure the PDF file is in the correct format and try again.")

# Add usage instructions
with st.expander("📖 How to Use"):
    st.markdown("""
    1. Click the 'Browse files' button above
    2. Select your PDF file
    3. Wait for the processing to complete
    4. Preview the extracted data
    5. Click the download link to get your Excel file
    
    The Excel file will include:
    - Bold text for topics
    - Horizontal lines between rows
    - Adjusted column widths for better readability
    """)

# Add footer
st.markdown("---")
st.markdown("Made with ❤️ by Bravin Mugangasia")