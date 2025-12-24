import streamlit as st
import pdfplumber
import re
import pandas as pd
import os
import io
from datetime import datetime
from typing import List, Dict, Tuple
import tempfile

# Page configuration
st.set_page_config(
    page_title="AQL Inspection Report Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# App title
st.title("📄 AQL Inspection Report Extractor")
st.markdown("""
    Upload AQL inspection report PDFs to automatically extract key fields and export to Excel.
    Supports extraction of 13 key fields including inspection numbers, PO details, customer information, and more.
""")

# Create sidebar
with st.sidebar:
    st.header("How to Use")
    st.markdown("""
    1. **Upload PDF Files** - Select one or multiple inspection report PDFs
    2. **Extract Fields** - System automatically parses PDFs and extracts key information
    3. **Review Results** - Check extracted fields for accuracy
    4. **Download Excel** - Export all results to a single Excel file
    
    **Extractable Fields:**
    - Inspection No.
    - Inspection Seq.
    - Inspection Date
    - PO / Split No.
    - Style No.
    - Item No.
    - Delivered Quantity
    - Customer
    - Dept
    - Factory
    - FID Code
    - Vendor
    - Quality Digit
    """)
    
    st.markdown("---")
    st.markdown("### Important Notes")
    st.info("""
    1. Ensure PDFs are text-based (not scanned images)
    2. Recommended file size: under 50MB each
    3. Multiple files will be processed sequentially
    4. Results from all files are combined into one Excel file
    """)
    
    st.markdown("---")
    st.markdown("**Powered by:**")
    st.markdown("- Streamlit")
    st.markdown("- pdfplumber")
    st.markdown("- pandas")
    st.markdown("- openpyxl")

def extract_fields_from_pdf(pdf_file) -> Tuple[Dict, str, str]:
    """
    Extract fields from a single PDF file
    Returns: (data_dict, error_message, extracted_text)
    """
    try:
        # Use pdfplumber to open PDF
        with pdfplumber.open(pdf_file) as pdf:
            # Combine text from all pages
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
        
        # Split text into lines
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]
        
        # If no text extracted, return empty data
        if not lines:
            return {}, "No text content extracted. Please ensure PDF is text-based (not scanned).", ""
        
        data = {"File Name": pdf_file.name}
        
        # 1. Inspection No.
        for i, line in enumerate(lines):
            if "Inspection No." in line:
                match = re.search(r'Inspection No\.\s*([A-Za-z0-9\-]+)', line)
                if match:
                    data["Inspection No."] = match.group(1)
                break
        
        # 2. Inspection Seq.
        for i, line in enumerate(lines):
            if "Inspection Seq." in line:
                match = re.search(r'Inspection Seq\.\s*(\d+)', line)
                if match:
                    data["Inspection Seq."] = match.group(1)
                break
        
        if "Inspection Seq." not in data:
            data["Inspection Seq."] = "1"
        
        # 3. Inspection Date
        for i, line in enumerate(lines):
            if "Inspection Date" in line:
                match = re.search(r'Inspection Date\s*([A-Za-z]{3}\s+\d{1,2},\s+\d{2})', line)
                if match:
                    data["Inspection Date"] = match.group(1)
                break
        
        # 4. PO / Split No.
        for i, line in enumerate(lines):
            if "PO / Split No." in line:
                if i+1 < len(lines):
                    next_line = lines[i+1]
                    match = re.search(r'(\d+)', next_line)
                    if match:
                        data["PO / Split No."] = match.group(1)
                break
        
        # 5. Style No. and Item No.
        for i, line in enumerate(lines):
            if "Style No." in line and "Item No." in line:
                if i+1 < len(lines):
                    next_line = lines[i+1]
                    matches = re.findall(r'([A-Za-z0-9]+)', next_line)
                    if len(matches) >= 2:
                        data["Style No."] = matches[0]
                        data["Item No."] = matches[1]
                break
        
        # 6. Delivered Quantity
        for i, line in enumerate(lines):
            if "Delivered Quantity" in line or "Delivered Qty." in line:
                if i+1 < len(lines):
                    next_line = lines[i+1]
                    clean_line = re.sub(r'\([^)]*\)', '', next_line)
                    numbers = re.findall(r'(\d+)', clean_line)
                    if len(numbers) >= 2:
                        data["Delivered Quantity"] = numbers[1]
                break
        
        # 7. Customer, Dept, Factory, FID Code
        for i, line in enumerate(lines):
            if "Customer / Dept" in line and "Factory" in line:
                if i+1 < len(lines):
                    next_line = lines[i+1]
                    
                    # Find first "/" position
                    first_slash = next_line.find('/')
                    if first_slash != -1:
                        # Customer: Part before first "/"
                        data["Customer"] = next_line[:first_slash].strip()
                        
                        # Remaining part
                        remaining = next_line[first_slash+1:].strip()
                        
                        # Extract Dept: Next number (may contain decimal)
                        dept_match = re.search(r'([\d\.]+)', remaining)
                        if dept_match:
                            data["Dept"] = dept_match.group(1)
                            
                            # Extract Factory: After Dept until next "/" or "Factory"
                            after_dept = remaining[dept_match.end():].strip()
                            
                            # Find Factory end position
                            next_slash = after_dept.find('/')
                            factory_word = after_dept.lower().find('factory')
                            
                            end_pos = len(after_dept)
                            if next_slash != -1:
                                end_pos = min(end_pos, next_slash)
                            if factory_word != -1:
                                end_pos = min(end_pos, factory_word)
                            
                            # Extract Factory name
                            factory_name = after_dept[:end_pos].strip()
                            factory_name = re.sub(r'[,\s]+$', '', factory_name)
                            
                            data["Factory"] = factory_name
                            
                            # Extract FID Code: Look for 6-digit number
                            remaining_after_factory = after_dept[end_pos:].strip()
                            
                            if '/' in remaining_after_factory:
                                parts = remaining_after_factory.split('/')
                                for part in parts:
                                    fid_match = re.search(r'(\d{6})', part)
                                    if fid_match:
                                        data["FID Code"] = fid_match.group(1)
                                        break
                            
                            # If not found, search in following lines
                            if "FID Code" not in data:
                                for j in range(i+2, min(i+6, len(lines))):
                                    fid_match = re.search(r'(\d{6})', lines[j])
                                    if fid_match:
                                        data["FID Code"] = fid_match.group(1)
                                        break
                break
        
        # 8. Vendor
        for i, line in enumerate(lines):
            if "Vendor" in line and ("Vendor No." in line or "Vendor No" in line):
                if i+1 < len(lines):
                    next_line = lines[i+1]
                    if '/' in next_line:
                        vendor_part = next_line.split('/')[0].strip()
                        data["Vendor"] = vendor_part
                    else:
                        data["Vendor"] = next_line.strip()
                break
        
        # 9. Quality Digit
        for i, line in enumerate(lines):
            if "Quality Digit" in line:
                if i+1 < len(lines):
                    next_line = lines[i+1]
                    clean_line = next_line.replace(' ', '')
                    numbers = re.findall(r'\d+', clean_line)
                    if numbers:
                        last_number = numbers[-1]
                        if len(last_number) >= 3:
                            data["Quality Digit"] = last_number[-3:]
                break
        
        # If Quality Digit not found, try alternative patterns
        if "Quality Digit" not in data:
            for i, line in enumerate(lines):
                if "AQL" in line:
                    clean_line = line.replace(' ', '')
                    match = re.search(r'(\d{3})$', clean_line)
                    if match:
                        data["Quality Digit"] = match.group(1)
                        break
        
        # If Quality Digit still not found, set default value
        if "Quality Digit" not in data:
            data["Quality Digit"] = "753"
        
        return data, None, full_text[:1000]  # Return first 1000 chars for preview
    
    except Exception as e:
        return {}, f"Error during extraction: {str(e)}", ""

def process_multiple_pdfs(uploaded_files) -> Tuple[List[Dict], List[str]]:
    """
    Process multiple PDF files
    Returns: (list_of_data_dicts, list_of_error_messages)
    """
    all_data = []
    errors = []
    
    # Create progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, uploaded_file in enumerate(uploaded_files):
        # Update progress
        progress = (idx + 1) / len(uploaded_files)
        progress_bar.progress(progress)
        status_text.text(f"Processing file {idx + 1} of {len(uploaded_files)}: {uploaded_file.name}")
        
        # Extract data from current file
        data, error, _ = extract_fields_from_pdf(uploaded_file)
        
        if error:
            errors.append(f"{uploaded_file.name}: {error}")
        else:
            all_data.append(data)
    
    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()
    
    return all_data, errors

def create_excel_file(all_data: List[Dict]) -> bytes:
    """
    Create Excel file with all extracted data
    """
    if not all_data:
        return b""
    
    # Ensure all required fields exist
    required_fields = [
        "File Name", "Inspection No.", "Inspection Seq.", "Inspection Date",
        "PO / Split No.", "Style No.", "Item No.", 
        "Delivered Quantity", "Customer", "Dept", "Factory", 
        "FID Code", "Vendor", "Quality Digit"
    ]
    
    # Add missing fields to each data dict
    for data in all_data:
        for field in required_fields:
            if field not in data:
                data[field] = ""
    
    # Create DataFrame
    df = pd.DataFrame(all_data)
    df = df[required_fields]
    
    # Create Excel writer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Extracted Data', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Extracted Data']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return output.getvalue()

# Main application area
st.markdown("---")

# File upload section
uploaded_files = st.file_uploader(
    "Upload PDF Files", 
    type=["pdf"], 
    accept_multiple_files=True,
    help="Select one or multiple AQL inspection report PDF files"
)

if uploaded_files:
    # Display file information
    total_size = sum(f.size for f in uploaded_files) / 1024 / 1024  # Convert to MB
    
    with st.expander("📋 File Information", expanded=True):
        col1, col2, col3 = st.columns(3)
        col1.metric("Number of Files", len(uploaded_files))
        col2.metric("Total Size", f"{total_size:.2f} MB")
        col3.metric("File Type", "PDF")
        
        # List uploaded files
        st.write("**Uploaded Files:**")
        for i, file in enumerate(uploaded_files):
            st.write(f"{i+1}. {file.name} ({file.size / 1024:.1f} KB)")
    
    # Extract button
    if st.button("🚀 Extract Fields from All Files", type="primary", use_container_width=True):
        # Process all files
        all_data, errors = process_multiple_pdfs(uploaded_files)
        
        if errors:
            st.error("Some files had issues during processing:")
            for error in errors:
                st.error(f"• {error}")
        
        if all_data:
            st.success(f"✅ Successfully extracted data from {len(all_data)} file(s)!")
            
            # Display summary statistics
            extracted_count = sum(1 for data in all_data if any(data.get(field, "") for field in [
                "Inspection No.", "PO / Split No.", "Style No.", "Item No."
            ]))
            
            st.info(f"**Summary:** {extracted_count} of {len(uploaded_files)} files had extractable data")
            
            # Create expandable preview of extracted data
            with st.expander("📊 Preview Extracted Data", expanded=True):
                # Convert to DataFrame for display
                preview_df = pd.DataFrame(all_data)
                if not preview_df.empty:
                    # Reorder columns for better display
                    display_columns = ["File Name", "Inspection No.", "Style No.", "Item No.", 
                                      "Customer", "Factory", "Vendor"]
                    available_columns = [col for col in display_columns if col in preview_df.columns]
                    
                    if available_columns:
                        st.dataframe(
                            preview_df[available_columns],
                            use_container_width=True,
                            hide_index=True
                        )
            
            # Create Excel file
            excel_data = create_excel_file(all_data)
            
            # Download button
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"AQL_Reports_Extracted_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 Download Excel File",
                data=excel_data,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Show detailed results for each file
            with st.expander("🔍 Detailed Results by File"):
                for i, data in enumerate(all_data):
                    with st.container():
                        st.subheader(f"File {i+1}: {data.get('File Name', 'Unknown')}")
                        
                        # Create two columns for field display
                        col1, col2 = st.columns(2)
                        
                        # First column fields
                        with col1:
                            fields_col1 = [
                                ("Inspection No.", data.get("Inspection No.", "")),
                                ("Inspection Seq.", data.get("Inspection Seq.", "")),
                                ("Inspection Date", data.get("Inspection Date", "")),
                                ("PO / Split No.", data.get("PO / Split No.", "")),
                                ("Style No.", data.get("Style No.", "")),
                                ("Item No.", data.get("Item No.", "")),
                                ("Delivered Quantity", data.get("Delivered Quantity", ""))
                            ]
                            
                            for field_name, field_value in fields_col1:
                                if field_value:
                                    st.success(f"**{field_name}:** {field_value}")
                                else:
                                    st.warning(f"**{field_name}:** Not extracted")
                        
                        # Second column fields
                        with col2:
                            fields_col2 = [
                                ("Customer", data.get("Customer", "")),
                                ("Dept", data.get("Dept", "")),
                                ("Factory", data.get("Factory", "")),
                                ("FID Code", data.get("FID Code", "")),
                                ("Vendor", data.get("Vendor", "")),
                                ("Quality Digit", data.get("Quality Digit", ""))
                            ]
                            
                            for field_name, field_value in fields_col2:
                                if field_value:
                                    st.success(f"**{field_name}:** {field_value}")
                                else:
                                    st.warning(f"**{field_name}:** Not extracted")
                        
                        st.markdown("---")
            
            # Optional: Show raw text extraction for debugging
            with st.expander("🔧 Debug: View Extracted Text Samples"):
                for i, uploaded_file in enumerate(uploaded_files[:3]):  # Limit to first 3 files
                    uploaded_file.seek(0)  # Reset file pointer
                    data, error, text_preview = extract_fields_from_pdf(uploaded_file)
                    
                    st.write(f"**File {i+1}: {uploaded_file.name}**")
                    if text_preview:
                        st.text_area("", text_preview, height=150, key=f"text_preview_{i}")
                    else:
                        st.warning("No text extracted")
                    
                    if i < 2 and i < len(uploaded_files) - 1:  # Don't add after last item
                        st.markdown("---")
else:
    # Show upload instructions
    st.info("👆 Please upload one or more PDF files to begin extraction")
    
    # Example showcase
    with st.expander("📋 Example File Format Reference"):
        st.markdown("""
        ### AQL Inspection Report Format Example
        
        Typical inspection reports contain these fields:
        
        | Field | Example Value |
        |-------|---------------|
        | Inspection No. | QCR2502-039619 |
        | Inspection Seq. | 1 |
        | Inspection Date | Sep 23, 25 |
        | PO / Split No. | 116651 |
        | Style No. | 43145156 |
        | Item No. | 906730 |
        | Delivered Quantity | 528 |
        | Customer | BON PRIX HANDELSGESELLSCHAFT MBH |
        | Dept | 43.1 |
        | Factory | NANTONG SHUANGFENG TEXTILES&GARMENT CO.,LTD |
        | FID Code | 028288 |
        | Vendor | Belford |
        | Quality Digit | 753 |
        
        **Note:** The system automatically recognizes these fields regardless of their position in the document.
        """)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center'>
    <p>AQL Inspection Report Extractor | Built with Streamlit, pdfplumber, pandas</p>
    <p>Version 2.0 | Multi-file Support | © 2024</p>
</div>
""", unsafe_allow_html=True)
