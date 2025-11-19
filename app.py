import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

# --- Page Configuration ---
st.set_page_config(page_title="Thai Contract Automation", page_icon="📝", layout="wide")

st.title("📝 Thai Contract Automation Tool")
st.markdown("Upload your contract template (.docx) and your data (.csv) to generate contracts automatically.")

# --- Sidebar: Uploads ---
st.sidebar.header("1. Upload Files")
docx_file = st.sidebar.file_uploader("Upload Contract Template (.docx)", type=["docx"])
csv_file = st.sidebar.file_uploader("Upload Data (.csv)", type=["csv"])

# --- Main Logic ---
if docx_file and csv_file:
    # 1. Read CSV
    try:
        # Attempt to read CSV (handle potential encoding issues common with Thai)
        try:
            df = pd.read_csv(csv_file, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding='tis-620') # Common legacy Thai encoding
        
        st.success("Files uploaded successfully!")
        
        # --- Column Mapping Interface ---
        st.subheader("2. Map CSV Columns to Contract Fields")
        st.info("Select which column in your CSV corresponds to the required field in the contract.")
        
        # Get column names
        columns = df.columns.tolist()
        
        # Create 2 columns for layout
        col1, col2 = st.columns(2)
        
        with col1:
            # Intelligent defaults based on user prompt description (0-based index logic)
            # User said: Prefix=3rd(idx 2), Name=4th(idx 3), Surname=5th(idx 4), ID=6th(idx 5), Address=10th(idx 9)
            
            col_prefix = st.selectbox("Column for Prefix (นาย/นางสาว)", columns, index=2 if len(columns) > 2 else 0)
            col_name = st.selectbox("Column for First Name", columns, index=3 if len(columns) > 3 else 0)
            col_surname = st.selectbox("Column for Surname", columns, index=4 if len(columns) > 4 else 0)
            
        with col2:
            col_id = st.selectbox("Column for ID Card (Identification Number)", columns, index=5 if len(columns) > 5 else 0)
            col_address = st.selectbox("Column for Address", columns, index=9 if len(columns) > 9 else 0)
            # Signing name usually combines First + Last, but we allow mapping just in case
            use_full_name_sign = st.checkbox("Auto-combine Name + Surname for Signature?", value=True)
        
        st.markdown("---")
        
        # --- Selection Interface ---
        st.subheader("3. Select Person to Generate Contract")
        
        # Create a display label for the dropdown
        df['display_label'] = df[col_prefix].astype(str) + " " + df[col_name].astype(str) + " " + df[col_surname].astype(str)
        
        selected_person_label = st.selectbox("Select a person from the list:", df['display_label'])
        
        # Filter data for the selected person
        selected_row = df[df['display_label'] == selected_person_label].iloc[0]
        
        # --- Preview Data ---
        with st.expander("View Data to be filled"):
            st.write(f"**Prefix:** {selected_row[col_prefix]}")
            st.write(f"**Name:** {selected_row[col_name]}")
            st.write(f"**Surname:** {selected_row[col_surname]}")
            st.write(f"**ID Card:** {selected_row[col_id]}")
            st.write(f"**Address:** {selected_row[col_address]}")
            
            sign_name_val = f"{selected_row[col_name]} {selected_row[col_surname]}" if use_full_name_sign else f"{selected_row[col_name]}"
            st.write(f"**Signature Name:** {sign_name_val}")

        # --- Generation Button ---
        if st.button("Generate Contract"):
            # Load the Docx Template
            doc = DocxTemplate(docx_file)
            
            # Prepare context dictionary
            context = {
                'prefix': str(selected_row[col_prefix]),
                'name': str(selected_row[col_name]),
                'surname': str(selected_row[col_surname]),
                'id_card': str(selected_row[col_id]),
                'address': str(selected_row[col_address]),
                'sign_name': sign_name_val
            }
            
            # Render the document
            doc.render(context)
            
            # Save to a memory stream
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            file_name = f"Contract_{selected_row[col_name]}_{selected_row[col_surname]}.docx"
            
            st.success(f"Contract generated for {selected_person_label}!")
            
            st.download_button(
                label="⬇️ Download Filled Contract (.docx)",
                data=bio,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
    except Exception as e:
        st.error(f"An error occurred processing the file: {e}")

else:
    st.info("Please upload both the .docx template and .csv data file to begin.")
