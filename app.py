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
            # Intelligent defaults based on user prompt (0-based index logic)
            # Prefix=3rd(idx 2), Name=4th(idx 3), Surname=5th(idx 4)
            col_prefix = st.selectbox("Column for Prefix (นาย/นางสาว)", columns, index=2 if len(columns) > 2 else 0)
            col_name = st.selectbox("Column for First Name", columns, index=3 if len(columns) > 3 else 0)
            col_surname = st.selectbox("Column for Surname", columns, index=4 if len(columns) > 4 else 0)
            
        with col2:
            # ID=6th(idx 5), Address=10th(idx 9)
            col_id = st.selectbox("Column for ID Card (Identification Number)", columns, index=5 if len(columns) > 5 else 0)
            col_address = st.selectbox("Column for Address", columns, index=9 if len(columns) > 9 else 0)
        
        st.markdown("---")
        
        # --- Selection Interface ---
        st.subheader("3. Select Person to Generate Contract")
        
        # Create a display label for the dropdown (Full Name)
        df['display_label'] = df[col_prefix].astype(str) + " " + df[col_name].astype(str) + " " + df[col_surname].astype(str)
        
        selected_person_label = st.selectbox("Select a person from the list:", df['display_label'])
        
        # Filter data for the selected person
        selected_row = df[df['display_label'] == selected_person_label].iloc[0]

        # --- Construct the data variables ---
        p_prefix = str(selected_row[col_prefix])
        p_name = str(selected_row[col_name])
        p_surname = str(selected_row[col_surname])
        p_id = str(selected_row[col_id])
        p_address = str(selected_row[col_address])
        
        # UPDATED: Sign name now includes Prefix + Name + Surname
        p_sign_name = f"{p_prefix} {p_name} {p_surname}"
        
        # --- Preview Data ---
        with st.expander("View Data to be filled"):
            st.write(f"**Prefix:** {p_prefix}")
            st.write(f"**Name:** {p_name}")
            st.write(f"**Surname:** {p_surname}")
            st.write(f"**ID Card:** {p_id}")
            st.write(f"**Address:** {p_address}")
            st.markdown(f"**Signature Name:** `{p_sign_name}`") # highlighted for visibility

        # --- Generation Button ---
        if st.button("Generate Contract"):
            # Load the Docx Template
            doc = DocxTemplate(docx_file)
            
            # Prepare context dictionary (Matches {{ }} in Word file)
            context = {
                'prefix': p_prefix,
                'name': p_name,
                'surname': p_surname,
                'id_card': p_id,
                'address': p_address,
                'sign_name': p_sign_name
            }
            
            # Render the document
            doc.render(context)
            
            # Save to a memory stream
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            file_name = f"Contract_{p_name}_{p_surname}.docx"
            
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
