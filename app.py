import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import re

# --- Page Configuration ---
st.set_page_config(page_title="Thai Contract Automation", page_icon="📝", layout="wide")

st.title("📝 Thai Contract Automation Tool")
st.markdown("Upload your contract template (.docx) and your data (.csv) to generate contracts.")

# --- Helper Function for Auto-Increment ---
def increment_id(start_id, index):
    # Find the last number in the string
    match = re.search(r'(\d+)$', start_id)
    if match:
        number_str = match.group(1)
        number_len = len(number_str)
        start_num = int(number_str)
        
        # Calculate new number
        new_num = start_num + index
        
        # Pad with zeros to match original length (e.g. 001 -> 002)
        new_number_str = str(new_num).zfill(number_len)
        
        # Replace the old number with the new number
        return start_id[:match.start()] + new_number_str + start_id[match.end():]
    else:
        # If no number found at end, just append -1, -2, etc.
        return f"{start_id}-{index + 1}"

# --- Sidebar: Uploads ---
st.sidebar.header("1. Upload Files")
docx_file = st.sidebar.file_uploader("Upload Contract Template (.docx)", type=["docx"])
csv_file = st.sidebar.file_uploader("Upload Data (.csv)", type=["csv"])

# --- Main Logic ---
if docx_file and csv_file:
    try:
        try:
            df = pd.read_csv(csv_file, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding='tis-620')
        
        st.success("Files uploaded successfully!")
        
        # --- Global Contract Settings ---
        st.subheader("2. Contract ID Settings")
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            # User inputs the FIRST ID
            start_contract_id = st.text_input(
                "Enter Starting Contract ID (e.g. CT-2568/001)", 
                value="CT-2568/001",
                help="The system will find the last number and increment it for each person selected."
            )
        
        # --- Column Mapping Interface ---
        st.subheader("3. Map CSV Columns")
        with st.expander("Click to adjust column mapping", expanded=False):
            columns = df.columns.tolist()
            col1, col2 = st.columns(2)
            with col1:
                col_prefix = st.selectbox("Prefix (นาย/นางสาว)", columns, index=2 if len(columns) > 2 else 0)
                col_name = st.selectbox("First Name", columns, index=3 if len(columns) > 3 else 0)
                col_surname = st.selectbox("Surname", columns, index=4 if len(columns) > 4 else 0)
            with col2:
                col_id = st.selectbox("ID Card", columns, index=5 if len(columns) > 5 else 0)
                col_address = st.selectbox("Address", columns, index=9 if len(columns) > 9 else 0)

        st.markdown("---")
        
        # --- Bulk Selection Interface ---
        st.subheader("4. Select People for Contracts")
        
        # Insert Select column
        df_with_selections = df.copy()
        df_with_selections.insert(0, "Select", False)
        
        # Display data editor
        edited_df = st.data_editor(
            df_with_selections,
            column_config={
                "Select": st.column_config.CheckboxColumn("Generate?", default=False)
            },
            disabled=df.columns,
            hide_index=True,
        )
        
        # Filter selected rows
        selected_rows = edited_df[edited_df.Select]
        
        # --- Generation Logic ---
        st.write(f"**Selected: {len(selected_rows)} people**")
        
        if st.button("Generate Selected Contracts", type="primary"):
            if len(selected_rows) == 0:
                st.warning("Please select at least one person.")
            else:
                zip_buffer = io.BytesIO()
                progress_bar = st.progress(0)
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    # Iterate through selected rows with an index counter (0, 1, 2...)
                    for i, (index, row) in enumerate(selected_rows.iterrows()):
                        
                        # GENERATE DYNAMIC ID
                        current_contract_id = increment_id(start_contract_id, i)
                        
                        # Prepare data
                        p_prefix = str(row[col_prefix])
                        p_name = str(row[col_name])
                        p_surname = str(row[col_surname])
                        p_id = str(row[col_id])
                        p_address = str(row[col_address])
                        p_sign_name = f"{p_prefix} {p_name} {p_surname}"
                        
                        # Context
                        context = {
                            'contract_id': current_contract_id, # Auto-running ID
                            'prefix': p_prefix,
                            'name': p_name,
                            'surname': p_surname,
                            'id_card': p_id,
                            'address': p_address,
                            'sign_name': p_sign_name
                        }
                        
                        # Render doc
                        docx_file.seek(0) 
                        doc = DocxTemplate(docx_file)
                        doc.render(context)
                        
                        # Save to stream
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # File name now includes the running ID for easy sorting
                        filename = f"{current_contract_id}_{p_name}.docx"
                        # Sanitize filename (remove slashes that might break zip)
                        filename = filename.replace("/", "-").replace("\\", "-")
                        
                        zf.writestr(filename, doc_io.getvalue())
                        
                        progress_bar.progress((i + 1) / len(selected_rows))
                
                zip_buffer.seek(0)
                st.success(f"Generated {len(selected_rows)} contracts starting from ID: {start_contract_id}")
                
                st.download_button(
                    label="⬇️ Download All as ZIP",
                    data=zip_buffer,
                    file_name="Contracts_Running_ID.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")

else:
    st.info("Please upload .docx template and .csv data.")
