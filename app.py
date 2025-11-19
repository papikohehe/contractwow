import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile

# --- Page Configuration ---
st.set_page_config(page_title="Thai Contract Automation", page_icon="📝", layout="wide")

st.title("📝 Thai Contract Automation Tool")
st.markdown("Upload your contract template (.docx) and your data (.csv) to generate contracts.")

# --- Sidebar: Uploads ---
st.sidebar.header("1. Upload Files")
docx_file = st.sidebar.file_uploader("Upload Contract Template (.docx)", type=["docx"])
csv_file = st.sidebar.file_uploader("Upload Data (.csv)", type=["csv"])

# --- Main Logic ---
if docx_file and csv_file:
    # 1. Read CSV
    try:
        try:
            df = pd.read_csv(csv_file, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding='tis-620')
        
        st.success("Files uploaded successfully!")
        
        # --- Global Contract Settings ---
        st.subheader("2. Global Contract Details")
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            # This variable 'contract_id' can be used in Word as {{ contract_id }}
            global_contract_id = st.text_input("Enter Contract ID / Date / Reference No.", value="CT-2025-001")
        
        # --- Column Mapping Interface ---
        st.subheader("3. Map CSV Columns")
        with st.expander("Click to adjust column mapping", expanded=False):
            columns = df.columns.tolist()
            col1, col2 = st.columns(2)
            with col1:
                # Defaults based on your description
                col_prefix = st.selectbox("Prefix (นาย/นางสาว)", columns, index=2 if len(columns) > 2 else 0)
                col_name = st.selectbox("First Name", columns, index=3 if len(columns) > 3 else 0)
                col_surname = st.selectbox("Surname", columns, index=4 if len(columns) > 4 else 0)
            with col2:
                col_id = st.selectbox("ID Card", columns, index=5 if len(columns) > 5 else 0)
                col_address = st.selectbox("Address", columns, index=9 if len(columns) > 9 else 0)

        st.markdown("---")
        
        # --- Bulk Selection Interface ---
        st.subheader("4. Select People for Contracts")
        
        # Create a clean display dataframe for selection
        display_df = df[[col_prefix, col_name, col_surname, col_id]].copy()
        display_df['Full Name'] = display_df[col_prefix].astype(str) + " " + display_df[col_name].astype(str) + " " + display_df[col_surname].astype(str)
        
        # Add a "Select" column to the dataframe for the data editor
        # We insert it at the beginning
        df_with_selections = df.copy()
        df_with_selections.insert(0, "Select", False)
        
        # Display data editor with checkboxes
        edited_df = st.data_editor(
            df_with_selections,
            column_config={
                "Select": st.column_config.CheckboxColumn(
                    "Generate?",
                    help="Select to generate contract",
                    default=False,
                )
            },
            disabled=df.columns, # Disable editing other columns
            hide_index=True,
        )
        
        # Filter to get only selected rows
        selected_rows = edited_df[edited_df.Select]
        
        # --- Generation Logic ---
        st.write(f"**Selected: {len(selected_rows)} people**")
        
        if st.button("Generate Selected Contracts", type="primary"):
            if len(selected_rows) == 0:
                st.warning("Please select at least one person from the table above.")
            else:
                # Create a ZIP buffer if multiple files, or single IO if one
                zip_buffer = io.BytesIO()
                generated_files = []
                
                progress_bar = st.progress(0)
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for idx, (index, row) in enumerate(selected_rows.iterrows()):
                        # Prepare data
                        p_prefix = str(row[col_prefix])
                        p_name = str(row[col_name])
                        p_surname = str(row[col_surname])
                        p_id = str(row[col_id])
                        p_address = str(row[col_address])
                        p_sign_name = f"{p_prefix} {p_name} {p_surname}"
                        
                        # Context for Jinja2 template
                        context = {
                            'contract_id': global_contract_id, # Added the manual input here
                            'prefix': p_prefix,
                            'name': p_name,
                            'surname': p_surname,
                            'id_card': p_id,
                            'address': p_address,
                            'sign_name': p_sign_name
                        }
                        
                        # Render
                        # We must reload the docx template fresh for every loop to avoid overwriting
                        # Reset file pointer of uploaded file
                        docx_file.seek(0) 
                        doc = DocxTemplate(docx_file)
                        doc.render(context)
                        
                        # Save individual doc to stream
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        filename = f"Contract_{global_contract_id}_{p_name}_{p_surname}.docx"
                        
                        # Add to ZIP
                        zf.writestr(filename, doc_io.getvalue())
                        generated_files.append(filename)
                        
                        # Update progress
                        progress_bar.progress((idx + 1) / len(selected_rows))
                
                zip_buffer.seek(0)
                
                st.success(f"Successfully generated {len(selected_rows)} contracts!")
                
                # Download Button
                st.download_button(
                    label=f"⬇️ Download All ({len(selected_rows)} Files)",
                    data=zip_buffer,
                    file_name=f"Contracts_Batch_{global_contract_id}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")

else:
    st.info("Please upload .docx template and .csv data.")
