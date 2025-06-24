import streamlit as st
import pandas as pd
import AMZ
import PA

st.set_page_config(
    page_title="BCA Summary Viewer",
    layout="wide"
)

st.title("BCA Summary Viewer")
st.subheader("Business Case Analysis Tool for AMZ and PA Files")

if 'input_file' not in st.session_state:
    st.session_state.input_file = None
if 'modified_file' not in st.session_state:
    st.session_state.modified_file = None
if 'input_summary_df' not in st.session_state:
    st.session_state.input_summary_df = None
if 'modified_summary_df' not in st.session_state:
    st.session_state.modified_summary_df = None
if 'input_defaults_df' not in st.session_state:
    st.session_state.input_defaults_df = None
if 'modified_defaults_df' not in st.session_state:
    st.session_state.modified_defaults_df = None

def process_file(uploaded_file):
    filename = uploaded_file.name.upper()
    if filename.startswith('AMZ'):
        return AMZ.getSummary(uploaded_file)
    elif filename.startswith('PA'):
        return PA.getSummary(uploaded_file)
    else:
        raise ValueError("File name must start with 'AMZ' or 'PA' to determine the correct processing method.")
    
# Main content tabs
input_tab, modified_tab, summary_tab = st.tabs(["Original File", "Modified File", "Summary Comparison"])

with input_tab:
    st.header("Upload Original BCA File")
    st.write("Upload your original Business Case Analysis file. The system will automatically detect whether it's an AMZ or PA file based on the filename.")
    
    input_file = st.file_uploader(
        "Choose your original Excel file", 
        type=['xlsx', 'xls'],
        help="File name must start with 'AMZ' or 'PA'"
    )
    
    if input_file is not None:
        file_type = "Amazon" if input_file.name.upper().startswith('AMZ') else "Parts Authority"
        if st.session_state.input_file != input_file:
            st.toast(f"Starting {file_type} Calculations...")
            st.session_state.input_file = input_file
            try:
                with st.spinner('Processing file... This may take a moment.'):
                    output, assumptions = process_file(input_file)
                    st.session_state.input_summary_df = output
                    st.session_state.input_defaults_df = assumptions
                st.toast("✅ File processed successfully!")
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")

        if st.session_state.input_defaults_df is not None:
            with st.expander("📋 View Defaults & Assumptions", expanded=False):
                st.dataframe(
                    st.session_state.input_defaults_df,
                    use_container_width=True,
                    hide_index=True
                )
        
        if st.session_state.input_summary_df is not None:
            with st.expander("📊 View Complete Summary", expanded=True):
                st.dataframe(
                    st.session_state.input_summary_df,
                    use_container_width=True,
                    hide_index=True
                )

with modified_tab:
    st.header("Upload Modified BCA File")
    st.write("Upload a modified version of your BCA file to compare changes and analyze the impact of modifications.")
    
    modified_file = st.file_uploader(
        "Choose your modified Excel file", 
        type=['xlsx', 'xls'],
        help="File name must start with 'AMZ' or 'PA' and match the original file type",
        key="modified_uploader"
    )
    
    if modified_file is not None:
        file_type = "Amazon" if modified_file.name.upper().startswith('AMZ') else "Parts Authority"
        if st.session_state.modified_file != modified_file:
            st.toast(f"Starting {file_type} Calculations...")
            st.session_state.modified_file = modified_file
            try:
                with st.spinner('Processing modified file... This may take a moment.'):
                    output, assumptions = process_file(modified_file)
                    st.session_state.modified_summary_df = output
                    st.session_state.modified_defaults_df = assumptions
                st.toast("✅ Modified file processed successfully!")
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")

        if st.session_state.modified_defaults_df is not None:
            with st.expander("📋 View Modified Defaults & Assumptions", expanded=False):
                st.dataframe(
                    st.session_state.modified_defaults_df,
                    use_container_width=True,
                    hide_index=True
                )
        
        if st.session_state.modified_summary_df is not None:
            with st.expander("📊 View Modified Summary", expanded=True):
                st.dataframe(
                    st.session_state.modified_summary_df,
                    use_container_width=True,
                    hide_index=True
                )

with summary_tab:
    if st.session_state.input_summary_df is not None and st.session_state.modified_summary_df is not None:
        input_file_type = "AMZ" if st.session_state.input_file.name.upper().startswith('AMZ') else "PA"
        modified_file_type = "AMZ" if st.session_state.modified_file.name.upper().startswith('AMZ') else "PA"
        
        if input_file_type != modified_file_type:
            st.error(f"❌ File Type Mismatch!")
            st.warning(f"""
            **Cannot compare files of different types:**
            
            - **Original File:** {input_file_type} ({st.session_state.input_file.name})
            - **Modified File:** {modified_file_type} ({st.session_state.modified_file.name})
            
            Please upload files of the same type to perform comparison.
            """)
        else:
            st.header("📊 Summary Comparison Analysis")
            st.success(f"✅ Comparing {input_file_type} files: {st.session_state.input_file.name} vs {st.session_state.modified_file.name}")
            if st.session_state.input_defaults_df is not None and st.session_state.modified_defaults_df is not None:
                st.subheader("📋 Changes in Defaults & Assumptions")
              
                input_defaults = st.session_state.input_defaults_df.set_index('Parameter')
                modified_defaults = st.session_state.modified_defaults_df.set_index('Parameter')
                common_params = input_defaults.index.intersection(modified_defaults.index)
                
                changes = []
                for param in common_params:
                    original_value = input_defaults.loc[param, 'Value']
                    modified_value = modified_defaults.loc[param, 'Value']
                    try:
                        if pd.isna(original_value) and pd.isna(modified_value):
                            continue
                        elif pd.isna(original_value) or pd.isna(modified_value):
                            difference = "N/A"
                        elif abs(float(original_value) - float(modified_value)) > 1e-10:
                            difference = float(modified_value) - float(original_value)
                        else:
                            continue 
                    except (ValueError, TypeError):
                        if str(original_value) != str(modified_value):
                            difference = "Text Changed"
                        else:
                            continue
                    
                    changes.append({
                        'Parameter': param,
                        'Original Value': original_value,
                        'Modified Value': modified_value,
                        'Change Description': f"{original_value} → {modified_value}"
                    })
                
                if changes:
                    changes_df = pd.DataFrame(changes)
                    st.dataframe(
                        changes_df,
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("✅ No changes detected in defaults and assumptions.")
                
                st.divider()
            
            st.subheader("📊 Summary Data Comparison")
            st.write("Below is the difference between your modified file and original file (Modified - Original):")
          
            try:
                numeric_result = st.session_state.modified_summary_df.iloc[1:, 1:] - st.session_state.input_summary_df.iloc[1:, 1:]
                result = pd.DataFrame(
                    numeric_result.values,
                    index=st.session_state.input_summary_df.iloc[1:, 0],
                    columns=st.session_state.input_summary_df.columns[1:]
                )
                st.dataframe(
                    result,
                    use_container_width=True,
                    height=600
                )
                
            except Exception as e:
                st.error(f"❌ Error calculating differences: {str(e)}")
                st.info("💡 Make sure both files have the same structure and data types.")
    