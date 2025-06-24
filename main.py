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
            with st.expander("📋 View Assumptions & Parameters", expanded=False):
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
            with st.expander("📋 View Modified Assumptions & Parameters", expanded=False):
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
        st.header("Summary Comparison Analysis")
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
            csv = result.to_csv()
            st.download_button(
                label="📥 Download Comparison as CSV",
                data=csv,
                file_name="bca_comparison.csv",
                mime="text/csv"
            )
            
        except Exception as e:
            st.error(f"❌ Error calculating differences: {str(e)}")
            st.info("💡 Make sure both files have the same structure and data types.")
    
    else:
        st.header("Getting Started")
        
        st.info("""
        To view the summary comparison:
        
        1. **Upload Original File** - Go to the "Original File" tab and upload your base BCA file
        2. **Upload Modified File** - Go to the "Modified File" tab and upload your modified BCA file  
        3. **View Comparison** - Return here to see the detailed comparison analysis
        
        Both files must be processed successfully before comparison can be performed.
        """)
        
        original_status = "✅ Uploaded" if st.session_state.input_summary_df is not None else "❌ Not uploaded"
        modified_status = "✅ Uploaded" if st.session_state.modified_summary_df is not None else "❌ Not uploaded"
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**Original File Status:** {original_status}")
        with col2:
            st.info(f"**Modified File Status:** {modified_status}")