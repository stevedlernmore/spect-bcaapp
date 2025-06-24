import streamlit as st
import pandas as pd
import AMZ
import PA
from io import BytesIO

def format_dataframe_for_display(df):
    display_df = df.copy()
    for col in display_df.columns:
        for idx in display_df.index:
            value = display_df.loc[idx, col]
            if pd.isna(value):
                continue
            try:
                num_value = float(value)
                if 'Metric' in display_df.columns:
                    row_name = str(display_df.loc[idx, 'Metric']) if 'Metric' in display_df.columns else str(idx)
                else:
                    row_name = str(idx)
                if row_name.endswith('%'):
                    percentage_value = num_value * 100
                    if percentage_value < 0:
                        display_df.loc[idx, col] = f"({abs(percentage_value):,.1f}%)"
                    else:
                        display_df.loc[idx, col] = f"{percentage_value:,.1f}%"
                elif 'Cumulative' in col:
                    if num_value < 0:
                        display_df.loc[idx, col] = f"({abs(num_value):,.0f})"
                    else:
                        display_df.loc[idx, col] = f"{num_value:,.0f}"
                elif 'Per Unit' in col:
                    if num_value < 0:
                        display_df.loc[idx, col] = f"({abs(num_value):,.2f})"
                    else:
                        display_df.loc[idx, col] = f"{num_value:,.2f}"
                else:
                    if num_value < 0:
                        display_df.loc[idx, col] = f"({abs(num_value):,.2f})"
                    else:
                        display_df.loc[idx, col] = f"{num_value:,.2f}"
            except (ValueError, TypeError):
                continue
    return display_df

def to_excel_bytes(df):
  output = BytesIO()
  with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Summary')
  output.seek(0)
  return output.getvalue()

st.set_page_config(
    page_title="Rapid BCA Analyzer",
    page_icon="logo-spectra-premium.jpg", 
    layout="wide"
)

st.markdown("""
<style>
    /* Hide main menu */
    #MainMenu {visibility: hidden !important;}
    
    /* Hide footer */
    footer {visibility: hidden !important;}
    
    /* Hide header */
    header {visibility: hidden !important;}
    
    /* Remove dead space from hidden top bar */
    .stApp > div:first-child {
        margin-top: -75px;
    }
    
    /* Hide deploy button, GitHub fork, and other toolbar items */
    .stDeployButton {display: none !important;}
    .stToolbar {display: none !important;}
    div[data-testid="stToolbar"] {display: none !important;}
    
    /* Hide decoration */
    .stDecoration {display: none !important;}
    div[data-testid="stDecoration"] {display: none !important;}
    
    /* Hide status widget */
    div[data-testid="stStatusWidget"] {display: none !important;}

    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Roboto', sans-serif;
    }

    .stApp {
        background-color: #FFFFFF;
        color: #002855;
    }

    h1, h2, h3, h4, h5, h6 {
        color: #002855;
    }

    /* Header container styling */
    .header-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 20px 0;
        margin-bottom: 20px;
    }

    .title-section {
        flex: 1;
    }

    .logo-section {
        flex: 0 0 auto;
        text-align: right;
    }

    /* Button styling */
    .css-18e3th9, .stButton>button {
        background-color: #002855;
        color: #FFFFFF;
        border-radius: 5px;
        border: none;
    }

    .stButton>button:hover {
        background-color: #004080;
        color: #FFFFFF;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 4px;
        margin-bottom: 20px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 8px;
        color: #002855;
        font-weight: 500;
        font-size: 16px;
        padding: 10px 20px;
        border: none;
        transition: all 0.2s ease;
    }

    .stTabs [aria-selected="true"] {
        background-color: #002855 !important;
        color: #FFFFFF !important;
        font-weight: 600;
        box-shadow: 0 2px 4px rgba(0,40,85,0.2);
    }

    .stTabs [data-baseweb="tab"]:hover {
        background-color: #e8ecf0;
        color: #002855;
        transform: translateY(-1px);
    }

    .stTabs [aria-selected="true"]:hover {
        background-color: #004080 !important;
        color: #FFFFFF !important;
    }

    /* Enhanced File uploader styling */
    .stFileUploader {
        background-color: #FFFFFF;
    }

    .stFileUploader > div {
        background-color: #f8f9fa !important;
        border: 2px dashed #002855 !important;
        border-radius: 10px !important;
        padding: 30px 20px !important;
        text-align: center;
        transition: all 0.3s ease;
    }

    .stFileUploader > div:hover {
        border-color: #004080 !important;
        background-color: #e8f4fd !important;
        box-shadow: 0 4px 8px rgba(0,40,85,0.1);
    }

    /* File uploader text */
    .stFileUploader label {
        color: #002855 !important;
        font-weight: 600 !important;
        font-size: 16px !important;
    }

    .stFileUploader small {
        color: #6c757d !important;
        font-size: 14px !important;
    }

    /* Upload area */
    .stFileUploader > div > div {
        background-color: transparent !important;
        border: none !important;
    }

    .stFileUploader [data-testid="stFileUploaderDropzone"] {
        background-color: transparent !important;
    }

    .stFileUploader [data-testid="stFileUploaderDropzoneInstructions"] {
        color: #002855 !important;
        font-weight: 500 !important;
    }

    /* Browse files button - ONLY affects the main upload button */
    .stFileUploader button[kind="primary"] {
        background-color: #002855 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
        font-size: 14px !important;
        transition: all 0.2s ease;
    }

    .stFileUploader button[kind="primary"]:hover {
        background-color: #004080 !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(0,40,85,0.2);
    }

    .stFileUploader button[kind="primary"]:active {
        transform: translateY(0);
    }

    /* Fallback for browse button if kind attribute is not present */
    .stFileUploader button:not([kind="secondary"]) {
        background-color: #002855 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
        font-size: 14px !important;
        transition: all 0.2s ease;
    }

    .stFileUploader button:not([kind="secondary"]):hover {
        background-color: #004080 !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(0,40,85,0.2);
    }

    /* X button (delete/remove) - TRANSPARENT BACKGROUND */
    .stFileUploader button[kind="secondary"] {
        background-color: transparent !important;
        color: #002855 !important;
        border: none !important;
        border-radius: 4px !important;
        padding: 4px !important;
    }

    .stFileUploader button[kind="secondary"]:hover {
        background-color: rgba(0, 40, 85, 0.1) !important;
        color: #002855 !important;
    }

    /* X icon styling - dark blue and visible */
    .stFileUploader button[kind="secondary"] svg {
        fill: #002855 !important;
        width: 16px !important;
        height: 16px !important;
    }

    .stFileUploader button[kind="secondary"] svg * {
        fill: #002855 !important;
    }

    /* Additional selectors for the X button */
    .stFileUploader button[title*="Remove"] svg *,
    .stFileUploader button[title*="Delete"] svg *,
    .stFileUploader button[title*="Close"] svg * {
        fill: #002855 !important;
    }

    /* Target the specific delete button */
    .stFileUploader [data-testid="fileDeleteBtn"] {
        background-color: transparent !important;
        color: #002855 !important;
        border: none !important;
    }

    .stFileUploader [data-testid="fileDeleteBtn"]:hover {
        background-color: rgba(0, 40, 85, 0.1) !important;
        color: #002855 !important;
    }

    .stFileUploader [data-testid="fileDeleteBtn"] svg,
    .stFileUploader [data-testid="fileDeleteBtn"] svg * {
        fill: #002855 !important;
    }

    /* File upload icon styling - for browse button only */
    .stFileUploader button:not([kind="secondary"]) svg {
        fill: #FFFFFF !important;
    }

    /* Expander styling */
    .streamlit-expanderHeader {
        background-color: #f8f9fa;
        border-radius: 5px;
        color: #002855;
        font-weight: 500;
    }

    /* Success/error message styling */
    .stSuccess {
        background-color: #d4edda;
        border-color: #c3e6cb;
        color: #155724;
    }

    .stError {
        background-color: #f8d7da;
        border-color: #f5c6cb;
        color: #721c24;
    }

    .stWarning {
        background-color: #fff3cd;
        border-color: #ffeaa7;
        color: #856404;
    }

    .stInfo {
        background-color: #d1ecf1;
        border-color: #bee5eb;
        color: #0c5460;
    }

    /* Hide dataframe download button */
    .stDataFrame [data-testid="stElementToolbar"] {
        display: none !important;
    }
    
    .stDataFrame [data-testid="stElementToolbarButton"] {
        display: none !important;
    }
    
    /* Alternative selectors for download buttons */
    .stDataFrame button[title="Download"] {
        display: none !important;
    }
    
    .stDataFrame [aria-label="Download"] {
        display: none !important;
    }

    /* Dataframe styling with alternating colors */
    .stDataFrame {
        border-radius: 8px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }

    /* Header row styling */
    .stDataFrame thead tr th {
        background-color: #002855 !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border: none !important;
        padding: 12px 8px !important;
        text-align: center !important;
    }

    /* Alternating row colors */
    .stDataFrame tbody tr:nth-child(odd) {
        background-color: #f8f9fa !important;
    }

    .stDataFrame tbody tr:nth-child(even) {
        background-color: #ffffff !important;
    }

    /* Row hover effect */
    .stDataFrame tbody tr:hover {
        background-color: #e3f2fd !important;
        transition: background-color 0.2s ease !important;
    }

    /* Cell styling */
    .stDataFrame tbody tr td {
        border: none !important;
        padding: 10px 8px !important;
        vertical-align: middle !important;
        color: #002855 !important;
    }

    /* First column (usually contains labels) styling */
    .stDataFrame tbody tr td:first-child {
        font-weight: 500 !important;
        background-color: #f0f2f6 !important;
    }

    /* Override first column hover */
    .stDataFrame tbody tr:hover td:first-child {
        background-color: #d1e7dd !important;
    }

    /* Grid lines */
    .stDataFrame table {
        border-collapse: collapse !important;
    }

    .stDataFrame td, .stDataFrame th {
        border-right: 1px solid #e0e0e0 !important;
    }

    .stDataFrame td:last-child, .stDataFrame th:last-child {
        border-right: none !important;
    }

</style>
""", unsafe_allow_html=True)

# Initialize session state
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

# Header section using container
header_container = st.container()
with header_container:
  col1, col2 = st.columns([3, 1])
  with col1:
    st.markdown("""
    <div class="title-section">
      <h1 style="margin-bottom: 5px;">Rapid BCA Analyzer</h1>
      <h3 style="color: #6c757d; font-weight: 400; margin-top: 0;">Instantly Analyze Business Cases</h3>
    </div>
    """, unsafe_allow_html=True)

with col2:
  st.markdown("""
  <div class="logo-section">
  """, unsafe_allow_html=True)
  st.image("logo-spectra-premium.jpg", width=300)
  st.markdown("</div>", unsafe_allow_html=True)

def process_file(uploaded_file):
  filename = uploaded_file.name.upper()
  if filename.startswith('AMZ'):
      return AMZ.getSummary(uploaded_file)
  elif filename.startswith('PA'):
      return PA.getSummary(uploaded_file)
  else:
      raise ValueError("File name must start with 'AMZ' or 'PA' to determine the correct processing method.")

# Main content container
main_container = st.container(border=True)
with main_container:
  input_tab, modified_tab, summary_tab = st.tabs(["Original File", "Modified File", "Summary Comparison"])

  with input_tab:
    st.header("Upload Original BCA File")
    st.write("Upload your original Business Case Analysis file. The system will automatically detect whether it's an AMZ or PA file based on the filename.")
    
    input_file = st.file_uploader(
      "Choose your original Excel file", 
      type=['xlsx', 'xls']
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
            format_dataframe_for_display(st.session_state.input_summary_df),
            use_container_width=True,
            hide_index=True,
          )

        file_name = f'{input_file.name.split(".")[0]}_export.xlsx'
        st.download_button(
          label="Download Summary",
          data=to_excel_bytes(st.session_state.input_summary_df),
          file_name=file_name,
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          key="download_input_summary"
        )

  with modified_tab:
    st.header("Upload Modified BCA File")
    st.write("Upload a modified version of your BCA file to compare changes and analyze the impact of modifications.")
    
    modified_file = st.file_uploader(
      "Choose your modified Excel file", 
      type=['xlsx', 'xls'],
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
            format_dataframe_for_display(st.session_state.modified_summary_df),
            use_container_width=True,
            hide_index=True
          )
        file_name = f'{modified_file.name.split(".")[0]}_export.xlsx'
        st.download_button(
          label="Download Summary",
          data=to_excel_bytes(st.session_state.modified_summary_df),
          file_name=file_name,
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          key="download_modified_summary"
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
              format_dataframe_for_display(result),
              use_container_width=True,
              height=600
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
