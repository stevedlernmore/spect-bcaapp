import streamlit as st
import pandas as pd
import AMZ
import ORL
import AZ
import PA
import UNITED
import VAST
import standard
from io import BytesIO
from openai import OpenAI

AI_client = OpenAI(
  api_key=st.secrets["OPEN_AI_KEY"]
)

def escape_markdown(text):
    escape_chars = ['$']
    for char in escape_chars:
        text = text.replace(char, f"\\{char}")
    return text

def format_dataframe_for_display(df):
  display_df = df.copy()
  for col in display_df.columns:
    for idx in display_df.index:
      value = display_df.loc[idx, col]
      if pd.isna(value) or value == '' or str(value).strip() == '':
        display_df.loc[idx, col] = ""  # Set to empty string instead of NaN
        continue
      try:
        num_value = float(value)
        if 'Metric' in display_df.columns:
          row_name = str(display_df.loc[idx, 'Metric']) if 'Metric' in display_df.columns else str(idx)
        else:
          row_name = str(idx)
        if row_name.endswith('%') and 'FACTORING' not in row_name:
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
        if pd.isna(value) or value == '' or str(value).strip() == '':
          display_df.loc[idx, col] = ""
        else:
          display_df.loc[idx, col] = str(value)
  return display_df

def to_excel_bytes(summary_df, defaults_df=None):
  output = BytesIO()
  with pd.ExcelWriter(output, engine='openpyxl') as writer:
    summary_df.to_excel(writer, index=False, sheet_name='Summary')
    if defaults_df is not None:
      defaults_df.to_excel(writer, sheet_name='Defaults & Assumptions')
  
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

    /* Hide profile container/account menu */
    .stProfileContainer {display: none !important;}
    div[data-testid="stProfileContainer"] {display: none !important;}
    
    /* Hide account menu button */
    button[title="Account"] {display: none !important;}
    button[aria-label="Account"] {display: none !important;}
    
    /* Hide user avatar/profile picture */
    .stUserAvatar {display: none !important;}
    div[data-testid="stUserAvatar"] {display: none !important;}
    
    /* Hide account menu dropdown */
    .stAccountMenu {display: none !important;}
    div[data-testid="stAccountMenu"] {display: none !important;}

    /* Additional selectors for profile-related elements */
    [data-testid="stHeader"] button[kind="header"] {display: none !important;}
    [data-testid="stHeader"] .stButton {display: none !important;}

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

def check_password():
  def password_entered():
    if (st.session_state["username"] == "BCA_User" and st.session_state["password"] == "Analyzer"):
      st.session_state["authenticated"] = True
    else:
      st.error("Incorrect username or password")
      st.session_state["authenticated"] = False
      
  col1, col2, col3 = st.columns([2, 1, 2])
  with col2:
      if "authenticated" not in st.session_state or st.session_state["authenticated"] is False:
        container = st.container(border=True)
        with container:
          st.image("logo-spectra-premium.jpg", width=300)
          st.session_state["username"] = st.text_input("Username", key="username_input")
          st.session_state["password"] = st.text_input("Password", type="password", key="password_input")
          st.button("Login", on_click=password_entered, key="login_button", use_container_width=True)
          return False
      elif not st.session_state["authenticated"]:
        return False
      else:
        return True

def getFileType(file_name):
  if file_name.startswith('AMZ'):
    return "Amazon"
  elif file_name.startswith('PA'):
    return "Parts Authority"
  elif file_name.startswith('ORL'):
    return "OReilly"
  elif file_name.startswith('AZ'):
    return "AZ"
  elif file_name.startswith('UNITED'):
    return "United"
  elif file_name.startswith('VAST'):
    return "VAST"
  else:
    return "Standard BCA"

if check_password():
  if 'input_file' not in st.session_state:
      st.session_state.input_file = None
  if 'input_summary_df' not in st.session_state:
      st.session_state.input_summary_df = None
  if 'input_defaults_df' not in st.session_state:
      st.session_state.input_defaults_df = None
  if 'input_assumptions_df' not in st.session_state:
      st.session_state.input_assumptions_df = None
  if 'user_assumptions_df' not in st.session_state:
      st.session_state.user_assumptions_df = None
  if 'user_defaults_df' not in st.session_state:
      st.session_state.user_defaults_df = None
  if 'user_summary_df' not in st.session_state:
      st.session_state.user_summary_df = None

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

  def process_file(uploaded_file, user_defaults_df=None, volume=0.0):
    filename = uploaded_file.name.upper()
    if filename.startswith('AMZ'):
      return AMZ.getSummary(uploaded_file, user_defaults_df, volume)
    elif filename.startswith('PA'):
      return PA.getSummary(uploaded_file, user_defaults_df, volume)
    elif filename.startswith('ORL'):
      return ORL.getSummary(uploaded_file, user_defaults_df, volume)
    elif filename.startswith('AZ'):
      return AZ.getSummary(uploaded_file, user_defaults_df, volume)
    elif filename.startswith('UNITED'):
      return UNITED.getSummary(uploaded_file, user_defaults_df, volume)
    elif filename.startswith('VAST'):
      return VAST.getSummary(uploaded_file, user_defaults_df, volume)
    else:
      return standard.getSummary(uploaded_file, user_defaults_df, volume)

  main_container = st.container(border=True)
  with main_container:
    input_tab, summary_tab = st.tabs(["Original File", "Summary Comparison"])

    with input_tab:
      st.header("Upload Original BCA File")
      st.write("Upload your original Business Case Analysis file. The system will automatically detect based on the filename.")
      input_file = st.file_uploader(
        "Choose your original Excel file", 
        type=['xlsx', 'xls']
      )
      if input_file is not None:
        file_type = getFileType(input_file.name)
        if st.session_state.input_file != input_file:
          st.toast(f"Starting {file_type} Calculations...")
          st.session_state.input_file = input_file
          try:
            with st.spinner('Processing file... This may take a moment.'):
              output, assumption_test, defaults_test = process_file(input_file)
              st.session_state.input_summary_df = output
              st.session_state.input_defaults_df = defaults_test
              st.session_state.input_assumptions_df = assumption_test
              if defaults_test is not None:
                st.session_state.user_defaults_df = defaults_test.copy()
              if assumption_test is not None:
                st.session_state.user_assumptions_df = assumption_test.copy()
              st.session_state.user_summary_df = None
          except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
        def formatter(name, value, defaults=False):
          print(name, value)
          if 'per unit' in name.lower() or 'cumulative' in name.lower() or 'big' in name.lower() or 'shipping' in name.lower() or 'return allowance' in name.lower() or 'pallets' in name.lower() or 'delivery' in name.lower() or 'inspect return' in name.lower() or 'rebox' in name.lower() or 'labor' in name.lower() or 'overhead' in name.lower() or 'special marketing' in name.lower():
            if defaults:
              value = f"% {value*100:.3f}"
            elif 'alliance' in name.lower():
              value = f"% {value*100:.3f}"
            else:
              value = f"$ {value:.2f}"
          elif 'qty' in name.lower():
            value = f"{value}"
          else:
            value = f"% {value*100:.3f}"
          return f'**{value}**'
        if file_type == "VAST":
          col1, col2 = st.columns(2)
          with col1:
            st.markdown("#### Current Values")
            with st.expander("See Current", expanded=True):
              for x in st.session_state.input_assumptions_df.get('Current'):
                st.write(f"{x}:", formatter(x,st.session_state.input_assumptions_df.get('Current')[x]))
          with col2:
            st.markdown("#### New Values")
            with st.expander("See New", expanded=True):
              for x in st.session_state.input_assumptions_df.get('New'):
                st.write(f"{x}:", formatter(x,st.session_state.input_assumptions_df.get('New')[x]))
          st.markdown("#### Assumptions")
          with st.expander("See Assumptions", expanded=True):
            for x in st.session_state.input_defaults_df:
              st.write(f"{x}:", formatter(x, st.session_state.input_defaults_df[x], True))
        else:
          if st.session_state.input_defaults_df is not None or st.session_state.input_assumptions_df is not None:
            qty_col1, qty_col2, qty_col3 = st.columns([1,2,1])
            with qty_col2:
              st.write("#### Quantity Volume Control")
              with st.container(border=True):
                column1, column2 = st.columns(2)
                with column1:
                  st.write('Volume Percentage Change:', formatter('Volume Percentage Change', 0.0))
                with column2:
                  volume = st.number_input(label="Volume Percentage Change", label_visibility="collapsed", value=0.0, key="volumne", help="Adjust the volume percentage to see how it affects the analysis.")
            column1, column2 = st.columns(2)
            with column1:
              st.markdown("#### Assumptions")
              with st.expander("See Assumptions", expanded=True):
                if st.session_state.input_assumptions_df is None:
                  st.write("No assumptions found in the original file.")
                else:
                  for x in st.session_state.input_assumptions_df:
                    subcol1, subcol2 = st.columns([3, 1])
                    with subcol1:
                      st.write(f"{x}:\n" + formatter(x, st.session_state.input_assumptions_df[x]))
                    with subcol2:
                      st.session_state.user_assumptions_df[x] = st.number_input(label=f"assumptions", value=st.session_state.input_assumptions_df[x], key=f"assumption_{x}", label_visibility="collapsed", format="%.4f")
            with column2:
              st.markdown("#### Defaults")
              with st.expander("See Defaults", expanded=True):
                if st.session_state.input_defaults_df is None:
                  st.write("No defaults found in the original file.")
                else:
                  for x in st.session_state.input_defaults_df:
                    subcol1, subcol2 = st.columns([3, 1])
                    with subcol1:
                      st.write(f"{x}:\n" + formatter(x, st.session_state.input_defaults_df[x], False))
                    with subcol2:
                      st.session_state.user_defaults_df[x] = st.number_input(label=f"defaults", value=st.session_state.input_defaults_df[x], key=f"defaults_{x}", label_visibility="collapsed", format="%.4f")  
            apply_changes = st.button("Calculate with Modified Values", key="apply_changes_button", use_container_width=True)
            if apply_changes:
                st.toast("Calculating with modified values... This may take a moment.")
                output, assumption_test, defaults_test = process_file(input_file, (st.session_state.user_defaults_df or {}) | (st.session_state.user_assumptions_df or {}), volume)
                st.toast("Calculation complete!")
                st.session_state.user_summary_df = output
                st.session_state.response = "Analysis in progress..."

        if st.session_state.input_summary_df is not None:
          with st.expander("📊 View Complete Summary", expanded=False):
            st.dataframe(
              format_dataframe_for_display(st.session_state.input_summary_df),
              use_container_width=True,
              hide_index=True,
            )
          file_name = f'{input_file.name.split(".")[0]}_export.xlsx'
          df = pd.DataFrame({
            "Values":(st.session_state.user_assumptions_df or {}) | (st.session_state.user_defaults_df or {})
          })
          st.download_button(
            label="Download Summary",
            data=to_excel_bytes(st.session_state.input_summary_df, df),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_input_summary"
          )

    with summary_tab:
      input_df = st.session_state.input_summary_df
      summary_df = st.session_state.user_summary_df
      if input_df is not None and summary_df is not None:
        input_defaults_df = pd.DataFrame({
          "Values": (st.session_state.input_defaults_df or {}) | (st.session_state.input_assumptions_df or {})
        })
        defaults_df = pd.DataFrame({
          "Values": (st.session_state.user_defaults_df or {}) | (st.session_state.user_assumptions_df or {})
        })
        st.header("📊 Summary Comparison Analysis")
        
        should_generate_analysis = False
        if 'response' not in st.session_state:
          st.session_state.response = "Click 'Calculate with Modified Values' to generate analysis..."
          should_generate_analysis = False
        elif st.session_state.response == "Analysis in progress...":
          should_generate_analysis = True
        
        if st.session_state.response == "Click 'Calculate with Modified Values' to generate analysis...":
          st.info(st.session_state.response)
        else:
          st.success(escape_markdown(st.session_state.response))
        if input_defaults_df is not None and defaults_df is not None:
          st.subheader("📋 Changes in Defaults & Assumptions")
          input_defaults = input_defaults_df
          modified_defaults = defaults_df
          common_params = input_defaults.index.intersection(modified_defaults.index)
          changes = []
          for param in common_params:
            original_value = input_defaults.loc[param].loc['Values']
            modified_value = modified_defaults.loc[param].loc['Values']
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
            if modified_value - original_value < 0:
              description = f"{original_value} → {modified_value} ({abs(difference):.3f} decrease)"
            else:
              description = f"{original_value} → {modified_value} ({abs(difference):.3f} increase)"
            changes.append({
                'Parameter': param,
                'Original Value': original_value,
                'Modified Value': modified_value,
                'Change Description': description
            })
          if volume != 0.0:
            if 0 - volume < 0:
              description = f"{volume}% increase"
            else:
              description = f"{volume}% decrease"
            changes.append({
              'Parameter': f"QTY Volume",
              'Original Value': 0,
              'Modified Value': volume,
              'Change Description': description
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
        st.write("Below is the difference between your modified (user-edited) file and original file (Modified - Original):")
        try:
            numeric_result = summary_df.iloc[:, 1:] - input_df.iloc[:, 1:]
            result = pd.DataFrame(
                numeric_result.values,
                index=input_df.iloc[:,0],
                columns=input_df.columns[1:,]
            )
            result_with_metrics = result.reset_index()
            result_with_metrics.rename(columns={'index': 'Metric'}, inplace=True)
            st.dataframe(
              format_dataframe_for_display(result),
              use_container_width=True,
              height=600
            )
            
            if should_generate_analysis:
                with st.spinner("Generating AI analysis..."):
                  response = AI_client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                      {
                        "role": "system",
                        "content": "You are an expert business case analyst specializing in financial impact assessment and strategic decision-making. Analyze parameter changes in Business Case Analysis (BCA) and provide structured, actionable insights."
                      },
                      {
                        "role": "user",
                        "content": f'''Analyze the following Business Case Analysis data and provide a structured response in the exact format specified:

            {{
              "analysis_context": {{
                "scenario": "Parameter optimization analysis comparing Modified BCA vs Original BCA",
                "data_structure": "Multiple product lines with Cumulative and Per Unit metrics",
                "calculation_method": "Difference analysis (Modified - Original values)"
              }},
              "input_data": {{
                "parameter_changes": {changes_df.to_csv(index=False) if file_type != "Amazon" else changes_df.to_string(index=False)},
                "financial_impact_differences": {result_with_metrics.to_csv(index=False) if file_type != "Amazon" else result_with_metrics.to_string(index=False)}
              }},
              "analysis_requirements": {{
                "executive_summary": {{
                  "word_limit": 100,
                  "focus_areas": ["Overall financial impact", "Contribution margin % trend"],
                  "format_requirements": ["Bold critical statements using **text**", "Clear positive/negative assessment"]
                }},
                "detailed_impact_analysis": {{
                  "bullet_count": "5-7 points",
                  "required_elements": [
                    "Parameter-specific business impact",
                    "Quantified dollar amounts and percentages", 
                    "Product line differentiation",
                    "Cumulative vs per-unit distinction",
                    "Risk/opportunity identification",
                    "Key financial metrics analysis",
                    "Actionable business insights"
                  ],
                  "format": "**Impact Title**: Detailed description with specific numbers (1-2 sentences)"
                }}
              }},
              "output_format": {{
                "structure": [
                  "### Executive Summary",
                  "### Detailed Impact Analysis"
                ],
                "guidelines": [
                  "Use specific dollar amounts and percentages from provided data",
                  "Distinguish between different product lines when applicable",
                  "Focus on business implications over numerical changes",
                  "Highlight critical metrics for decision-making",
                  "Use appropriate positive/negative language for financial impacts"
                ]
              }}
            }}

            RESPOND WITH:
            ## Executive Summary
            [100-word assessment focusing on contribution margin % as primary indicator. Bold key statements with **text**]

            ## Detailed Impact Analysis
            [5-7 bullet points in format: **Impact Title**: Description with specific dollar amounts and percentages]'''
                      }
                    ],
                    max_tokens=3500,
                    temperature=0.2
                  )
                  st.session_state.response = response.choices[0].message.content
                  st.rerun() 
            file_name = f'{st.session_state.input_file.name.split(".")[0]}_comparison.xlsx'
            st.download_button(
              label="Download Summary",
              data=to_excel_bytes(result_with_metrics, defaults_df),
              file_name=file_name,
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              key="download_comparison"
            )
        except Exception as e:
          st.error(f"❌ Error calculating differences: {str(e)}")
          st.info("💡 Make sure both files have the same structure and data types.")
      else:
        st.header("Getting Started")
        st.info("""
        To view the summary comparison:
        
        1. **Upload Original File** - Go to the "Original File" tab and upload your base BCA file
        2. **Edit Defaults & Assumptions** - Edit values and click 'Apply Changes' to see the impact
        3. **View Comparison** - Return here to see the detailed comparison analysis
        
        Both files must be processed successfully before comparison can be performed.
        """)
        original_status = "✅ Uploaded" if st.session_state.input_summary_df is not None else "❌ Not uploaded"
        modified_status = "✅ Calculated" if st.session_state.get('user_edit_summary_df') is not None else "❌ Not calculated"
        col1, col2 = st.columns(2)
        with col1:
          st.info(f"**Original File Status:** {original_status}")
        with col2:
          st.info(f"**Modified (User-Edited) Status:** {modified_status}")
