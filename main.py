from datetime import datetime
import streamlit as st
import pandas as pd
import standard
from openai import OpenAI
from UtilityLibrary import ExcelExport
from BCA_Matrix_tab import BCA_Matrix_tab
import sqlite3
import uuid
import hashlib
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

auto_size = JsCode("""
function(params) {
    params.api.sizeColumnsToFit();
}
""")

@st.dialog("Register")
def register():
  st.write("Welcome!")
  username = st.text_input("Username")
  fullname = st.text_input("Full Name")
  password = st.text_input("New Password", type="password")
  confirm_password = st.text_input("Confirm Password", type="password")
  if st.button("Submit", use_container_width=True):
    
    if password != confirm_password:
      st.toast("Passwords do not match", icon="🚨")
    elif not username or not fullname or not password:
      st.toast("All fields are required!", icon="🚨")
    elif cursor.execute('''SELECT * FROM users WHERE username = ?''', (username,)).fetchone():
      st.toast("Username already exists", icon="🚨")
    else:
      user_id = str(uuid.uuid4())
      password_hash = hashlib.scrypt(password.encode(), salt=b'secret_salt', n=16384, r=8, p=1).hex()
      cursor.execute('''INSERT INTO users (id, username, fullName, password) VALUES (?, ?, ?, ?)''', (user_id, username, fullname, password_hash))
      conn.commit()
      st.toast("User registered successfully")

      st.session_state["authenticated"] = True
      st.session_state["username"] = username
      st.session_state["fullName"] = fullname
      st.session_state["role"] = "User"
      st.session_state["password"] = password
      st.session_state["id"] = user_id
      st.rerun()

conn = sqlite3.connect('database.db', check_same_thread=False)
cursor = conn.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS users (id TEXT PRIMARY KEY, username TEXT, fullName TEXT, password TEXT, role TEXT DEFAULT 'User')''')
cursor.execute('''CREATE TABLE IF NOT EXISTS files (id TEXT PRIMARY KEY, user_id TEXT, input_fileName TEXT, output_fileName TEXT, upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (user_id) REFERENCES users (id))''')

volume = 0

AI_client = OpenAI(
  api_key=st.secrets["OPEN_AI_KEY"]
)

def escape_markdown(text):
  escape_chars = ['$']
  for char in escape_chars:
      text = text.replace(char, f"\\{char}")
  return text

def format_value_for_input(name, value):
  """Format a value for display in text input based on its type"""
  if value is None:
    return ""
  if 'per unit' in name.lower() or 'cumulative' in name.lower() or 'big' in name.lower() or \
      'shipping' in name.lower() or 'pallets' in name.lower() or \
      'delivery' in name.lower() or 'inspect return' in name.lower() or 'rebox' in name.lower() or \
      'labor' in name.lower() or 'overhead' in name.lower() or 'special marketing' in name.lower():
    if 'alliance' in name.lower():
      return f"{value*100:.2f}%"
    else:
      return f"${value:.2f}"
  elif 'qty' in name.lower():
    return f"{value}"
  else:
    return f"{value*100:.2f}%"

def parse_input_value(name, input_str):
  if not input_str or input_str.strip() == "":
    return 0.0
  input_str = input_str.strip()
  try:
    if input_str.endswith('%'):
      numeric_part = input_str[:-1].strip()
      value = float(numeric_part) / 100.0
    elif input_str.startswith('$'):
      numeric_part = input_str[1:].strip()
      value = float(numeric_part)
    else:
      value = float(input_str)
    return value
  except:
    return 0.0

def auto_format_input(name, input_str):
  if not input_str or input_str.strip() == "":
    return ""
  input_str = input_str.strip()
  if input_str.endswith('%') or input_str.startswith('$'):
    return input_str
  is_percentage_field = not ('per unit' in name.lower() or 'cumulative' in name.lower() or 
                            'big' in name.lower() or 'shipping' in name.lower() or 
                            'return allowance' in name.lower() or 'pallets' in name.lower() or 
                            'delivery' in name.lower() or 'inspect return' in name.lower() or 
                            'rebox' in name.lower() or 'labor' in name.lower() or 
                            'overhead' in name.lower() or 'special marketing' in name.lower() or
                            'qty' in name.lower())
  is_dollar_field = ('per unit' in name.lower() or 'cumulative' in name.lower() or 
                    'big' in name.lower() or 'shipping' in name.lower()  or 'pallets' in name.lower() or 
                    'delivery' in name.lower() or 'inspect return' in name.lower() or 
                    'rebox' in name.lower() or 'labor' in name.lower() or 
                    'overhead' in name.lower() or 'special marketing' in name.lower()) and not 'alliance' in name.lower()
  try:
    float(input_str)
    if is_percentage_field:
      return f"{input_str}%"
    elif is_dollar_field:
      return f"${input_str}"
    else: 
      return input_str
  except:
    return input_str

def validate_input_format(name, input_str):
  if not input_str or input_str.strip() == "":
    return True, ""
  try:
    auto_format_input(name, input_str)
    return True, ""
  except:
    return False, "Invalid number format"

def format_dataframe_for_display(df):
  display_df = df.copy()
  if 'Metric' in display_df.columns:
    metric_col_idx = display_df.columns.get_loc('Metric')
    display_df.insert(metric_col_idx + 1, ' ', '')
    original_data_cols = [col for col in df.columns if col != 'Metric']
    blank_col_counter = 1
    for i in range(2, len(original_data_cols), 2):
      target_col = original_data_cols[i]
      current_col_idx = display_df.columns.get_loc(target_col)
      blank_col_name = ' ' * (blank_col_counter + 1)
      display_df.insert(current_col_idx, blank_col_name, '')
      blank_col_counter += 1
  blank_row_triggers = ['Defect %', 'NET SALES', 'MARGIN %', 'SG&A', 'FACTORING %']
  rows_to_insert = []
  for idx in display_df.index:
    if 'Metric' in display_df.columns:
      metric_value = str(display_df.loc[idx, 'Metric'])
    else:
      metric_value = str(idx)
    if any(trigger in metric_value for trigger in blank_row_triggers):
      rows_to_insert.append(idx)
  for row_idx in reversed(rows_to_insert):
    blank_row = pd.Series([''] * len(display_df.columns), index=display_df.columns)
    display_df = pd.concat([
      display_df.iloc[:display_df.index.get_loc(row_idx) + 1],
      pd.DataFrame([blank_row]),
      display_df.iloc[display_df.index.get_loc(row_idx) + 1:]
    ], ignore_index=True)
  for col in display_df.columns:
    for idx in display_df.index:
      value = display_df.loc[idx, col]
      if pd.isna(value) or value == '' or str(value).strip() == '':
        display_df.loc[idx, col] = ""
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
      except:
        if pd.isna(value) or value == '' or str(value).strip() == '':
          display_df.loc[idx, col] = ""
        else:
          display_df.loc[idx, col] = str(value)
  return display_df

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

def password_entered():
  username = cursor.execute('''SELECT * FROM users WHERE username = ?''', (st.session_state["username"],)).fetchone()
  print(username)
  if not username:
    st.toast("**Username does not exist**", icon="🚨")
  elif hashlib.scrypt(st.session_state["password"].encode(), salt=b'secret_salt', n=16384, r=8, p=1).hex() == username[3]:
    st.session_state["authenticated"] = True
    st.session_state["username"] = username[1]
    st.session_state["fullName"] = username[2]
    st.session_state['role'] = username[4]
    st.session_state["id"] = username[0]
  else:
    st.toast("**Incorrect username or password**", icon="🚨")
    st.session_state["authenticated"] = False

def check_password():
  _, col2, _ = st.columns([2, 1, 2])
  with col2:
      if "authenticated" not in st.session_state or st.session_state["authenticated"] is False:
        with st.container(border=True):
          st.image("logo-spectra-premium.jpg", width=300)
          st.session_state["username"] = st.text_input("Username", key="username_input")
          st.session_state["password"] = st.text_input("Password", type="password", key="password_input")
          with st.container(horizontal=True):
            st.button("Login", on_click=password_entered, key="login_button", use_container_width=True)
            st.button("Register", type="secondary", key="register_button", on_click=register, use_container_width=True)
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
      st.markdown("""
      </div>
      """, unsafe_allow_html=True)

  def process_file(uploaded_file, user_defaults_df=None, volume=0.0):
    return standard.getSummary(uploaded_file, user_defaults_df, volume)

  main_container = st.container(border=True)
  with main_container:
    if st.session_state['role'] == 'Admin':
      input_tab, summary_tab, bca_matrix_tab, bca_files_tab, users_tab = st.tabs(["Original File", "Summary Comparison", "BCA Matrix", "BCA Files", "User Management"])
    else:
      input_tab, summary_tab, bca_matrix_tab, bca_files_tab = st.tabs(["Original File", "Summary Comparison", "BCA Matrix", "BCA Files"])

    with input_tab:
      st.header("Upload Original BCA File")
      st.write("Upload your original Business Case Analysis file. The system will automatically detect based on the filename.")
      input_file = st.file_uploader(
        "Choose your original Excel file", 
        label_visibility="collapsed",
        type=['xlsx', 'xls']
      )
      if input_file is not None:
        file_type = getFileType(input_file.name)
        if st.session_state.input_file != input_file:
          st.toast(f"Starting {file_type} Calculations...")
          st.session_state.input_file = input_file
          try:
            with st.spinner('Processing file... This may take a moment.'):
              output, assumption_test, defaults_test, st.session_state.format = process_file(input_file)
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
        if st.session_state.input_defaults_df is not None or st.session_state.input_assumptions_df is not None:
          try:
            header, date = input_file.name.split(".")[0].split("-")
          except:
            header, date = "Unformatted File Name", "Unknown"
          st.subheader(f'Business Case Analysis: {header}')
          st.write(f"**Date:** {date}")
          st.write('You can adjust the following parameters to see their impact on the analysis:')
          qty_col1, qty_col2, qty_col3 = st.columns([1,2,1])
          with qty_col2:
            st.write("#### Quantity Volume Control")
            with st.container(border=True):
              column1, column2 = st.columns(2)
              with column1:
                st.write('Volume Percentage Change:', formatter('Volume Percentage Change', 0.0))
              with column2:
                if "volume_text" not in st.session_state:
                  st.session_state.volume_text = "0.0%"
                
                volume_text = st.text_input(
                  label="Volume Percentage Change", 
                  value=st.session_state.volume_text,
                  key="volume_input", 
                  label_visibility="collapsed", 
                  help="Adjust the volume percentage to see how it affects the analysis (e.g., 5.0% for 5% increase)"
                )
                
                formatted_volume = auto_format_input("Volume Percentage Change", volume_text)
                
                is_valid, error_msg = validate_input_format("Volume Percentage Change", formatted_volume)
                if not is_valid:
                  st.error(error_msg)
                  volume = 0.0
                else:
                  volume = parse_input_value("Volume Percentage Change", formatted_volume)
                  if formatted_volume != volume_text:
                    st.session_state.volume_text = formatted_volume
                    st.rerun()
                  else:
                    st.session_state.volume_text = formatted_volume
          column1, column2 = st.columns(2)
          with column1:
            st.markdown("#### Assumptions")
            with st.expander("See Assumptions", expanded=True):
              if st.session_state.input_assumptions_df is None:
                st.write("No assumptions found in the original file.")
              else:
                header_col1, header_col2, header_col3 = st.columns([3, 1.5, 1.5])
                with header_col1:
                  st.markdown("**Parameter**")
                with header_col2:
                  st.markdown("**Original Value**")
                with header_col3:
                  st.markdown("**Your Value**")
                for x in st.session_state.input_assumptions_df:
                  param_col1, param_col2, param_col3 = st.columns([3, 1.5, 1.5])
                  
                  with param_col1:
                    st.markdown(f"{x}")
                  
                  with param_col2:
                    original_formatted = format_value_for_input(x, st.session_state.input_assumptions_df[x])
                    st.markdown(f"`{original_formatted}`")
                  
                  with param_col3:
                    text_key = f"assumption_text_{x}"
                    st.session_state[text_key] = format_value_for_input(x, st.session_state.user_assumptions_df[x])

                    text_value = st.text_input(
                      label=f"Edit {x}", 
                      value=st.session_state[text_key],
                      key=f"assumption_{x}", 
                      label_visibility="collapsed",
                      help=f"Enter your value for {x}",
                      placeholder="Enter value..."
                    )

                    formatted_value = auto_format_input(x, text_value)
                    st.session_state.user_assumptions_df[x] = parse_input_value(x, formatted_value)
                    if formatted_value != text_value:
                      st.session_state[text_key] = formatted_value
                      st.rerun()
                    else:
                      st.session_state[text_key] = formatted_value
          with column2:
            st.markdown("#### Defaults")
            with st.expander("See Defaults", expanded=True):
              if st.session_state.input_defaults_df is None:
                st.write("No defaults found in the original file.")
              elif file_type == "Standard BCA":
                st.session_state.lines = {n.split(' ', 1)[0] for n in st.session_state.input_defaults_df}
                st.session_state.metrics = {n.split(' ', 1)[1] for n in st.session_state.input_defaults_df}
                st.session_state.selected = st.selectbox("Select product line", options=st.session_state.lines)
                header_col1, header_col2, header_col3 = st.columns([3, 1.5, 1.5])
                with header_col1:
                  st.markdown("**Parameter**")
                with header_col2:
                  st.markdown("**Original Value**")
                with header_col3:
                  st.markdown("**Your Value**")
                for x in st.session_state.metrics:
                  if x not in st.session_state.format.index:
                    continue
                  param_col1, param_col2, param_col3 = st.columns([3, 1.5, 1.5])
                  
                  with param_col1:
                    st.markdown(f'{x}')
                  
                  with param_col2:
                    original_formatted = format_value_for_input(f'{st.session_state.selected} {x}', st.session_state.input_defaults_df[f'{st.session_state.selected} {x}'])
                    st.markdown(f"`{original_formatted}`")
                  
                  with param_col3:
                    text_key = f"defaults_text_{st.session_state.selected} {x}"
                    st.session_state[text_key] = format_value_for_input(f'{st.session_state.selected} {x}', st.session_state.user_defaults_df[f'{st.session_state.selected} {x}'])
                    text_value = st.text_input(
                      label=f"Edit f'{st.session_state.selected} {x}'", 
                      value=st.session_state[text_key],
                      key=f"defaults_{st.session_state.selected} {x}", 
                      label_visibility="collapsed",
                      help=f"Enter your value for f'{st.session_state.selected} {x}'",
                      placeholder="Enter value..."
                    )

                    formatted_value = auto_format_input(f'{st.session_state.selected} {x}', text_value)
                    
                    is_valid, error_msg = validate_input_format(f'{st.session_state.selected} {x}', formatted_value)
                    if not is_valid:
                      st.error(error_msg)
                    else:
                      st.session_state.user_defaults_df[f'{st.session_state.selected} {x}'] = parse_input_value(f'{st.session_state.selected} {x}', formatted_value)
                      if formatted_value != text_value:
                        st.session_state[text_key] = formatted_value
                        st.rerun()
                      else:
                        st.session_state[text_key] = formatted_value  
              else:
                header_col1, header_col2, header_col3 = st.columns([3, 1.5, 1.5])
                with header_col1:
                  st.markdown("**Parameter**")
                with header_col2:
                  st.markdown("**Original Value**")
                with header_col3:
                  st.markdown("**Your Value**")
                for x in st.session_state.input_defaults_df:
                  param_col1, param_col2, param_col3 = st.columns([3, 1.5, 1.5])
                  
                  with param_col1:
                    st.markdown(f"{x.replace('-', '')}")
                  
                  with param_col2:
                    original_formatted = format_value_for_input(x, st.session_state.input_defaults_df[x])
                    st.markdown(f"`{original_formatted}`")
                  
                  with param_col3:
                    text_key = f"defaults_text_{x}"
                    st.session_state[text_key] = format_value_for_input(x, st.session_state.input_defaults_df[x])
                    
                    text_value = st.text_input(
                      label=f"Edit {x}", 
                      value=st.session_state[text_key],
                      key=f"defaults_{x}", 
                      label_visibility="collapsed",
                      help=f"Enter your value for {x}",
                      placeholder="Enter value..."
                    )
                    
                    formatted_value = auto_format_input(x, text_value)
                    
                    is_valid, error_msg = validate_input_format(x, formatted_value)
                    if not is_valid:
                      st.error(error_msg)
                    else:
                      st.session_state.user_defaults_df[x] = parse_input_value(x, formatted_value)
                      if formatted_value != text_value:
                        st.session_state[text_key] = formatted_value
                        st.rerun()
                      else:
                        st.session_state[text_key] = formatted_value  
          apply_changes = st.button("Calculate with Modified Values", key="apply_changes_button", use_container_width=True)
          if apply_changes:
            st.toast("Calculating with modified values... This may take a moment.")
            output, assumption_test, defaults_test, format_test = process_file(input_file, (st.session_state.user_defaults_df or {}) | (st.session_state.user_assumptions_df or {}), volume)
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
            "Values":(st.session_state.input_assumptions_df or {}) | (st.session_state.input_defaults_df or {})
          })

          excel_bytes = ExcelExport.summaryBCA(st.session_state.input_summary_df, df, st.session_state.lines, st.session_state.metrics)

          bca_files_folder = os.path.join(os.getcwd(), "BCA Files")
          os.makedirs(bca_files_folder, exist_ok=True)
          bca_file_path = os.path.join(bca_files_folder, file_name)

          with open(bca_file_path, 'wb') as f:
            f.write(excel_bytes)

          input_files_folder = os.path.join(os.getcwd(), "Input Files")
          os.makedirs(input_files_folder, exist_ok=True)
          input_file_path = os.path.join(input_files_folder, input_file.name)
          
          with open(input_file_path, 'wb') as f:
            f.write(input_file.getvalue())
          
          file_id = str(uuid.uuid4())
          check = cursor.execute('''SELECT * FROM files WHERE user_id = ? AND input_fileName = ?''', (st.session_state["id"], input_file.name)).fetchone()
          if check is None: 
            cursor.execute('''INSERT INTO files (id, user_id, input_fileName, output_fileName) VALUES (?, ?, ?, ?)''', (file_id, st.session_state["id"], input_file.name, file_name))
          else:
            cursor.execute('''UPDATE files SET output_fileName = ? WHERE user_id = ? AND input_fileName = ?''', (file_name, st.session_state["id"], input_file.name))
          conn.commit()


          st.download_button(
            label="Download Summary",
            data=ExcelExport.summaryBCA(st.session_state.input_summary_df, df, st.session_state.lines, st.session_state.metrics),
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
              description = f"{volume*100}% increase"
            else:
              description = f"{volume*100}% decrease"
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
              format_dataframe_for_display(result_with_metrics),
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
                "parameter_changes": {changes_df.to_string(index=False)},
                "financial_impact_differences": {result_with_metrics.to_string(index=False)}
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
            default_input = pd.DataFrame({
              "Values":(st.session_state.input_assumptions_df or {}) | (st.session_state.input_defaults_df or {})
            })
            default_user = pd.DataFrame({
              "Values":(st.session_state.user_assumptions_df or {}) | (st.session_state.user_defaults_df or {})
            })
            st.download_button(
              label="Download Overall Summary and Comparison",
              data=ExcelExport.overallBCA(input_df, summary_df, default_input, default_user, st.session_state.lines, st.session_state.metrics),
              file_name=file_name,
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              key="comparison_download"
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

    with bca_matrix_tab:
      BCA_Matrix_tab.show()

    with bca_files_tab:
      st.header("All BCA Files")
      records = cursor.execute('''SELECT
        files.id as "File ID", 
        users.fullName as "Full Name of User",
        files.input_fileName as "Input Filename", 
        files.output_fileName as "BCA Filename", 
        files.upload_date as "Upload Date"
      FROM files 
      JOIN users ON files.user_id = users.id''')
      df_bca = pd.DataFrame(records, columns=["File ID", "Full Name of User", "Input Filename", "BCA Filename", "Upload Date"])
      df_bca["Upload Date"] = pd.to_datetime(df_bca["Upload Date"], format='mixed').dt.strftime("%B %d, %Y - %I:%M %p")
      df_bca.set_index("File ID", inplace=True)
      
      if df_bca.empty:
        st.info("No BCA files found. Upload files in the 'Original File' tab to see them listed here.")
      else:
        st.write("Select a row to download files:")
        
        gb = GridOptionsBuilder.from_dataframe(df_bca)
        gb.configure_selection(selection_mode='single', use_checkbox=True)
        gb.configure_grid_options(domLayout='autoHeight')
        gb.configure_column("Full Name of User", flex=1)
        gb.configure_column("Input Filename", headerTooltip="Click row checkbox to download", flex=1)
        gb.configure_column("BCA Filename", headerTooltip="Click row checkbox to download", flex=1)
        gb.configure_column("Upload Date", flex=1)
        gb.configure_grid_options(onFirstDataRendered=auto_size)
        
        gridOptions = gb.build()

        grid_response = AgGrid(
            df_bca,
            gridOptions=gridOptions,
            theme='balham',
            fit_columns_on_grid_load=True,
            enable_enterprise_modules=False,
            reload_data=True,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            allow_unsafe_jscode=True 
        )
        
        selected_rows = grid_response['selected_rows']
        
        if selected_rows is not None and not selected_rows.empty:
          selected_row = selected_rows.iloc[0]
          input_filename = selected_row['Input Filename']
          output_filename = selected_row['BCA Filename']
          
          col1, col2 = st.columns(2)
          
          with col1:
            input_file_path = os.path.join(os.getcwd(), "Input Files", input_filename)
            if os.path.exists(input_file_path):
              with open(input_file_path, 'rb') as f:
                st.download_button(
                  label=f"📥 Download _**Input File**_",
                  data=f.read(),
                  file_name=input_filename,
                  type="primary",
                  use_container_width=True,
                  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  key=f"download_input_{selected_row.name}"
                )
            else:
              st.error(f"Input file not found: {input_filename}")
          
          with col2:
            output_file_path = os.path.join(os.getcwd(), "BCA Files", output_filename)
            if os.path.exists(output_file_path):
              with open(output_file_path, 'rb') as f:
                st.download_button(
                  label=f"📥 Download _**BCA Output File**_",
                  data=f.read(),
                  file_name=output_filename,
                  use_container_width=True,
                  type="primary",
                  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  key=f"download_output_{selected_row.name}"
                )
            else:
              st.error(f"BCA file not found: {output_filename}")

    if st.session_state['role'] == 'Admin':
      with users_tab:
        st.header("Users List")
        st.write("Prototype: Admin can add, delete, edit users here.")
        users = cursor.execute('''
  SELECT fullName as "User Full Name", username as "Username", role as "Role" FROM users
  ''')
        user_df = pd.DataFrame(users, columns=["User Full Name", "Username", "Role"])
        gb = GridOptionsBuilder.from_dataframe(user_df)
        gb.configure_grid_options(domLayout='autoHeight')
        gb.configure_column("User Full Name", flex=1)
        gb.configure_column("Username", flex=1)
        gb.configure_column("Role", flex=1)
        gb.configure_grid_options(onFirstDataRendered=auto_size)
        gridOptions = gb.build()
        AgGrid(
          user_df,
          theme='balham',
          gridOptions=gridOptions,
          enable_enterprise_modules=False,
          allow_unsafe_jscode=True,
          fit_columns_on_grid_load=True,
        )

  footer_col1, footer_col2, footer_col3 = st.columns([1, 2, 1])
  with footer_col2:
    st.markdown("""
    <div style="text-align: center; color: #6c757d; font-size: 14px; margin-top: 30px;">
      <p style="margin-bottom: 0px;"><strong>Rapid BCA Analyzer</strong></p>
      <a style="margin-bottom: 0px; text-decoration: none;" href="https://lernmoreconsulting.com">© 2025 LernMore Consulting</a>
      <p>Secure • Fast • Accurate</p>
    </div>
    """, unsafe_allow_html=True)
