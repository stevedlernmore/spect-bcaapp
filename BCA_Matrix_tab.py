import streamlit as st
from UtilityLibrary import BCA_Matrix, ExcelExport, getFileDetails
import pandas as pd

class BCA_Matrix_tab:
  
  def show():
    st.subheader("Upload Files", anchor=False)
    with st.expander("View details"):
      st.warning("**Please note the filename format `<customer name>-<date>.xlsx` or `<customer name>-<date>.xlsx`**", icon="⚠️")
      st.markdown("**Upload your Excel files containing BCA matrices. Supported formats: `.xlsx`, `.xls`.**")
      if 'input_files' not in st.session_state:
        st.session_state.input_files = st.file_uploader(
          "file input",
          accept_multiple_files=True,
          type=['xlsx', 'xls'],
          label_visibility="collapsed"
        )
      else:
        uploaded_files = st.file_uploader(
          "file input",
          accept_multiple_files=True,
          type=['xlsx', 'xls'],
          label_visibility="collapsed"
        )
        if uploaded_files:
          st.session_state.input_files = uploaded_files
    if st.session_state.input_files:
      st.success(f"**Detected {len(st.session_state.input_files)} file(s) successfully!**", icon="✅")
    if st.session_state.input_files:
      st.session_state.fileDetails = getFileDetails(st.session_state.input_files)
      column1, column2 = st.columns([2, 1])
      with column1:
        st.subheader("Detected Customers Details", anchor=False)
        with st.expander("View details"):
          _, Plines, Metrics, PerUnitMetrics = BCA_Matrix.getBCAs(st.session_state.input_files)
          st.dataframe(st.session_state.fileDetails, hide_index=True)
      with column2:
        st.subheader("Detected Product Lines", anchor=False)
        with st.expander("View details"):
          plines = pd.DataFrame(Plines, columns=["Product Line/s"])
          st.dataframe(plines, hide_index=True)
      st.subheader("Difference Matrix", anchor=False)
      column1, column2 = st.columns([1, 4])
      with column1:
        st.session_state.mode = st.radio("Select View", ["Cumulative", "Per Unit"], index=0)
      with column2:
        if st.session_state.mode == "Cumulative":
          selectedMetric = st.selectbox("Select Metric for Comparison", options=Metrics, index=Metrics.index("Contribution Margin"))
        else:
          selectedMetric = st.selectbox("Select Metric for Comparison", options=PerUnitMetrics, index=PerUnitMetrics.index("Contribution Margin"))
  
      comparison = pd.DataFrame(columns=['Customer'])
      comparison['Customer'] = st.session_state.fileDetails["Customer Name"].unique()
      for pline in Plines:
        values = []
        for customer in comparison['Customer']:
          try:
            if st.session_state.mode == "Per Unit":
              values.append(st.session_state[customer].loc[selectedMetric, f'{pline} Per Unit'])
            else:
              values.append(st.session_state[customer].loc[selectedMetric, f'{pline} Cumulative'])
          except:
            values.append('-')
        comparison[pline] = values

      styled_df = comparison.set_index('Customer').map(ExcelExport.table_format, metric=selectedMetric)
      styled = styled_df.style.map(ExcelExport.highlight_negative)

      st.dataframe(styled)
      descripancy = pd.DataFrame(BCA_Matrix.getDescripancy(comparison))
      st.subheader("Discrepancy Report", anchor=False)
      st.write("These are the **product lines** with price discrepancies from the other retailers. Consider the listed retailers to be outliers in certain product lines.")
      if descripancy.empty:
        st.success("No discrepancies found.", icon="✅")
      else:
        st.dataframe(descripancy, hide_index=True)
    else:
      st.error("Please upload at least one Excel file to proceed.", icon="❗")