import pandas as pd

def getSumGivenColumn(Pline, dataframe, column):
  if Pline == 'All':
    return dataframe[column].sum()
  else:
    if 'Pline' in dataframe.columns:
      filtered_data = dataframe[dataframe['Pline'] == Pline]
    elif 'P Line' in dataframe.columns:
      filtered_data = dataframe[dataframe['P Line'] == Pline]
    else:
      raise KeyError("Neither 'Pline' nor 'P Line' column found in dataframe.")
    total_gross = filtered_data[column].sum()
    return total_gross
  
def getQTYTotalAndDefect(dataframe):
  cumulative_columns = [col for col in dataframe.columns if 'Cumulative' in col]
  for column in cumulative_columns:
    QTY_gross = dataframe.loc[dataframe['Metric'] == 'QTY Gross', column].iloc[0]
    QTY_defect = dataframe.loc[dataframe['Metric'] == 'QTY Defect', column].iloc[0]
    dataframe.loc[dataframe['Metric'] == 'QTY Total', column] = QTY_gross + QTY_defect
    dataframe.loc[dataframe['Metric'] == 'Defect %', column] = QTY_defect / QTY_gross if QTY_gross != 0 else 0
  return dataframe

def getSumAndPerUnit(Pline, dataframe, column):
  if Pline == 'All':
    total = dataframe[column].sum()
    per_unit = total / dataframe['L12 Shipped'].sum() if dataframe['L12 Shipped'].sum() != 0 else 0
    return total, per_unit
  else:
    filtered_data = dataframe[dataframe['Pline'] == Pline] or dataframe[dataframe['P Line'] == Pline]
    total = filtered_data[column].sum()
    per_unit = total / filtered_data['L12 Shipped'].sum() if filtered_data['L12 Shipped'].sum() != 0 else 0
    return total, per_unit