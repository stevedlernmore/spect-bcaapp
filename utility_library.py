from ctypes import util
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

# STANDARD HELPER
def extract_parentheses_content(text):
  start = text.find('(')
  end = text.find(')')
  if start != -1 and end != -1 and start < end:
    content = text[start+1:end]
    return [item.strip() for item in content.split(',')]
  return []

# STANDARD HELPER
def getPerUnit(row, column, output):
  return output.loc[row, column] / output.loc['QTY Gross', column]

# STANDARD
def QTY_CALCULATIONS(PRODUCT_LINES, output, DATA, ASSUMPTIONS, volume):
  for line in PRODUCT_LINES:
    gross = getSumGivenColumn(line, DATA, 'Qty')
    output.loc['QTY Gross', f'{line} Cumulative'] = gross + (gross/100*volume)
    output.loc['Defect %', f'{line} Cumulative'] = ASSUMPTIONS.loc[line, 'Defect %']
  for line in PRODUCT_LINES:
    output.loc['QTY Defect', f'{line} Cumulative'] = output.loc['QTY Gross', f'{line} Cumulative'] * output.loc['Defect %', f'{line} Cumulative'] * -1
    output.loc['QTY Total', f'{line} Cumulative'] = output.loc['QTY Gross', f'{line} Cumulative'] + output.loc['QTY Defect', f'{line} Cumulative']
  return output

# STANDARD
def GET_NETSALES_MARGIN_METRICS(FORMAT):
  defect_idx = FORMAT.index.get_loc('Defect %')
  net_sales_idx = FORMAT.index.get_loc('NET SALES')
  total_variable_cost_idx = FORMAT.index.get_loc('TOTAL VARIABLE COST')

  net_sales_metrics = FORMAT.index[defect_idx+1:net_sales_idx]
  net_sales_list = net_sales_metrics.tolist()
  if 'Agency Rep' in net_sales_list:
    net_sales_list.remove('Agency Rep')
    net_sales_list.append('Agency Rep')
    net_sales_metrics = pd.Index(net_sales_list)

  margin_metrics = FORMAT.index[net_sales_idx+1:total_variable_cost_idx]
  margin_list = margin_metrics.tolist()
  if 'Scrap Return Rate' in margin_list:
    margin_list.remove('Scrap Return Rate')
    margin_list.append('Scrap Return Rate')
    margin_metrics = pd.Index(margin_list)

  return net_sales_metrics, margin_metrics

# STANDARD
def NET_SALES_CALCULATIONS(output, DATA, ASSUMPTIONS, NET_SALES_METRICS, PRODUCT_LINES, DEFAULTS, volume):
  for metric in NET_SALES_METRICS:
    specific_pline = extract_parentheses_content(metric)
    if metric == 'Defect':
      for line in PRODUCT_LINES:
        output.loc[metric, f'{line} Cumulative'] = output.loc['Sales', f'{line} Cumulative'] * ASSUMPTIONS.loc[line, 'Defect %']
        output.loc[metric, f'{line} Per Unit'] = getPerUnit(metric, f'{line} Cumulative', output)
    elif metric not in DEFAULTS:
      for line in PRODUCT_LINES:
        temp_sale = getSumGivenColumn(line, DATA, metric)
        output.loc[metric, f'{line} Cumulative'] = temp_sale + (temp_sale/100*volume)
        output.loc[metric, f'{line} Per Unit'] = getPerUnit(metric, f'{line} Cumulative', output)
    else:
      loss = -1
      if 'in Lieu' in metric:
        loss = 1
      for line in PRODUCT_LINES:
        if line in specific_pline or specific_pline == []:
          output.loc[metric, f'{line} Cumulative'] = output.loc['Sales', f'{line} Cumulative'] * DEFAULTS.get(metric) * loss
        else:
          output.loc[metric, f'{line} Cumulative'] = 0
        output.loc[metric, f'{line} Per Unit'] = getPerUnit(metric, f'{line} Cumulative', output)

  for line in PRODUCT_LINES:
    agency_rep = 0
    for metric in NET_SALES_METRICS:
      if metric != 'Agency Rep':
        agency_rep += output.loc[metric, f'{line} Cumulative']
    output.loc['Agency Rep', f'{line} Cumulative'] = agency_rep * -1 * DEFAULTS.get('Agency Rep', 0)
    output.loc['Agency Rep', f'{line} Per Unit'] = getPerUnit('Agency Rep', f'{line} Cumulative', output)

  for line in PRODUCT_LINES:
    net_sales = 0
    for metric in NET_SALES_METRICS:
      net_sales += output.loc[metric, f'{line} Cumulative']
    output.loc['NET SALES', f'{line} Cumulative'] = net_sales
    output.loc['NET SALES', f'{line} Per Unit'] = getPerUnit('NET SALES', f'{line} Cumulative', output)

  return output

# STANDARD
def MARGIN_CALCULATIONS(PRODUCT_LINES, MARGIN_METRICS, output, DATA, DEFAULTS, FORMAT):
  for line in PRODUCT_LINES:
    for metric in MARGIN_METRICS:
      try:
        if metric == 'Cost':
          output.loc[metric, f'{line} Cumulative'] = getSumGivenColumn(line, DATA, 'Total Cost')
        else:
          output.loc[metric, f'{line} Cumulative'] = getSumGivenColumn(line, DATA, metric)
      except:
        output.loc[metric, f'{line} Cumulative'] = 0
      output.loc[metric, f'{line} Per Unit'] = getPerUnit(metric, f'{line} Cumulative', output)

  for line in PRODUCT_LINES:
    scrap_return = 0
    for metric in MARGIN_METRICS:
      if metric != 'Scrap Return Rate':
        if FORMAT.loc[metric].iloc[1] == 1:
          print(f'{metric} included in scrap return rate')
          scrap_return += output.loc[metric, f'{line} Per Unit']
        else:
          print(f'{metric} not included in scrap return rate')
    return_allowance = output.loc['Return Allowance', f'{line} Cumulative']
    per_unit_sales = output.loc['NET SALES', f'{line} Per Unit']
    output.loc['Scrap Return Rate', f'{line} Cumulative'] = (1-DEFAULTS.get('Scrap Return Rate'))*(return_allowance/per_unit_sales)*scrap_return
    output.loc['Scrap Return Rate', f'{line} Per Unit'] = getPerUnit('Scrap Return Rate', f'{line} Cumulative', output)

  for line in PRODUCT_LINES:
    variable_cost = 0
    for metric in MARGIN_METRICS:
      variable_cost += output.loc[metric, f'{line} Cumulative']
    output.loc['TOTAL VARIABLE COST', f'{line} Cumulative'] = variable_cost
    output.loc['TOTAL VARIABLE COST', f'{line} Per Unit'] = getPerUnit('TOTAL VARIABLE COST', f'{line} Cumulative', output)
    net_sales = output.loc['NET SALES', f'{line} Cumulative']
    margin = net_sales - output.loc['TOTAL VARIABLE COST', f'{line} Cumulative']
    output.loc['MARGIN', f'{line} Cumulative'] = margin
    output.loc['MARGIN', f'{line} Per Unit'] = getPerUnit('MARGIN', f'{line} Cumulative', output)
    margin_perc = margin / net_sales
    output.loc['MARGIN %', f'{line} Cumulative'] = margin_perc

  return output

# STANDARD
def SGA_CALCULATIONS(PRODUCT_LINES, output, ASSUMPTIONS, DEFAULTS, SGA_METRICS):
  for line in PRODUCT_LINES:
    output.loc['Handling/Shipping', f'{line} Per Unit'] = ASSUMPTIONS.loc[line, 'Handling/Shipping']
    output.loc['Handling/Shipping', f'{line} Cumulative'] = ASSUMPTIONS.loc[line, 'Handling/Shipping'] * output.loc['QTY Gross', f'{line} Cumulative']
    output.loc['Fill Rate Fines', f'{line} Cumulative'] = output.loc['NET SALES', f'{line} Cumulative'] * DEFAULTS.get('Fill Rate Fines', 0)
    output.loc['Fill Rate Fines', f'{line} Per Unit'] = getPerUnit('Fill Rate Fines', f'{line} Cumulative', output)
    output.loc['Inspect Return', f'{line} Per Unit'] = ASSUMPTIONS.loc[line, 'Handling/Shipping']
    output.loc['Return Allowance Put Away/Rebox', f'{line} Per Unit'] = ASSUMPTIONS.loc[line, 'Handling/Shipping']+ASSUMPTIONS.loc[line, 'Return Allowance Put Away/Rebox']
    output.loc['Pallets / Wrapping', f'{line} Per Unit'] = ASSUMPTIONS.loc[line, 'Pallets / Wrapping']
    output.loc['Pallets / Wrapping', f'{line} Cumulative'] = ASSUMPTIONS.loc[line, 'Pallets / Wrapping'] * output.loc['QTY Gross', f'{line} Cumulative']
    output.loc['Delivery Cost', f'{line} Per Unit'] = ASSUMPTIONS.loc[line, 'Delivery Cost']
    output.loc['Delivery Cost', f'{line} Cumulative'] = ASSUMPTIONS.loc[line, 'Delivery Cost'] * output.loc['QTY Gross', f'{line} Cumulative']
    output.loc['Special Marketing (We Control)', f'{line} Per Unit'] = ASSUMPTIONS.loc[line, 'Special Marketing (We Control)']
    output.loc['Special Marketing (We Control)', f'{line} Cumulative'] = ASSUMPTIONS.loc[line, 'Special Marketing (We Control)'] * output.loc['QTY Gross', f'{line} Cumulative']
    return_allowance = output.loc['Return Allowance', f'{line} Cumulative']
    per_unit_sales = output.loc['NET SALES', f'{line} Per Unit']
    per_unit_inspect = output.loc['Inspect Return', f'{line} Per Unit']
    per_unit_rebox = output.loc['Return Allowance Put Away/Rebox', f'{line} Per Unit']
    if DEFAULTS.get('Scrap Return Rate') == 1:
      output.loc['Inspect Return', f'{line} Cumulative'] = 0
    else:
      output.loc['Inspect Return', f'{line} Cumulative'] = (return_allowance/per_unit_sales)*per_unit_inspect*-1
    output.loc['Return Allowance Put Away/Rebox', f'{line} Cumulative'] = (return_allowance/per_unit_sales)*per_unit_rebox*-1*(1-DEFAULTS.get('Scrap Return Rate', 0))
    SGA = 0
    for metric in SGA_METRICS:
      SGA += output.loc[metric, f'{line} Cumulative']
    output.loc['SG&A', f'{line} Cumulative'] = SGA
    output.loc['SG&A', f'{line} Per Unit'] = getPerUnit('SG&A', f'{line} Cumulative', output)

  return output