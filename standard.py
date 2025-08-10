import pandas as pd
import utility_library as util

def getSummary(file, user_defaults_df=None, volume=0.0):
  print(user_defaults_df)
  FORMAT = pd.read_excel(file, sheet_name="Format", index_col="METRIC")
  NET_SALES_METRICS, MARGIN_METRICS = util.GET_NETSALES_MARGIN_METRICS(FORMAT)
  SGA_METRICS = [
    'Handling/Shipping',
    'Fill Rate Fines',
    'Inspect Return',
    'Return Allowance Put Away/Rebox',
    'Pallets / Wrapping',
    'Delivery Cost',
    'Special Marketing (We Control)']

  DATA = pd.read_excel(file, sheet_name="Data")
  PRODUCT_LINES = DATA['P Line'].unique()
  if user_defaults_df is not None:
    ASSUMPTIONS = pd.DataFrame()
    DEFAULTS = user_defaults_df
    for data in user_defaults_df:
      for pline in PRODUCT_LINES:
        if pline in data:
          row, col = data.split(' ', 1)
          ASSUMPTIONS.at[row, col] = user_defaults_df[data]
  else:
    ASSUMPTIONS = pd.read_excel(file, sheet_name="Defaults & Assumptions", index_col="P Line")
    DEFAULTS = {}
    for metric in FORMAT.index:
      if(pd.notna(FORMAT.loc[metric].iloc[0])):
        if "Tariffs" not in metric: 
          DEFAULTS[metric] = FORMAT.loc[metric].iloc[0]
  print(ASSUMPTIONS)

  columns = []
  for line in PRODUCT_LINES:
    columns.append(f'{line} Cumulative')
    columns.append(f'{line} Per Unit')

  output = pd.DataFrame(index=FORMAT.index, columns=columns)

  output = util.QTY_CALCULATIONS(PRODUCT_LINES, output, DATA, ASSUMPTIONS, volume)
  output = util.NET_SALES_CALCULATIONS(output, DATA, ASSUMPTIONS, NET_SALES_METRICS, PRODUCT_LINES, DEFAULTS, volume)
  output = util.MARGIN_CALCULATIONS(PRODUCT_LINES, MARGIN_METRICS, output, DATA, DEFAULTS, FORMAT)
  output = util.SGA_CALCULATIONS(PRODUCT_LINES, output, ASSUMPTIONS, DEFAULTS, SGA_METRICS)

  for line in PRODUCT_LINES:
    net_sales = output.loc['NET SALES', f'{line} Cumulative']
    output.loc['FACTORING %', f'{line} Cumulative'] = net_sales * DEFAULTS['FACTORING %']
    output.loc['FACTORING %', f'{line} Per Unit'] = util.getPerUnit('FACTORING %', f'{line} Cumulative', output)
    margin = output.loc['MARGIN', f'{line} Cumulative']
    SGA = output.loc['SG&A', f'{line} Cumulative']
    factoring = output.loc['FACTORING %', f'{line} Cumulative']
    output.loc['Contribution Margin', f'{line} Cumulative'] = margin - SGA - factoring
    output.loc['Contribution Margin', f'{line} Per Unit'] = util.getPerUnit('Contribution Margin', f'{line} Cumulative', output)
    output.loc['Contribution Margin %', f'{line} Cumulative'] = output.loc['Contribution Margin', f'{line} Cumulative'] / net_sales

  assumptions_list = {}
  for column in ASSUMPTIONS.columns:
    for index in ASSUMPTIONS.index:
      if 'Unnamed' in column:
        continue
      parameter_name = f"{index} {column}"
      value = ASSUMPTIONS.loc[index, column]
      assumptions_list[parameter_name] = round(float(value),4)

  print(output)
  print(DEFAULTS)
  print(assumptions_list)
  return output.reset_index().rename(columns={'METRIC': 'Metric'}), DEFAULTS, assumptions_list