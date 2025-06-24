import pandas as pd
import utility_library as util

def getSummary(file):

  print('Reading PA DATA sheet...')
  DATA = pd.read_excel(file, sheet_name='Data', header=1, na_values='', keep_default_na=False)
  DATA = DATA.dropna(subset=['Pline'])
  print('Reading Default and Assumptions sheets...')
  DEFAULTS = pd.read_excel(file, sheet_name='Defaults & Assumptions')
  print('Reading Default and Assumptions sheets...')
  SGA_PER_UNIT = pd.read_excel(file, sheet_name='SG&A Per Unit', index_col=0).T

  # I'll assume that the per unit of freight interco is 0 since not defined in the Default and Assumptions sheet
  FREIGHT_INTERCO_PER_UNIT = 0.0
  SPECIAL_MARKETING_PER_UNIT = 0.0
  SPECIAL_MARKETING_CUMULATIVE = 0.0
  SCRAP_RETURN_RATE = DEFAULTS['Scrap Return Rate'][0]
  FILL_RATE_FINES = DEFAULTS['Fill Rate Fines'][0]
  FACTORING_PERCENT = DEFAULTS['FACTORING %'][0]

  NET_SALES_METRICS = [
    'Sales',
    'Rebate',
    'Logistic Rebate N/A',
    'Agency Rep (N.A. Williams)',
    'Return Allowance',
    'Defect Rebate in Lieu of Returns',
    'Defect',
  ]
  VARIABLE_METRICS = [
    'Cost',
    'Scrap Return Rate',
    'Variable - Overhead',
    'Labor',
    'Duty',
    'Tariffs',
    'Freight interco',
  ]
  SGA_METRICS = [
    'Handling/Shipping',
    'Fill Rate Fines',
    'Inspect Return',
    'Return Allowance Put Away/Rebox',
    'Pallets / Wrapping',
    'Delivery',
    'Special Marketing (We Control)',
  ]

  def getSummaryTotal(metrics_to_sum, metric_name, SGA=False):
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      total = 0
      for metric in metrics_to_sum:
        value = output.loc[output['Metric'] == metric, column].iloc[0]
        total += value
      output.loc[output['Metric'] == metric_name, column] = total
      if SGA:
        qty_total = output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'].iloc[0]
      else:
        qty_total = output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == metric_name, per_unit_column] = total / qty_total
      else:
        output.loc[output['Metric'] == metric_name, per_unit_column] = 0

  def getAllLineScrap():
    returnAllowance = output.loc[output['Metric'] == 'Return Allowance', 'All Lines Cumulative'].iloc[0]
    netSalesPerUnit = output.loc[output['Metric'] == 'NET SALES', 'All Lines Per Unit'].iloc[0]
    costPerUnit = output.loc[output['Metric'] == 'Cost', 'All Lines Per Unit'].iloc[0]
    variableOverheadPerUnit = output.loc[output['Metric'] == 'Variable - Overhead', 'All Lines Per Unit'].iloc[0]
    laborPerUnit = output.loc[output['Metric'] == 'Labor', 'All Lines Per Unit'].iloc[0]
    dutyPerUnit = output.loc[output['Metric'] == 'Duty', 'All Lines Per Unit'].iloc[0]
    tariffsPerUnit = output.loc[output['Metric'] == 'Tariffs', 'All Lines Per Unit'].iloc[0]
    return (1-SCRAP_RETURN_RATE)*(returnAllowance/netSalesPerUnit)*(costPerUnit + variableOverheadPerUnit + laborPerUnit + dutyPerUnit + tariffsPerUnit)

  def getMarginAndMarginPercent():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
      margin = net_sales - output.loc[output['Metric'] == 'TOTAL VARIABLE COST', column].iloc[0]
      output.loc[output['Metric'] == 'MARGIN', column] = margin
      output.loc[output['Metric'] == 'MARGIN %', column] = margin / net_sales if net_sales != 0 else 0

  def getInspectReturnCumulative(line):
    if SCRAP_RETURN_RATE == 1:
      return 0
    else:
      returnAllowance = output.loc[output['Metric'] == 'Return Allowance', f'{line} Cumulative'].iloc[0]
      netSalesPerUnit = output.loc[output['Metric'] == 'NET SALES', f'{line} Per Unit'].iloc[0]
      perUnit = output.loc[output['Metric'] == 'Inspect Return', f'{line} Per Unit'].iloc[0]
      return (returnAllowance/netSalesPerUnit)*perUnit*-1

  def getReturnAllowanceCumulative(line):
    returnAllowance = output.loc[output['Metric'] == 'Return Allowance', f'{line} Cumulative'].iloc[0]
    netSalesPerUnit = output.loc[output['Metric'] == 'NET SALES', f'{line} Per Unit'].iloc[0]
    perUnit = output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{line} Per Unit'].iloc[0]
    return (returnAllowance/netSalesPerUnit)*perUnit*-1*(1-SCRAP_RETURN_RATE)

  def getContributionMargin():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      contribution_margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0] - output.loc[output['Metric'] == 'FACTORING %', column].iloc[0] - output.loc[output['Metric'] == 'SG&A', column].iloc[0]
      output.loc[output['Metric'] == 'Contribution Margin', column] = contribution_margin
      qty_total = output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'Contribution Margin', per_unit_column] = contribution_margin / qty_total
      else:
        output.loc[output['Metric'] == 'Contribution Margin', per_unit_column] = 0
      contribution_margin_percent = (contribution_margin / output.loc[output['Metric'] == 'NET SALES', column].iloc[0]) if output.loc[output['Metric'] == 'NET SALES', column].iloc[0] != 0 else 0
      output.loc[output['Metric'] == 'Contribution Margin %', column] = contribution_margin_percent

  print('Preparing output structure...')
  METRICS = [
    'QTY Gross',
    'QTY Defect',
    'QTY Total',
    'Defect %',

    'Sales',
    'Rebate',
    'Logistic Rebate N/A',
    'Agency Rep (N.A. Williams)',
    'Return Allowance',
    'Defect Rebate in Lieu of Returns',
    'Defect',
    'NET SALES',

    'Cost',
    'Scrap Return Rate',
    'Variable - Overhead',
    'Labor',
    'Duty',
    'Tariffs',
    'Freight interco',
    'TOTAL VARIABLE COST',
    'MARGIN',
    'MARGIN %',

    'Handling/Shipping',
    'Fill Rate Fines',
    'Inspect Return',
    'Return Allowance Put Away/Rebox',
    'Pallets / Wrapping',
    'Delivery',
    'Special Marketing (We Control)',
    'SG&A',

    'FACTORING %',

    'Contribution Margin',
    'Contribution Margin %'
  ]
  COLUMNS = [
    'Metric',
    'All Lines Cumulative',
    'All Lines Per Unit',
  ]
  Plines = DATA['Pline'].unique()

  for Pline in Plines:
    COLUMNS.append(f'{Pline} Cumulative')
    COLUMNS.append(f'{Pline} Per Unit')

  output = pd.DataFrame(columns= COLUMNS)
  output['Metric'] = METRICS

  # Calculate QTY Gross, Defect, Total, and Defect %
  QTY_column = 'L12 Shipped'
  Defect_column = 'Defect Count'
  print('Calculating Sum of QTY Gross and Defect for all lines...')
  output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, QTY_column)
  output.loc[output['Metric'] == 'QTY Defect', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Defect_column)
  print('Calculating Sum of QTY Gross and Defect for each Pline...')
  for line in Plines:
    output.loc[output['Metric'] == 'QTY Gross', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, QTY_column)
    output.loc[output['Metric'] == 'QTY Defect', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Defect_column)
  output = util.getQTYTotalAndDefect(output)

  # Net Sales Data
  Sales_column = 'Total Sales'
  Rebate_column = 'VR Total'
  Logistic_Rebate_column = 'Logistcs Rebate N/A'
  Agency_Rep_column = 'Agency Rep (N.A. Williams) '
  Return_Allowance_column = 'Return Allowance '
  Defect_Rebate_column = 'Defect Rebate in Lieu of Returns'
  Defect_Total = 'Defect All Lines'
  Defect_Per_Line = 'Defect Plines'

  print('Calculating data for Net Sales for all lines...')
  output.loc[output['Metric'] == 'Sales', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Sales_column)
  output.loc[output['Metric'] == 'Rebate', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Rebate_column)*-1
  output.loc[output['Metric'] == 'Logistic Rebate N/A', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Logistic_Rebate_column)
  output.loc[output['Metric'] == 'Agency Rep (N.A. Williams)', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Agency_Rep_column)
  output.loc[output['Metric'] == 'Return Allowance', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Return_Allowance_column)
  output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Defect_Rebate_column)
  output.loc[output['Metric'] == 'Defect', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Defect_Total)

  QTY = output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Sales', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Sales', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Rebate', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Rebate', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Logistic Rebate N/A', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Logistic Rebate N/A', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Agency Rep (N.A. Williams)', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Agency Rep (N.A. Williams)', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Return Allowance', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Return Allowance', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Defect', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Defect', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0

  print('Calculating data for Net Sales for each Pline...')
  for line in Plines:
    output.loc[output['Metric'] == 'Sales', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Sales_column)
    output.loc[output['Metric'] == 'Rebate', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Rebate_column)*-1
    output.loc[output['Metric'] == 'Logistic Rebate N/A', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Logistic_Rebate_column)
    output.loc[output['Metric'] == 'Agency Rep (N.A. Williams)', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Agency_Rep_column)
    output.loc[output['Metric'] == 'Return Allowance', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Return_Allowance_column)
    output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Defect_Rebate_column)
    output.loc[output['Metric'] == 'Defect', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Defect_Per_Line)

    QTY = output.loc[output['Metric'] == 'QTY Gross', f'{line} Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Sales', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Sales', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Rebate', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Rebate', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Logistic Rebate N/A', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Logistic Rebate N/A', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Agency Rep (N.A. Williams)', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Agency Rep (N.A. Williams)', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Return Allowance', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Return Allowance', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Defect', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Defect', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0

  print('Calculating NET Sales...')
  getSummaryTotal(NET_SALES_METRICS, 'NET SALES')

  # VARIABLE COST DATA, MARGIN, and MARGIN %
  Cost_column_1 = 'Total Cost'
  Cost_column_2 = 'Total PO Cost'
  variable_overhead_column = 'Total Overhead'
  labor_column = 'Total Labor'
  duty_column = 'Total Duty'
  tariffs_column = 'Total Tariffs'
  freight_interco_column = 'Total Interco'
  scrap_return_column = 'Scrap Return PLINES'

  print('Calculating Total Variable Cost Data for all lines...')
  output.loc[output['Metric'] == 'Cost', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, Cost_column_1) + util.getSumGivenColumn('All', DATA, Cost_column_2)
  output.loc[output['Metric'] == 'Variable - Overhead', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, variable_overhead_column)
  output.loc[output['Metric'] == 'Labor', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, labor_column)
  output.loc[output['Metric'] == 'Duty', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, duty_column)
  output.loc[output['Metric'] == 'Tariffs', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, tariffs_column)
  output.loc[output['Metric'] == 'Freight interco', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, freight_interco_column)

  QTY = output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Cost', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Cost', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Variable - Overhead', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Variable - Overhead', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Labor', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Labor', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Duty', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Duty', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Tariffs', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Tariffs', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Freight interco', 'All Lines Per Unit'] = FREIGHT_INTERCO_PER_UNIT

  print('Calculating Total Variable Cost Data for each Pline...')
  for line in Plines:
    output.loc[output['Metric'] == 'Cost', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, Cost_column_1) + util.getSumGivenColumn(line, DATA, Cost_column_2)
    output.loc[output['Metric'] == 'Variable - Overhead', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, variable_overhead_column)
    output.loc[output['Metric'] == 'Labor', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, labor_column)
    output.loc[output['Metric'] == 'Duty', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, duty_column)
    output.loc[output['Metric'] == 'Tariffs', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, tariffs_column)
    output.loc[output['Metric'] == 'Freight interco', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, freight_interco_column)
    output.loc[output['Metric'] == 'Scrap Return Rate', f'{line} Cumulative'] = util.getSumGivenColumn(line, DATA, scrap_return_column)

    QTY = output.loc[output['Metric'] == 'QTY Gross', f'{line} Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Cost', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Cost', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Variable - Overhead', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Variable - Overhead', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Labor', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Labor', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Duty', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Duty', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Tariffs', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Tariffs', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Freight interco', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Freight interco', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Scrap Return Rate', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Scrap Return Rate', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0

  output.loc[output['Metric'] == 'Scrap Return Rate', 'All Lines Cumulative'] = getAllLineScrap()
  QTY = output.loc[output['Metric'] == 'QTY Gross', f'All Lines Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Scrap Return Rate', f'All Lines Per Unit'] = output.loc[output['Metric'] == 'Scrap Return Rate', f'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0

  print('Calculating TOTAL VARIABLE COST...')
  getSummaryTotal(VARIABLE_METRICS, 'TOTAL VARIABLE COST')
  getMarginAndMarginPercent()

  # SG&A and FACTORING %
  handling_shipping_column = 'Handling.1'

  print('Calculating SG&A Data for all lines...')
  QTY = output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Handling/Shipping', 'All Lines Per Unit'] = SGA_PER_UNIT.loc['Handling/Shipping ', 'ALL Lines']
  output.loc[output['Metric'] == 'Fill Rate Fines', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'Fill Rate Fines', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0
  output.loc[output['Metric'] == 'Inspect Return', 'All Lines Per Unit'] = SGA_PER_UNIT.loc['Inspect Return', 'ALL Lines']
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'All Lines Per Unit'] = SGA_PER_UNIT.loc['Reutrn Allowance Put Away/Rebox', 'ALL Lines']
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'All Lines Per Unit'] = SGA_PER_UNIT.loc['Pallets / Wrapping', 'ALL Lines']
  output.loc[output['Metric'] == 'Delivery', 'All Lines Per Unit'] = SGA_PER_UNIT.loc['Delivery ', 'ALL Lines']
  output.loc[output['Metric'] == 'Special Marketing (We Control)', 'All Lines Per Unit'] = SPECIAL_MARKETING_PER_UNIT

  output.loc[output['Metric'] == 'FACTORING %', 'All Lines Cumulative'] = FACTORING_PERCENT*output.loc[output['Metric'] == 'NET SALES', 'All Lines Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'FACTORING %', 'All Lines Per Unit'] = output.loc[output['Metric'] == 'FACTORING %', 'All Lines Cumulative'].iloc[0] / QTY if QTY != 0 else 0

  output.loc[output['Metric'] == 'Handling/Shipping', 'All Lines Cumulative'] = util.getSumGivenColumn('All', DATA, handling_shipping_column)
  output.loc[output['Metric'] == 'Fill Rate Fines', 'All Lines Cumulative'] = output.loc[output['Metric'] == 'NET SALES', 'All Lines Cumulative'].iloc[0] * FILL_RATE_FINES
  output.loc[output['Metric'] == 'Inspect Return', 'All Lines Cumulative'] = getInspectReturnCumulative('All Lines')
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'All Lines Cumulative'] = getReturnAllowanceCumulative('All Lines')
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'All Lines Cumulative'] = SGA_PER_UNIT.loc['Pallets / Wrapping', 'ALL Lines'] * QTY
  output.loc[output['Metric'] == 'Delivery', 'All Lines Cumulative'] = SGA_PER_UNIT.loc['Delivery ', 'ALL Lines'] * QTY
  output.loc[output['Metric'] == 'Special Marketing (We Control)', 'All Lines Cumulative'] = SPECIAL_MARKETING_CUMULATIVE

  print('Calculating SG&A Data for each Pline...')
  for line in Plines:
    QTY = output.loc[output['Metric'] == 'QTY Gross', f'{line} Cumulative'].iloc[0]

    output.loc[output['Metric'] == 'Handling/Shipping', f'{line} Per Unit'] = SGA_PER_UNIT.loc['Handling/Shipping ', line]
    output.loc[output['Metric'] == 'Fill Rate Fines', f'{line} Per Unit'] = output.loc[output['Metric'] == 'Fill Rate Fines', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0
    output.loc[output['Metric'] == 'Inspect Return', f'{line} Per Unit'] = SGA_PER_UNIT.loc['Inspect Return', line]
    output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{line} Per Unit'] = SGA_PER_UNIT.loc['Reutrn Allowance Put Away/Rebox', line]
    output.loc[output['Metric'] == 'Pallets / Wrapping', f'{line} Per Unit'] = SGA_PER_UNIT.loc['Pallets / Wrapping', line]
    output.loc[output['Metric'] == 'Delivery', f'{line} Per Unit'] = SGA_PER_UNIT.loc['Delivery ', line]
    output.loc[output['Metric'] == 'Special Marketing (We Control)', f'{line} Per Unit'] = SPECIAL_MARKETING_PER_UNIT
    
    output.loc[output['Metric'] == 'FACTORING %', f'{line} Cumulative'] = FACTORING_PERCENT*output.loc[output['Metric'] == 'NET SALES', f'{line} Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'FACTORING %', f'{line} Per Unit'] = output.loc[output['Metric'] == 'FACTORING %', f'{line} Cumulative'].iloc[0] / QTY if QTY != 0 else 0

    output.loc[output['Metric'] == 'Handling/Shipping', f'{line} Cumulative'] = SGA_PER_UNIT.loc['Handling/Shipping ', line] * QTY
    output.loc[output['Metric'] == 'Fill Rate Fines', f'{line} Cumulative'] = output.loc[output['Metric'] == 'NET SALES', f'{line} Cumulative'].iloc[0] * FILL_RATE_FINES
    output.loc[output['Metric'] == 'Inspect Return', f'{line} Cumulative'] = getInspectReturnCumulative(line)
    output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{line} Cumulative'] = getReturnAllowanceCumulative(line)
    output.loc[output['Metric'] == 'Pallets / Wrapping', f'{line} Cumulative'] = SGA_PER_UNIT.loc['Pallets / Wrapping', line] * QTY
    output.loc[output['Metric'] == 'Delivery', f'{line} Cumulative'] = SGA_PER_UNIT.loc['Delivery ', line] * QTY
    output.loc[output['Metric'] == 'Special Marketing (We Control)', f'{line} Cumulative'] = SPECIAL_MARKETING_CUMULATIVE

  print('Calculating SG&A for all lines...')
  getSummaryTotal(SGA_METRICS, 'SG&A', SGA=True)

  print('Calculating Contribution Margin...')
  getContributionMargin()

  assumptions_data = {
    'Parameter': [
      'Freight Interco Per Unit',
      'Special Marketing Per Unit',
      'Special Marketing Cumulative',
      'Scrap Return Rate',
      'Fill Rate Fines',
      'FACTORING %'
    ],
    'Value': [
      FREIGHT_INTERCO_PER_UNIT,
      SPECIAL_MARKETING_PER_UNIT,
      SPECIAL_MARKETING_CUMULATIVE,
      SCRAP_RETURN_RATE,
      FILL_RATE_FINES,
      FACTORING_PERCENT
    ]
  }
  
  assumptions_df = pd.DataFrame(assumptions_data)

  return output, assumptions_df