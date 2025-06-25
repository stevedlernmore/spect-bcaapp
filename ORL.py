import pandas as pd
import utility_library as util

input_fileName = "ORL_ETB.xlsx"
output_fileName = f"RESULT-{input_fileName}"

def getSummary(file):

  def getQTYDefect(column):
    return output.loc[output['Metric'] == 'QTY Gross', column].iloc[0] * DEFECT_PERCENT

  def getQTYTotal(column):
    return output.loc[output['Metric'] == 'QTY Gross', column].iloc[0] + getQTYDefect(column)

  def getPerUnit(row, column):
    return output.loc[output['Metric'] == row, column].iloc[0] / output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]

  def getOffInvoice(column):
    return output.loc[output['Metric'] == 'Sales', column].iloc[0] * OFF_INVOICE*-1

  def getRebates(column):
    return -1 * output.loc[output['Metric'] == 'Sales', column].iloc[0] * REBATES

  def getDefect(column):
    return DEFECT_PERCENT * output.loc[output['Metric'] == 'Sales', column].iloc[0]

  def getReturnAllowance(column):
    return -1 * output.loc[output['Metric'] == 'Sales', column].iloc[0] * RETURN_ALLOWANCE

  def getAgentRep(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    off_invoice = output.loc[output['Metric'] == 'Off Invoice', column].iloc[0]
    defect = output.loc[output['Metric'] == 'Defect', column].iloc[0]
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    return -1 * (sales + off_invoice + defect+ return_allowance) * AGENCY_REP

  def getCoreDevaluation(column):
    return CORE_DEVALUATION* output.loc[output['Metric'] == 'Sales', column].iloc[0]

  def getNetSales(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    off_invoice = output.loc[output['Metric'] == 'Off Invoice', column].iloc[0]
    rebates = output.loc[output['Metric'] == 'Rebates (Logistics 4%, Special 3%)', column].iloc[0]
    defect = output.loc[output['Metric'] == 'Defect', column].iloc[0]
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    agent_rep = output.loc[output['Metric'] == 'Agency Rep (ROLSM)', column].iloc[0]
    core_devaluation = output.loc[output['Metric'] == 'Core Devaluation', column].iloc[0]
    return sales + off_invoice + rebates + defect + return_allowance + agent_rep + core_devaluation

  def getScrapReturnRate(column):
    returnAllowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    cost = output.loc[output['Metric'] == 'Cost', column].iloc[0]
    variable_overhead = output.loc[output['Metric'] == 'Variable - Overhead', column].iloc[0]
    labor = output.loc[output['Metric'] == 'Labor', column].iloc[0]
    duty = output.loc[output['Metric'] == 'Duty', column].iloc[0]
    tariffs = output.loc[output['Metric'] == 'Tariffs', column].iloc[0]
    freight_interco = output.loc[output['Metric'] == 'Freight interco', column].iloc[0]
    return (1-SCRAP_RETURN_RATE)*(returnAllowance/net_sales)*(cost + variable_overhead + labor + duty + tariffs + freight_interco)

  def getTotalVariableCost(column):
    cost = output.loc[output['Metric'] == 'Cost', column].iloc[0]
    scrap_return_rate = output.loc[output['Metric'] == 'Scrap Return Rate', column].iloc[0]
    variable_overhead = output.loc[output['Metric'] == 'Variable - Overhead', column].iloc[0]
    labor = output.loc[output['Metric'] == 'Labor', column].iloc[0]
    duty = output.loc[output['Metric'] == 'Duty', column].iloc[0]
    tariffs = output.loc[output['Metric'] == 'Tariffs', column].iloc[0]
    freight_interco = output.loc[output['Metric'] == 'Freight interco', column].iloc[0]
    return cost + variable_overhead + labor + duty + tariffs + freight_interco + scrap_return_rate

  def getMarginAndMarginPercent(column, dataframe):
    total_variable_cost = dataframe.loc[output['Metric'] == 'TOTAL VARIABLE COST', column].iloc[0]
    net_sales = dataframe.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    margin = net_sales - total_variable_cost
    margin_percent = margin / net_sales if net_sales != 0 else 0
    dataframe.loc[output['Metric'] == 'MARGIN', column] = margin
    dataframe.loc[output['Metric'] == 'MARGIN %', column] = margin_percent
    return dataframe

  def getInspectReturn(column):
    if SCRAP_RETURN_RATE == 1:
      return 0
    else:
      returnAllowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
      column = column.replace('Cumulative', 'Per Unit')
      net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
      return returnAllowance/net_sales*HANDLING_SHIPPING*-1

  def getReturnAllowancePutAwayRebox(column):
    returnAllowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    per_unit = output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', column].iloc[0]
    return returnAllowance/net_sales*-1*(1-SCRAP_RETURN_RATE)*per_unit

  def getSGA(column):
    handling_shipping = output.loc[output['Metric'] == 'Handling/Shipping', column].iloc[0]
    fill_rate_fines = output.loc[output['Metric'] == 'Fill Rate Fines', column].iloc[0]
    inspect_return = output.loc[output['Metric'] == 'Inspect Return', column].iloc[0]
    return_allowance_put_away_rebox = output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', column].iloc[0]
    pallets_wrapping = output.loc[output['Metric'] == 'Pallets / Wrapping', column].iloc[0]
    delivery = output.loc[output['Metric'] == 'Delivery', column].iloc[0]
    special_marketing = output.loc[output['Metric'] == 'Special Marketing (We Control)', column].iloc[0]
    return handling_shipping + fill_rate_fines + inspect_return + return_allowance_put_away_rebox + pallets_wrapping + delivery + special_marketing

  def getContributionMargin():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      contribution_margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0] - output.loc[output['Metric'] == 'FACTORING %', column].iloc[0] - output.loc[output['Metric'] == 'SG&A', column].iloc[0]
      output.loc[output['Metric'] == 'Contribution Margin', column] = contribution_margin
      qty_total = output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'Contribution Margin', per_unit_column] = contribution_margin / qty_total
      else:
        output.loc[output['Metric'] == 'Contribution Margin', per_unit_column] = 0
      contribution_margin_percent = (contribution_margin / output.loc[output['Metric'] == 'NET SALES', column].iloc[0]) if output.loc[output['Metric'] == 'NET SALES', column].iloc[0] != 0 else 0
      output.loc[output['Metric'] == 'Contribution Margin %', column] = contribution_margin_percent

  print('Reading DATA sheet...')
  DATA = pd.read_excel(file, sheet_name='Data', header=1)
  print('Reading Assumptions...')
  ASSUMPTIONS = pd.read_excel(file, sheet_name='Defaults & Assumptions', header=3)
  ASSUMPTIONS = ASSUMPTIONS.iloc[:, 1:]
  print('Reading Defaults...')
  DEFAULTS = pd.read_excel(file, sheet_name='Defaults & Assumptions', nrows=2)
  DEFAULTS = DEFAULTS.iloc[:, 1:]

  METRICS = [
    'QTY Gross',
    'QTY Defect',
    'QTY Total',
    'Defect %',

    'Sales',
    'Off Invoice',
    'Rebates (Logistics 4%, Special 3%)',
    'Agency Rep (ROLSM)',
    'Return Allowance',
    'Core Devaluation',
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
    'GTN/CGT Current Cumulative',
    'GTN/CGT Current Per Unit',
    'GTN/CGT Blended Cumulative',
    'GTN/CGT Blended Per Unit',
  ]

  DEFECT_PERCENT = DEFAULTS['Defect %'][0]
  OFF_INVOICE = DEFAULTS['Off Invoice '][0]
  REBATES = DEFAULTS['Rebates (Logistics 4%, Special 3%)'][0]
  AGENCY_REP = DEFAULTS['Agency Rep (ROLSM)'][0]
  RETURN_ALLOWANCE = DEFAULTS['Return Allowance '][0]
  CORE_DEVALUATION = DEFAULTS['Core Devaluation '][0]
  SCRAP_RETURN_RATE = DEFAULTS['Scrap Return Rate'][0]
  print(ASSUMPTIONS.columns)
  VARIABLE_OVERHEAD = ASSUMPTIONS['Variable - Overhead'][0]
  LABOR = ASSUMPTIONS['Labor'][0]
  FREIGHT_INTERCO_CURRENT_PER_UNIT = ASSUMPTIONS['Freight interco \nper unit\ncurrent pricing'][0]
  FREIGHT_INTERCO_BLENDED_PER_UNIT = ASSUMPTIONS['Freight interco \nper unit\nblended pricing'][0]
  HANDLING_SHIPPING = ASSUMPTIONS['Handling/Shipping '][0]
  INSPECT_RETURN = ASSUMPTIONS['Inspect Return'][0]
  RETURN_ALLOWANCE_PUT_AWAY_REBOX = ASSUMPTIONS['Reutrn Allowance Put Away/Rebox'][0]
  FILL_RATE_FINES = DEFAULTS['Fill Rate Fines'][0]
  PALLETS_WRAPPING = ASSUMPTIONS['Pallets / Wrapping'][0]
  DELIVERY = ASSUMPTIONS['Delivery '][0]
  SPECIAL_MARKETING = ASSUMPTIONS['Special Marketing (We Control)'][0]
  FACTORING = DEFAULTS['FACTORING %'][0]

  print('Initializing output DataFrame...')
  output = pd.DataFrame(columns= COLUMNS)
  output['Metric'] = METRICS

  print('Calculation QTY Gross...')
  output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Current Cumulative'] = util.getSumGivenColumn('All', DATA, 'L12 Actual')
  output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Blended Cumulative'] = util.getSumGivenColumn('All', DATA, 'L12 Actual')

  print('Calculating Defect %...')
  output.loc[output['Metric'] == 'Defect %', 'GTN/CGT Current Cumulative'] = DEFECT_PERCENT
  output.loc[output['Metric'] == 'Defect %', 'GTN/CGT Blended Cumulative'] = DEFECT_PERCENT

  print('Calculating QTY Defect...')
  output.loc[output['Metric'] == 'QTY Defect', 'GTN/CGT Current Cumulative'] = getQTYDefect('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'QTY Defect', 'GTN/CGT Blended Cumulative'] = getQTYDefect('GTN/CGT Blended Cumulative')

  print('Calculating QTY Total...')
  output.loc[output['Metric'] == 'QTY Total', 'GTN/CGT Current Cumulative'] = getQTYTotal('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'QTY Total', 'GTN/CGT Blended Cumulative'] = getQTYTotal('GTN/CGT Blended Cumulative')

  print('Calculating Sales...')
  output.loc[output['Metric'] == 'Sales', 'GTN/CGT Current Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total Sales (Currrent)')
  output.loc[output['Metric'] == 'Sales', 'GTN/CGT Blended Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total Sales (Adjusted) ')
  output.loc[output['Metric'] == 'Sales', 'GTN/CGT Current Per Unit'] = getPerUnit('Sales', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Sales', 'GTN/CGT Blended Per Unit'] = getPerUnit('Sales', 'GTN/CGT Blended Cumulative')

  print('Calculating Off Invoice...')
  output.loc[output['Metric'] == 'Off Invoice', 'GTN/CGT Current Cumulative'] = getOffInvoice('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Off Invoice', 'GTN/CGT Blended Cumulative'] = getOffInvoice('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Off Invoice', 'GTN/CGT Current Per Unit'] = getPerUnit('Off Invoice', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Off Invoice', 'GTN/CGT Blended Per Unit'] = getPerUnit('Off Invoice', 'GTN/CGT Blended Cumulative')

  print('Calculating Rebates...')
  output.loc[output['Metric'] == 'Rebates (Logistics 4%, Special 3%)', 'GTN/CGT Current Cumulative'] = getRebates('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Rebates (Logistics 4%, Special 3%)', 'GTN/CGT Blended Cumulative'] = getRebates('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Rebates (Logistics 4%, Special 3%)', 'GTN/CGT Current Per Unit'] = getPerUnit('Rebates (Logistics 4%, Special 3%)', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Rebates (Logistics 4%, Special 3%)', 'GTN/CGT Blended Per Unit'] = getPerUnit('Rebates (Logistics 4%, Special 3%)', 'GTN/CGT Blended Cumulative')

  print('Calculating Defect...')
  output.loc[output['Metric'] == 'Defect', 'GTN/CGT Current Cumulative'] = getDefect('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Defect', 'GTN/CGT Blended Cumulative'] = getDefect('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Defect', 'GTN/CGT Current Per Unit'] = getPerUnit('Defect', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Defect', 'GTN/CGT Blended Per Unit'] = getPerUnit('Defect', 'GTN/CGT Blended Cumulative')

  print('Calculating Return Allowance...')
  output.loc[output['Metric'] == 'Return Allowance', 'GTN/CGT Current Cumulative'] = getReturnAllowance('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'GTN/CGT Blended Cumulative'] = getReturnAllowance('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'GTN/CGT Current Per Unit'] = getPerUnit('Return Allowance', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'GTN/CGT Blended Per Unit'] = getPerUnit('Return Allowance', 'GTN/CGT Blended Cumulative')

  print('Calculating Agency Rep (ROLSM)...')
  output.loc[output['Metric'] == 'Agency Rep (ROLSM)', 'GTN/CGT Current Cumulative'] = getAgentRep('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Agency Rep (ROLSM)', 'GTN/CGT Blended Cumulative'] = getAgentRep('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Agency Rep (ROLSM)', 'GTN/CGT Current Per Unit'] = getPerUnit('Agency Rep (ROLSM)', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Agency Rep (ROLSM)', 'GTN/CGT Blended Per Unit'] = getPerUnit('Agency Rep (ROLSM)', 'GTN/CGT Blended Cumulative')

  print('Calculating Core Devaluation...')
  output.loc[output['Metric'] == 'Core Devaluation', 'GTN/CGT Current Cumulative'] = getCoreDevaluation('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Core Devaluation', 'GTN/CGT Blended Cumulative'] = getCoreDevaluation('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Core Devaluation', 'GTN/CGT Current Per Unit'] = getPerUnit('Core Devaluation', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Core Devaluation', 'GTN/CGT Blended Per Unit'] = getPerUnit('Core Devaluation', 'GTN/CGT Blended Cumulative')

  print('Calculating NET SALES...')
  output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Current Cumulative'] = getNetSales('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Blended Cumulative'] = getNetSales('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Current Per Unit'] = getPerUnit('NET SALES', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Blended Per Unit'] = getPerUnit('NET SALES', 'GTN/CGT Blended Cumulative')

  print('Calculating Cost...')
  output.loc[output['Metric'] == 'Cost', 'GTN/CGT Current Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total PO Cost')
  output.loc[output['Metric'] == 'Cost', 'GTN/CGT Blended Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total PO Cost')
  output.loc[output['Metric'] == 'Cost', 'GTN/CGT Current Per Unit'] = getPerUnit('Cost', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Cost', 'GTN/CGT Blended Per Unit'] = getPerUnit('Cost', 'GTN/CGT Blended Cumulative')

  print('Calculating Labor...')
  output.loc[output['Metric'] == 'Labor', 'GTN/CGT Current Cumulative'] = LABOR
  output.loc[output['Metric'] == 'Labor', 'GTN/CGT Blended Cumulative'] = LABOR
  output.loc[output['Metric'] == 'Labor', 'GTN/CGT Current Per Unit'] = getPerUnit('Labor', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Labor', 'GTN/CGT Blended Per Unit'] = getPerUnit('Labor', 'GTN/CGT Blended Cumulative')

  print('Calculating Variable - Overhead...')
  output.loc[output['Metric'] == 'Variable - Overhead', 'GTN/CGT Current Cumulative'] = VARIABLE_OVERHEAD
  output.loc[output['Metric'] == 'Variable - Overhead', 'GTN/CGT Blended Cumulative'] = VARIABLE_OVERHEAD
  output.loc[output['Metric'] == 'Variable - Overhead', 'GTN/CGT Current Per Unit'] = getPerUnit('Variable - Overhead', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Variable - Overhead', 'GTN/CGT Blended Per Unit'] = getPerUnit('Variable - Overhead', 'GTN/CGT Blended Cumulative')

  print('Calculating Duty...')
  output.loc[output['Metric'] == 'Duty', 'GTN/CGT Current Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total Duty')
  output.loc[output['Metric'] == 'Duty', 'GTN/CGT Blended Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total Duty')
  output.loc[output['Metric'] == 'Duty', 'GTN/CGT Current Per Unit'] = getPerUnit('Duty', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Duty', 'GTN/CGT Blended Per Unit'] = getPerUnit('Duty', 'GTN/CGT Blended Cumulative')

  print('Calculating Tariffs...')
  output.loc[output['Metric'] == 'Tariffs', 'GTN/CGT Current Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total Tariffs')
  output.loc[output['Metric'] == 'Tariffs', 'GTN/CGT Blended Cumulative'] = util.getSumGivenColumn('All', DATA, 'Total Tariffs')
  output.loc[output['Metric'] == 'Tariffs', 'GTN/CGT Current Per Unit'] = getPerUnit('Tariffs', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Tariffs', 'GTN/CGT Blended Per Unit'] = getPerUnit('Tariffs', 'GTN/CGT Blended Cumulative')

  print('Calculating Freight interco...')
  output.loc[output['Metric'] == 'Freight interco', 'GTN/CGT Current Per Unit'] = FREIGHT_INTERCO_CURRENT_PER_UNIT
  output.loc[output['Metric'] == 'Freight interco', 'GTN/CGT Blended Per Unit'] = FREIGHT_INTERCO_BLENDED_PER_UNIT
  output.loc[output['Metric'] == 'Freight interco', 'GTN/CGT Current Cumulative'] = FREIGHT_INTERCO_CURRENT_PER_UNIT * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Freight interco', 'GTN/CGT Blended Cumulative'] = FREIGHT_INTERCO_BLENDED_PER_UNIT * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Blended Cumulative'].iloc[0]

  print('Calculate Scrap Return Rate...')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'GTN/CGT Current Cumulative'] = getScrapReturnRate('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'GTN/CGT Blended Cumulative'] = getScrapReturnRate('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'GTN/CGT Current Per Unit'] = getPerUnit('Scrap Return Rate', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'GTN/CGT Blended Per Unit'] = getPerUnit('Scrap Return Rate', 'GTN/CGT Blended Cumulative')

  print('Calculating TOTAL VARIABLE COST...')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'GTN/CGT Current Cumulative'] = getTotalVariableCost('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'GTN/CGT Blended Cumulative'] = getTotalVariableCost('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'GTN/CGT Current Per Unit'] = getPerUnit('TOTAL VARIABLE COST', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'GTN/CGT Blended Per Unit'] = getPerUnit('TOTAL VARIABLE COST', 'GTN/CGT Blended Cumulative')

  print('Calculating MARGIN and MARGIN %...')
  output = getMarginAndMarginPercent('GTN/CGT Current Cumulative', output)
  output = getMarginAndMarginPercent('GTN/CGT Blended Cumulative', output)
  output.loc[output['Metric'] == 'MARGIN', 'GTN/CGT Current Per Unit'] = getPerUnit('MARGIN', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'MARGIN', 'GTN/CGT Blended Per Unit'] = getPerUnit('MARGIN', 'GTN/CGT Blended Cumulative')

  print('Calculating Handling/Shipping...')
  output.loc[output['Metric'] == 'Handling/Shipping', 'GTN/CGT Current Cumulative'] = HANDLING_SHIPPING * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Handling/Shipping', 'GTN/CGT Blended Cumulative'] = HANDLING_SHIPPING * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Blended Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Handling/Shipping', 'GTN/CGT Current Per Unit'] = HANDLING_SHIPPING
  output.loc[output['Metric'] == 'Handling/Shipping', 'GTN/CGT Blended Per Unit'] = HANDLING_SHIPPING

  print('Calculating Fill Rate Fines...')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'GTN/CGT Current Cumulative'] = FILL_RATE_FINES * output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Fill Rate Fines', 'GTN/CGT Blended Cumulative'] = FILL_RATE_FINES * output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Blended Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Fill Rate Fines', 'GTN/CGT Current Per Unit'] = getPerUnit('Fill Rate Fines', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'GTN/CGT Blended Per Unit'] = getPerUnit('Fill Rate Fines', 'GTN/CGT Blended Cumulative')

  print('Calculating Inspect Return...')
  output.loc[output['Metric'] == 'Inspect Return', 'GTN/CGT Current Cumulative'] = getInspectReturn('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Inspect Return', 'GTN/CGT Blended Cumulative'] = getInspectReturn('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'Inspect Return', 'GTN/CGT Current Per Unit'] = INSPECT_RETURN
  output.loc[output['Metric'] == 'Inspect Return', 'GTN/CGT Blended Per Unit'] = INSPECT_RETURN

  print('Calculating Return Allowance Put Away/Rebox...')
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'GTN/CGT Current Per Unit'] = RETURN_ALLOWANCE_PUT_AWAY_REBOX
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'GTN/CGT Blended Per Unit'] = RETURN_ALLOWANCE_PUT_AWAY_REBOX
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'GTN/CGT Current Cumulative'] = getReturnAllowancePutAwayRebox('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'GTN/CGT Blended Cumulative'] = getReturnAllowancePutAwayRebox('GTN/CGT Blended Cumulative')

  print('Calculating Pallets / Wrapping...')
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'GTN/CGT Current Per Unit'] = PALLETS_WRAPPING
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'GTN/CGT Blended Per Unit'] = PALLETS_WRAPPING
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'GTN/CGT Current Cumulative'] = PALLETS_WRAPPING * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'GTN/CGT Blended Cumulative'] = PALLETS_WRAPPING * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Blended Cumulative'].iloc[0]

  print('Calculating Delivery...')
  output.loc[output['Metric'] == 'Delivery', 'GTN/CGT Current Cumulative'] = DELIVERY * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Delivery', 'GTN/CGT Blended Cumulative'] = DELIVERY * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Blended Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Delivery', 'GTN/CGT Current Per Unit'] = DELIVERY
  output.loc[output['Metric'] == 'Delivery', 'GTN/CGT Blended Per Unit'] = DELIVERY

  print('Calculating Special Marketing (We Control)...')
  output.loc[output['Metric'] == 'Special Marketing (We Control)', 'GTN/CGT Current Cumulative'] = SPECIAL_MARKETING * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Special Marketing (We Control)', 'GTN/CGT Blended Cumulative'] = SPECIAL_MARKETING * output.loc[output['Metric'] == 'QTY Gross', 'GTN/CGT Blended Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Special Marketing (We Control)', 'GTN/CGT Current Per Unit'] = SPECIAL_MARKETING
  output.loc[output['Metric'] == 'Special Marketing (We Control)', 'GTN/CGT Blended Per Unit'] = SPECIAL_MARKETING

  print('Calculating SG&A...')
  output.loc[output['Metric'] == 'SG&A', 'GTN/CGT Current Cumulative'] = getSGA('GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'GTN/CGT Blended Cumulative'] = getSGA('GTN/CGT Blended Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'GTN/CGT Current Per Unit'] = getPerUnit('SG&A', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'GTN/CGT Blended Per Unit'] = getPerUnit('SG&A', 'GTN/CGT Blended Cumulative')

  print('Calculating FACTORING %...')
  output.loc[output['Metric'] == 'FACTORING %', 'GTN/CGT Current Cumulative'] = FACTORING * output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Current Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'FACTORING %', 'GTN/CGT Blended Cumulative'] = FACTORING * output.loc[output['Metric'] == 'NET SALES', 'GTN/CGT Blended Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'FACTORING %', 'GTN/CGT Current Per Unit'] = getPerUnit('FACTORING %', 'GTN/CGT Current Cumulative')
  output.loc[output['Metric'] == 'FACTORING %', 'GTN/CGT Blended Per Unit'] = getPerUnit('FACTORING %', 'GTN/CGT Blended Cumulative')

  print('Calculating Contribution Margin...')
  getContributionMargin()

  assumptions_list = []
  for column in ASSUMPTIONS.columns:
    for index in ASSUMPTIONS.index:
      parameter_name = f"{column}"
      value = ASSUMPTIONS.loc[index, column]
      assumptions_list.append({'Parameter': parameter_name, 'Value': value})
  
  defaults_list = []
  for column in DEFAULTS.columns:
    for index in DEFAULTS.index:
      parameter_name = f"{column}"
      value = DEFAULTS.loc[index, column]
      defaults_list.append({'Parameter': parameter_name, 'Value': value})
  
  combined_assumptions = assumptions_list + defaults_list
  
  assumptions_df = pd.DataFrame(combined_assumptions)
  
  assumptions_df = assumptions_df.dropna()
  
  return output, assumptions_df