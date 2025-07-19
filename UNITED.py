import pandas as pd
import utility_library as util

def getSummary(file, user_defaults_df=None):
  def getQTYTotal(column):
    gross = output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
    defect = output.loc[output['Metric'] == 'QTY Defect', column].iloc[0]
    return gross + defect

  def getQTYDefect(column, percent):
    return output.loc[output['Metric'] == 'QTY Gross', column].iloc[0] * percent * -1

  def getSales(DATA):
    values = DATA['Annual usage'] * DATA['NEW Invoice']
    return values.sum()

  def getPerUnit(row, column):
    return output.loc[output['Metric'] == row, column].iloc[0] / output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]

  def getMarketingAllowance(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * DEFAULTS.iloc[0]['Marketing allowance']*-1

  def getWarrantyAllowance(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * DEFAULTS.iloc[0]['Warranty allowance']*-1

  def getReturnAllowance(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * DEFAULTS.iloc[0]['Return Allowance']*-1

  def getAgencyFees(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    marketing_allowance = output.loc[output['Metric'] == 'Marketing allowance', column].iloc[0]
    warranty_allowance = output.loc[output['Metric'] == 'Warranty allowance', column].iloc[0]
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    return DEFAULTS.iloc[0]['Agency fees'] * (sales + marketing_allowance + warranty_allowance + return_allowance) * -1

  def getNetSales(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    marketing_allowance = output.loc[output['Metric'] == 'Marketing allowance', column].iloc[0]
    warranty_allowance = output.loc[output['Metric'] == 'Warranty allowance', column].iloc[0]
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    agency_fees = output.loc[output['Metric'] == 'Agency fees', column].iloc[0]
    return sales + marketing_allowance + warranty_allowance + return_allowance + agency_fees

  def getFGPOCost(DATA):
    return (DATA['Annual usage'] * DATA['V43.3 LC cost']).sum() 

  def getDuty(DATA):
    return (DATA['Annual usage'] * DATA['Duty']).sum()

  def getFreight(DATA):
    return (DATA['Annual usage'] * DATA['Interco Freight']).sum()

  def getHandlingReceiving(DATA):
    return (DATA['Annual usage'] * DATA['Handling and receiving']).sum()

  def getScrapReturn(column):
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    F0G_PO_cost = output.loc[output['Metric'] == 'FG PO cost', column].iloc[0]
    variable_overhead = output.loc[output['Metric'] == 'Variable - Overhead', column].iloc[0]
    labor = output.loc[output['Metric'] == 'Labor', column].iloc[0]
    duty = output.loc[output['Metric'] == 'Duty', column].iloc[0]
    interco_freight = output.loc[output['Metric'] == 'Interco Freight', column].iloc[0]
    handling_receiving = output.loc[output['Metric'] == 'Handling and receiving', column].iloc[0]
    return (F0G_PO_cost + variable_overhead + labor + duty + interco_freight + handling_receiving)*(1- DEFAULTS.iloc[0]['Scrap Return % (Resellable %)'])*(return_allowance/net_sales)

  def getTotalVariableCost(column):
    FG_PO_cost = output.loc[output['Metric'] == 'FG PO cost', column].iloc[0]
    variable_overhead = output.loc[output['Metric'] == 'Variable - Overhead', column].iloc[0]
    labor = output.loc[output['Metric'] == 'Labor', column].iloc[0]
    duty = output.loc[output['Metric'] == 'Duty', column].iloc[0]
    scrap_return = output.loc[output['Metric'] == 'Scrap Return % (Resellable %)', column].iloc[0]
    interco_freight = output.loc[output['Metric'] == 'Interco Freight', column].iloc[0]
    handling_receiving = output.loc[output['Metric'] == 'Handling and receiving', column].iloc[0]
    return FG_PO_cost + variable_overhead + labor + duty + interco_freight + handling_receiving + scrap_return

  def getMargin(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    total_variable_cost = output.loc[output['Metric'] == 'TOTAL VARIABLE COST', column].iloc[0]
    return net_sales - total_variable_cost

  def getMarginPercent(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0]
    return margin / net_sales

  def getFillRateFines(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    agency_fees = output.loc[output['Metric'] == 'Agency fees', column].iloc[0]
    return (sales + agency_fees) * DEFAULTS.iloc[0]['Fill rate fines']

  def getInspectReturnPerUnit(column):
    column = column.replace('Cumulative', 'Per Unit')
    return output.loc[output['Metric'] == 'Handling/Shipping', column].iloc[0]

  def getInspectReturnCumulative(column):
    if DEFAULTS.iloc[0]['Scrap Return % (Resellable %)'] == 1:
      return 0
    else:
      return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
      column = column.replace('Cumulative', 'Per Unit')
      net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
      inspect_return = output.loc[output['Metric'] == 'Inspect of Return', column].iloc[0]
      return return_allowance/net_sales* inspect_return*-1

  def getPutAwayPerUnit(column):
    column = column.replace('Cumulative', 'Per Unit')
    return output.loc[output['Metric'] == 'Handling/Shipping', column].iloc[0] + 2.42

  def getPutAwayCumulative(column):
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    put_away = output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', column].iloc[0]
    return (1-DEFAULTS.iloc[0]['Scrap Return % (Resellable %)'])*-1*put_away*(return_allowance/net_sales)

  def getSGandA(column):
    handling_shipping = output.loc[output['Metric'] == 'Handling/Shipping', column].iloc[0]
    fill_rate_fines = output.loc[output['Metric'] == 'Fill Rate Fines', column].iloc[0]
    inspect_return = output.loc[output['Metric'] == 'Inspect of Return', column].iloc[0]
    put_away = output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', column].iloc[0]
    special_marketing = output.loc[output['Metric'] == 'Special Marketing (we control)', column].iloc[0]
    pallets_wrapping = output.loc[output['Metric'] == 'Pallets / Wrapping', column].iloc[0]
    delivery = output.loc[output['Metric'] == 'Delivery ', column].iloc[0]
    return handling_shipping + fill_rate_fines + inspect_return + put_away + special_marketing + pallets_wrapping + delivery

  def getPaymentTerm(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    return net_sales * DEFAULTS.iloc[0]['Payment term']

  def getContributionMargin(column):
    margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0]
    sg_and_a = output.loc[output['Metric'] == 'SG&A', column].iloc[0]
    payment_term = output.loc[output['Metric'] == 'Payment term', column].iloc[0]
    return margin - sg_and_a - payment_term

  def getContributionMarginPercent(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    contribution_margin = output.loc[output['Metric'] == 'Contribution Margin', column].iloc[0]
    return contribution_margin / net_sales

  METRICS = [
    'QTY Gross',
    'QTY Defect',
    'QTY Total',
    'Defect %',

    'Sales',
    'Marketing allowance',
    'Warranty allowance',
    'Return Allowance',
    'Agency fees',
    'NET SALES',

    'FG PO cost',
    'Variable - Overhead',
    'Labor',
    'Duty',
    'Scrap Return % (Resellable %)',
    'Interco Freight',
    'Handling and receiving',
    'TOTAL VARIABLE COST',
    'MARGIN',
    'MARGIN %',

    'Handling/Shipping',
    'Fill Rate Fines',
    'Inspect of Return',
    'Return Allowance Put Away Costs, for usable product',
    'Special Marketing (we control)',
    'Pallets / Wrapping',
    'Delivery ',
    'SG&A',

    'Payment term',

    'Contribution Margin',
    'Contribution Margin %'
  ]
  COLUMNS = [
    'Metric',
    'New Cumulative',
    'New Per Unit'
  ]

  print('Reading DATA sheet...')
  DATA = pd.read_excel(file, sheet_name='CIA-CIN DOM', header=1)

  print('Reading Defaults...')
  if user_defaults_df is not None:
    DEFAULTS = user_defaults_df.set_index('Parameter').T
  else:
    DEFAULTS = pd.read_excel(file, sheet_name='Defaults & Assumptions')

  print('Initializing output DataFrame...')
  output = pd.DataFrame(columns= COLUMNS)
  output['Metric'] = METRICS

  print('Calculating QTY Gross...')
  output.loc[output['Metric'] == 'QTY Gross', 'New Cumulative'] = util.getSumGivenColumn('All', DATA, 'Annual usage')

  print('Calculating QTY Defect %...')
  output.loc[output['Metric'] == 'Defect %', 'New Cumulative'] = DEFAULTS.iloc[0]['Defect %']

  print('Calculating QTY Defect...')
  output.loc[output['Metric'] == 'QTY Defect', 'New Cumulative'] = getQTYDefect('New Cumulative', DEFAULTS.iloc[0]['Defect %'])

  print('Calculating QTY Total...')
  output.loc[output['Metric'] == 'QTY Total', 'New Cumulative'] = getQTYTotal('New Cumulative')

  print('Calculating Sales...')
  output.loc[output['Metric'] == 'Sales', 'New Cumulative'] = getSales(DATA)
  output.loc[output['Metric'] == 'Sales', 'New Per Unit'] = getPerUnit('Sales', 'New Cumulative')

  print('Calculating Marketing allowance...')
  output.loc[output['Metric'] == 'Marketing allowance', 'New Cumulative'] = getMarketingAllowance('New Cumulative')
  output.loc[output['Metric'] == 'Marketing allowance', 'New Per Unit'] = getPerUnit('Marketing allowance', 'New Cumulative')

  print('Calculating Warranty allowance...')
  output.loc[output['Metric'] == 'Warranty allowance', 'New Cumulative'] = getWarrantyAllowance('New Cumulative')
  output.loc[output['Metric'] == 'Warranty allowance', 'New Per Unit'] = getPerUnit('Warranty allowance', 'New Cumulative')

  print('Calculating Return Allowance...')
  output.loc[output['Metric'] == 'Return Allowance', 'New Cumulative'] = getReturnAllowance('New Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'New Per Unit'] = getPerUnit('Return Allowance', 'New Cumulative')

  print('Calculating Agency fees...')
  output.loc[output['Metric'] == 'Agency fees', 'New Cumulative'] = getAgencyFees('New Cumulative')
  output.loc[output['Metric'] == 'Agency fees', 'New Per Unit'] = getPerUnit('Agency fees', 'New Cumulative')

  print('Calculating NET SALES...')
  output.loc[output['Metric'] == 'NET SALES', 'New Cumulative'] = getNetSales('New Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'New Per Unit'] = getPerUnit('NET SALES', 'New Cumulative')

  print('Calculating FG PO cost...')
  output.loc[output['Metric'] == 'FG PO cost', 'New Cumulative'] = getFGPOCost(DATA)
  output.loc[output['Metric'] == 'FG PO cost', 'New Per Unit'] = getPerUnit('FG PO cost', 'New Cumulative')

  print('Calculating Variable - Overhead...')
  output.loc[output['Metric'] == 'Variable - Overhead', 'New Cumulative'] = DEFAULTS.iloc[0]['Variable - Overhead Cumulative']
  output.loc[output['Metric'] == 'Variable - Overhead', 'New Per Unit'] = getPerUnit('Variable - Overhead', 'New Cumulative')

  print('Calculating Labor...')
  output.loc[output['Metric'] == 'Labor', 'New Cumulative'] = DEFAULTS.iloc[0]['Labor Cumulative']
  output.loc[output['Metric'] == 'Labor', 'New Per Unit'] = getPerUnit('Labor', 'New Cumulative')

  print('Calculating Duty...')
  output.loc[output['Metric'] == 'Duty', 'New Cumulative'] = getDuty(DATA)
  output.loc[output['Metric'] == 'Duty', 'New Per Unit'] = getPerUnit('Duty', 'New Cumulative')

  print('Calculating Interco Freight...')
  output.loc[output['Metric'] == 'Interco Freight', 'New Cumulative'] = getFreight(DATA)
  output.loc[output['Metric'] == 'Interco Freight', 'New Per Unit'] = getPerUnit('Interco Freight', 'New Cumulative')

  print('Calculating Handling and receiving...')
  output.loc[output['Metric'] == 'Handling and receiving', 'New Cumulative'] = getHandlingReceiving(DATA)
  output.loc[output['Metric'] == 'Handling and receiving', 'New Per Unit'] = getPerUnit('Handling and receiving', 'New Cumulative')

  print('Calculating Scrap Return...')
  output.loc[output['Metric'] == 'Scrap Return % (Resellable %)', 'New Cumulative'] = getScrapReturn('New Cumulative')
  output.loc[output['Metric'] == 'Scrap Return % (Resellable %)', 'New Per Unit'] = getPerUnit('Scrap Return % (Resellable %)', 'New Cumulative')

  print('Calculating TOTAL VARIABLE COST...')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'New Cumulative'] = getTotalVariableCost('New Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'New Per Unit'] = getPerUnit('TOTAL VARIABLE COST', 'New Cumulative')

  print('Calculating MARGIN...')
  output.loc[output['Metric'] == 'MARGIN', 'New Cumulative'] = getMargin('New Cumulative')
  output.loc[output['Metric'] == 'MARGIN', 'New Per Unit'] = getPerUnit('MARGIN', 'New Cumulative')

  print('Calculating MARGIN %...')
  output.loc[output['Metric'] == 'MARGIN %', 'New Cumulative'] = getMarginPercent('New Cumulative')

  print('Calculating Handling/Shipping...')
  output.loc[output['Metric'] == 'Handling/Shipping', 'New Per Unit'] = DEFAULTS.iloc[0]['Handling/Shipping Per Unit']
  output.loc[output['Metric'] == 'Handling/Shipping', 'New Cumulative'] = output.loc[output['Metric'] == 'Handling/Shipping', 'New Per Unit'].iloc[0] * output.loc[output['Metric'] == 'QTY Gross', 'New Cumulative'].iloc[0]

  print('Calculating Fill Rate Fines...')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'New Cumulative'] = getFillRateFines('New Cumulative')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'New Per Unit'] = getPerUnit('Fill Rate Fines', 'New Cumulative')

  print('Calculating Inspect Return...')
  output.loc[output['Metric'] == 'Inspect of Return', 'New Per Unit'] = getInspectReturnPerUnit('New Cumulative')
  output.loc[output['Metric'] == 'Inspect of Return', 'New Cumulative'] = getInspectReturnCumulative('New Cumulative')

  print('Calculating Return Allowance Put Away Costs, for usable product...')
  output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', 'New Per Unit'] = getPutAwayPerUnit('New Cumulative')
  output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', 'New Cumulative'] = getPutAwayCumulative('New Cumulative')

  print('Calculating Special Marketing (we control)...')
  output.loc[output['Metric'] == 'Special Marketing (we control)', 'New Cumulative'] = DEFAULTS.iloc[0]['Special Marketing Cumulative']
  output.loc[output['Metric'] == 'Special Marketing (we control)', 'New Per Unit'] = DEFAULTS.iloc[0]['Special Marketing Per Unit']

  print('Calculating Pallets / Wrapping...')
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'New Per Unit'] = DEFAULTS.iloc[0]['Pallets / Wrapping Per Unit']
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'New Cumulative'] = output.loc[output['Metric'] == 'Pallets / Wrapping', 'New Per Unit'].iloc[0] * output.loc[output['Metric'] == 'QTY Gross', 'New Cumulative'].iloc[0]

  print('Calculating Delivery ...')
  output.loc[output['Metric'] == 'Delivery ', 'New Per Unit'] = DEFAULTS.iloc[0]['Delivery Per Unit']
  output.loc[output['Metric'] == 'Delivery ', 'New Cumulative'] = output.loc[output['Metric'] == 'Delivery ', 'New Per Unit'].iloc[0] * output.loc[output['Metric'] == 'QTY Gross', 'New Cumulative'].iloc[0]

  print('Calculating SG&A...')
  output.loc[output['Metric'] == 'SG&A', 'New Cumulative'] = getSGandA('New Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'New Per Unit'] = getPerUnit('SG&A', 'New Cumulative')

  print('Calculating Payment term...')
  output.loc[output['Metric'] == 'Payment term', 'New Cumulative'] = getPaymentTerm('New Cumulative')
  output.loc[output['Metric'] == 'Payment term', 'New Per Unit'] = getPerUnit('Payment term', 'New Cumulative')

  print('Calculating Contribution Margin...')
  output.loc[output['Metric'] == 'Contribution Margin', 'New Cumulative'] = getContributionMargin('New Cumulative')
  output.loc[output['Metric'] == 'Contribution Margin', 'New Per Unit'] = getPerUnit('Contribution Margin', 'New Cumulative')

  print('Calculating Contribution Margin %...')
  output.loc[output['Metric'] == 'Contribution Margin %', 'New Cumulative'] = getContributionMarginPercent('New Cumulative')

  defaults_list = []
  for column in DEFAULTS.columns:
    for index in DEFAULTS.index:
      parameter_name = f"{column}"
      value = DEFAULTS.loc[index, column]
      defaults_list.append({'Parameter': parameter_name, 'Value': value})
  
  assumptions_df = pd.DataFrame(defaults_list)
  
  return output, assumptions_df