import pandas as pd
import utility_library as util

def getSummary(file, user_defaults_df=None, volume=0.0):
  def getQTYDefect(column, percent):
    return output.loc[output['Metric'] == 'QTY Gross', column].iloc[0] * percent * -1

  def getQTYTotal(column):
    gross = output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
    defect = output.loc[output['Metric'] == 'QTY Defect', column].iloc[0]
    return gross + defect

  def getPerUnit(row, column):
    return output.loc[output['Metric'] == row, column].iloc[0] / output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]

  def getPrimaryRebate(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * PRIMARY_REBATE * -1

  def getMDMVendorNetPromotionalAllowance(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * MDM_VENDOR_NET_PROMOTIONAL_ALLOWANCE * -1

  def getReturnAllowance(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * RETURN_ALLOWANCE * -1

  def getFreightAllowance(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * FREIGHT_ALLOWANCE * -1

  def getDefect(column):
    defect = output.loc[output['Metric'] == 'Defect %', column].iloc[0]
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    return sales * defect * -1

  def getAgenctRep(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    primary_rebate = output.loc[output['Metric'] == 'Primary Rebate', column].iloc[0]
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    defect = output.loc[output['Metric'] == 'Defect', column].iloc[0]
    return AGENCY_REP * (sales + primary_rebate + return_allowance + defect) * -1

  def getNetSales(column):
    sales = output.loc[output['Metric'] == 'Sales', column].iloc[0]
    primary_rebate = output.loc[output['Metric'] == 'Primary Rebate', column].iloc[0]
    mdm_vendor_net_promotional_allowance = output.loc[output['Metric'] == 'MDM / Vendor Net Promotional Allowance + Display Allowance ', column].iloc[0]
    agency_rep = output.loc[output['Metric'] == 'Agency Rep', column].iloc[0]
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    freight_allowance = output.loc[output['Metric'] == 'Freight Allowance ', column].iloc[0]
    defect = output.loc[output['Metric'] == 'Defect', column].iloc[0]
    return sales + primary_rebate + mdm_vendor_net_promotional_allowance + return_allowance + freight_allowance + defect + agency_rep

  def getScrapReturnRate(column):
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    po_cost = output.loc[output['Metric'] == 'PO Cost ', column].iloc[0]
    duty = output.loc[output['Metric'] == 'Duty', column].iloc[0]
    baseline_tariff = output.loc[output['Metric'] == 'Baseline Tariff', column].iloc[0]
    two_four_tariff = output.loc[output['Metric'] == '2/4 Tariff ', column].iloc[0]
    three_four_tariff = output.loc[output['Metric'] == '3/4 Tariff ', column].iloc[0]
    additional_25_tariff = output.loc[output['Metric'] == 'Additional 25% Tariff ', column].iloc[0]
    return (1-SCRAP_RETURN_RATE)*(return_allowance/net_sales)*(po_cost + duty + baseline_tariff + two_four_tariff + three_four_tariff + additional_25_tariff)

  def getVariableCost(column):
    po_cost = output.loc[output['Metric'] == 'PO Cost ', column].iloc[0]
    inbound_shipping = output.loc[output['Metric'] == 'Inbound Shipping', column].iloc[0]
    additional_costs = output.loc[output['Metric'] == 'Additional Costs', column].iloc[0]
    scrap_return_rate = getScrapReturnRate(column)
    duty = output.loc[output['Metric'] == 'Duty', column].iloc[0]
    baseline_tariff = output.loc[output['Metric'] == 'Baseline Tariff', column].iloc[0]
    two_four_tariff = output.loc[output['Metric'] == '2/4 Tariff ', column].iloc[0]
    three_four_tariff = output.loc[output['Metric'] == '3/4 Tariff ', column].iloc[0]
    additional_25_tariff = output.loc[output['Metric'] == 'Additional 25% Tariff ', column].iloc[0]
    
    return po_cost + inbound_shipping + additional_costs + scrap_return_rate + duty + baseline_tariff + two_four_tariff + three_four_tariff + additional_25_tariff

  def getMargin(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    variable_cost = output.loc[output['Metric'] == 'TOTAL VARIABLE COST', column].iloc[0]
    return net_sales - variable_cost

  def getMarginPercent(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0]
    return (margin / net_sales) if net_sales != 0 else 0

  def getHandlingShipping(column):
    gross = output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    handling_shipping = output.loc[output['Metric'] == 'Handling/Shipping', column].iloc[0]
    return (handling_shipping*gross)

  def getFillRateFines(column):
    return output.loc[output['Metric'] == 'NET SALES', column].iloc[0] * FILL_RATE_FINES

  def getInspectReturn(column):
    if SCRAP_RETURN_RATE == 1:
      return 0
    else:
      return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
      column = column.replace('Cumulative', 'Per Unit')
      net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
      inspect_return = output.loc[output['Metric'] == 'Inspect Return', column].iloc[0]
      return (return_allowance/net_sales*inspect_return)*-1 if net_sales != 0 else 0

  def getRebox(column):
    return_allowance = output.loc[output['Metric'] == 'Return Allowance', column].iloc[0]
    column = column.replace('Cumulative', 'Per Unit')
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    rebox = output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', column].iloc[0]
    return return_allowance/net_sales*rebox*-1*(1-SCRAP_RETURN_RATE) if net_sales != 0 else 0

  def getSGandA(column):
    handling_shipping = output.loc[output['Metric'] == 'Handling/Shipping', column].iloc[0]
    fill_rate_fines = output.loc[output['Metric'] == 'Fill Rate Fines', column].iloc[0]
    inspect_return = output.loc[output['Metric'] == 'Inspect Return', column].iloc[0]
    rebox = output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', column].iloc[0]
    pallets_wrapping = output.loc[output['Metric'] == 'Pallets / Wrapping', column].iloc[0]
    delivery = output.loc[output['Metric'] == 'Delivery', column].iloc[0]
    additional_marketing = output.loc[output['Metric'] == 'Additional Marketing ', column].iloc[0]
    
    return handling_shipping + fill_rate_fines + inspect_return + rebox + pallets_wrapping + delivery + additional_marketing

  def getContributionMargin(column):
    margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0]
    sg_and_a = output.loc[output['Metric'] == 'SG&A', column].iloc[0]
    factoring = output.loc[output['Metric'] == 'FACTORING %', column].iloc[0]
    return margin - sg_and_a - factoring

  def getContributionMarginPercent(column):
    net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
    contribution_margin = output.loc[output['Metric'] == 'Contribution Margin', column].iloc[0]
    return (contribution_margin/net_sales) if contribution_margin != 0 else 0

  METRICS = [
    'QTY Gross',
    'QTY Defect',
    'QTY Total',
    'Defect %',

    'Sales',
    'Primary Rebate',
    'MDM / Vendor Net Promotional Allowance + Display Allowance ',
    'Agency Rep',
    'Return Allowance',
    'Freight Allowance ',
    'Defect',
    'NET SALES',

    'PO Cost ',
    'Inbound Shipping',
    'Additional Costs',
    'Scrap Return Rate',
    'Duty',
    'Baseline Tariff',
    '2/4 Tariff ',
    '3/4 Tariff ',
    'Additional 25% Tariff ',
    'TOTAL VARIABLE COST',
    'MARGIN',
    'MARGIN %',

    'Handling/Shipping',
    'Fill Rate Fines',
    'Inspect Return',
    'Return Allowance Put Away/Rebox',
    'Pallets / Wrapping',
    'Delivery',
    'Additional Marketing ',
    'SG&A',

    'FACTORING %',

    'Contribution Margin',
    'Contribution Margin %'
  ]
  COLUMNS = [
    'Metric',
    'FM Cumulative',
    'FM Per Unit',
    'SU Cumulative',
    'SU Per Unit',
  ]

  print('Reading DATA sheet...')
  DATA = pd.read_excel(file, sheet_name='Data')

  if user_defaults_df is not None:
    print('Using user defaults...')
    print(user_defaults_df)
    ASSUMPTIONS = pd.DataFrame()
    DEFAULTS = pd.DataFrame()
    for data in user_defaults_df:
      if 'SU' in data or 'FM' in data:
        row, col = data.split(' ', 1)
        ASSUMPTIONS.at[row, col] = user_defaults_df[data]
      else:
        DEFAULTS = pd.DataFrame({'Values': user_defaults_df}).T
  else:
    ASSUMPTIONS = pd.read_excel(file, sheet_name='Defaults & Assumptions', header=4, index_col=0)
    print('Reading Defaults...')
    DEFAULTS = pd.read_excel(file, sheet_name='Defaults & Assumptions', nrows=2)


  PRIMARY_REBATE = DEFAULTS['Primary Rebate'].iloc[0]
  MDM_VENDOR_NET_PROMOTIONAL_ALLOWANCE = DEFAULTS['MDM / Vendor Net Promotional Allowance + Display Allowance '].iloc[0]
  AGENCY_REP = DEFAULTS['Agency Rep'].iloc[0]
  RETURN_ALLOWANCE = DEFAULTS['Return Allowance '].iloc[0]
  FREIGHT_ALLOWANCE = DEFAULTS['Freight Allowance '].iloc[0]
  SCRAP_RETURN_RATE = DEFAULTS['Scrap Return Rate'].iloc[0]
  FACTORING_PERCENT = DEFAULTS['Factoring'].iloc[0]
  FILL_RATE_FINES = DEFAULTS['Fill Rate Fines'].iloc[0]

  print('Initializing output DataFrame...')
  output = pd.DataFrame(columns= COLUMNS)
  output['Metric'] = METRICS

  print('Calculation QTY Gross...')
  gross_FM = util.getSumGivenColumn('FM', DATA, 'L13 Sales')
  gross_SU = util.getSumGivenColumn('SU', DATA, 'L13 Sales')
  output.loc[output['Metric'] == 'QTY Gross', 'FM Cumulative'] = gross_FM+(gross_FM*volume/100)
  output.loc[output['Metric'] == 'QTY Gross', 'SU Cumulative'] = gross_SU+(gross_SU*volume/100)

  print('Calculation Defect Percent...')
  output.loc[output['Metric'] == 'Defect %', 'FM Cumulative'] = ASSUMPTIONS['Defect %'].loc['FM']
  output.loc[output['Metric'] == 'Defect %', 'SU Cumulative'] = ASSUMPTIONS['Defect %'].loc['SU']

  print('Calculation QTY Defect...')
  output.loc[output['Metric'] == 'QTY Defect', 'FM Cumulative'] = getQTYDefect('FM Cumulative', ASSUMPTIONS['Defect %'].loc['FM'])
  output.loc[output['Metric'] == 'QTY Defect', 'SU Cumulative'] = getQTYDefect('SU Cumulative', ASSUMPTIONS['Defect %'].loc['SU'])

  print('Calculation QTY Total...')
  output.loc[output['Metric'] == 'QTY Total', 'FM Cumulative'] = getQTYTotal('FM Cumulative')
  output.loc[output['Metric'] == 'QTY Total', 'SU Cumulative'] = getQTYTotal('SU Cumulative')

  print('Calculation Sales...')
  FM_sales = util.getSumGivenColumn('FM', DATA, 'Total Sales')
  FM_sales = FM_sales + (FM_sales * volume/100)
  SU_sales = util.getSumGivenColumn('SU', DATA, 'Total Sales')
  SU_sales = SU_sales + (SU_sales * volume/100)
  output.loc[output['Metric'] == 'Sales', 'FM Cumulative'] = FM_sales
  output.loc[output['Metric'] == 'Sales', 'SU Cumulative'] = SU_sales
  output.loc[output['Metric'] == 'Sales', 'FM Per Unit'] = getPerUnit('Sales', 'FM Cumulative')
  output.loc[output['Metric'] == 'Sales', 'SU Per Unit'] = getPerUnit('Sales', 'SU Cumulative')

  print('Calculation Primary Rebate...')
  output.loc[output['Metric'] == 'Primary Rebate', 'FM Cumulative'] = getPrimaryRebate('FM Cumulative')
  output.loc[output['Metric'] == 'Primary Rebate', 'SU Cumulative'] = getPrimaryRebate('SU Cumulative')
  output.loc[output['Metric'] == 'Primary Rebate', 'FM Per Unit'] = getPerUnit('Primary Rebate', 'FM Cumulative')
  output.loc[output['Metric'] == 'Primary Rebate', 'SU Per Unit'] = getPerUnit('Primary Rebate', 'SU Cumulative')

  print('Calculation MDM / Vendor Net Promotional Allowance + Display Allowance...')
  output.loc[output['Metric'] == 'MDM / Vendor Net Promotional Allowance + Display Allowance ', 'FM Cumulative'] = getMDMVendorNetPromotionalAllowance('FM Cumulative')
  output.loc[output['Metric'] == 'MDM / Vendor Net Promotional Allowance + Display Allowance ', 'SU Cumulative'] = getMDMVendorNetPromotionalAllowance('SU Cumulative')
  output.loc[output['Metric'] == 'MDM / Vendor Net Promotional Allowance + Display Allowance ', 'FM Per Unit'] = getPerUnit('MDM / Vendor Net Promotional Allowance + Display Allowance ', 'FM Cumulative')
  output.loc[output['Metric'] == 'MDM / Vendor Net Promotional Allowance + Display Allowance ', 'SU Per Unit'] = getPerUnit('MDM / Vendor Net Promotional Allowance + Display Allowance ', 'SU Cumulative')

  print('Calculation Return Allowance...')
  output.loc[output['Metric'] == 'Return Allowance', 'FM Cumulative'] = getReturnAllowance('FM Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'SU Cumulative'] = getReturnAllowance('SU Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'FM Per Unit'] = getPerUnit('Return Allowance', 'FM Cumulative')
  output.loc[output['Metric'] == 'Return Allowance', 'SU Per Unit'] = getPerUnit('Return Allowance', 'SU Cumulative')

  print('Calculation Freight Allowance...')
  output.loc[output['Metric'] == 'Freight Allowance ', 'FM Cumulative'] = getFreightAllowance('FM Cumulative')
  output.loc[output['Metric'] == 'Freight Allowance ', 'SU Cumulative'] = getFreightAllowance('SU Cumulative')
  output.loc[output['Metric'] == 'Freight Allowance ', 'FM Per Unit'] = getPerUnit('Freight Allowance ', 'FM Cumulative')
  output.loc[output['Metric'] == 'Freight Allowance ', 'SU Per Unit'] = getPerUnit('Freight Allowance ', 'SU Cumulative')

  print('Calculation Defect...')
  output.loc[output['Metric'] == 'Defect', 'FM Cumulative'] = getDefect('FM Cumulative')
  output.loc[output['Metric'] == 'Defect', 'SU Cumulative'] = getDefect('SU Cumulative')
  output.loc[output['Metric'] == 'Defect', 'FM Per Unit'] = getPerUnit('Defect', 'FM Cumulative')
  output.loc[output['Metric'] == 'Defect', 'SU Per Unit'] = getPerUnit('Defect', 'SU Cumulative')

  print('Calculation Agency Rep...')
  output.loc[output['Metric'] == 'Agency Rep', 'FM Cumulative'] = getAgenctRep('FM Cumulative')
  output.loc[output['Metric'] == 'Agency Rep', 'SU Cumulative'] = getAgenctRep('SU Cumulative')
  output.loc[output['Metric'] == 'Agency Rep', 'FM Per Unit'] = getPerUnit('Agency Rep', 'FM Cumulative')
  output.loc[output['Metric'] == 'Agency Rep', 'SU Per Unit'] = getPerUnit('Agency Rep', 'SU Cumulative')

  print('Calculation NET SALES...')
  output.loc[output['Metric'] == 'NET SALES', 'FM Cumulative'] = getNetSales('FM Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'SU Cumulative'] = getNetSales('SU Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'FM Per Unit'] = getPerUnit('NET SALES', 'FM Cumulative')
  output.loc[output['Metric'] == 'NET SALES', 'SU Per Unit'] = getPerUnit('NET SALES', 'SU Cumulative')

  print('Calculation PO Cost...')
  output.loc[output['Metric'] == 'PO Cost ', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, 'Total Buy Cost')
  output.loc[output['Metric'] == 'PO Cost ', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, 'Total Buy Cost')
  output.loc[output['Metric'] == 'PO Cost ', 'FM Per Unit'] = getPerUnit('PO Cost ', 'FM Cumulative')
  output.loc[output['Metric'] == 'PO Cost ', 'SU Per Unit'] = getPerUnit('PO Cost ', 'SU Cumulative')

  print('Calculation Inbound Shipping...')
  output.loc[output['Metric'] == 'Inbound Shipping', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, 'Total Freight')
  output.loc[output['Metric'] == 'Inbound Shipping', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, 'Total Freight')
  output.loc[output['Metric'] == 'Inbound Shipping', 'FM Per Unit'] = getPerUnit('Inbound Shipping', 'FM Cumulative')
  output.loc[output['Metric'] == 'Inbound Shipping', 'SU Per Unit'] = getPerUnit('Inbound Shipping', 'SU Cumulative')

  print('Calculation Additional Costs...')
  output.loc[output['Metric'] == 'Additional Costs', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, 'Total Additional Cost')
  output.loc[output['Metric'] == 'Additional Costs', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, 'Total Additional Cost')
  output.loc[output['Metric'] == 'Additional Costs', 'FM Per Unit'] = getPerUnit('Additional Costs', 'FM Cumulative')
  output.loc[output['Metric'] == 'Additional Costs', 'SU Per Unit'] = getPerUnit('Additional Costs', 'SU Cumulative')

  print('Calculation Duty...')
  output.loc[output['Metric'] == 'Duty', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, 'Total Duty')
  output.loc[output['Metric'] == 'Duty', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, 'Total Duty')
  output.loc[output['Metric'] == 'Duty', 'FM Per Unit'] = getPerUnit('Duty', 'FM Cumulative')
  output.loc[output['Metric'] == 'Duty', 'SU Per Unit'] = getPerUnit('Duty', 'SU Cumulative')

  print('Calculation Baseline Tariff...')
  output.loc[output['Metric'] == 'Baseline Tariff', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, 'Total Tariffs')
  output.loc[output['Metric'] == 'Baseline Tariff', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, 'Total Tariffs')
  output.loc[output['Metric'] == 'Baseline Tariff', 'FM Per Unit'] = getPerUnit('Baseline Tariff', 'FM Cumulative')
  output.loc[output['Metric'] == 'Baseline Tariff', 'SU Per Unit'] = getPerUnit('Baseline Tariff', 'SU Cumulative')

  print('Calculation 2/4 Tariff...')
  output.loc[output['Metric'] == '2/4 Tariff ', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, '2/4 Total')
  output.loc[output['Metric'] == '2/4 Tariff ', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, '2/4 Total')
  output.loc[output['Metric'] == '2/4 Tariff ', 'FM Per Unit'] = getPerUnit('2/4 Tariff ', 'FM Cumulative')
  output.loc[output['Metric'] == '2/4 Tariff ', 'SU Per Unit'] = getPerUnit('2/4 Tariff ', 'SU Cumulative')

  print('Calculation 3/4 Tariff...')
  output.loc[output['Metric'] == '3/4 Tariff ', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, '3/4 Total')
  output.loc[output['Metric'] == '3/4 Tariff ', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, '3/4 Total')
  output.loc[output['Metric'] == '3/4 Tariff ', 'FM Per Unit'] = getPerUnit('3/4 Tariff ', 'FM Cumulative')
  output.loc[output['Metric'] == '3/4 Tariff ', 'SU Per Unit'] = getPerUnit('3/4 Tariff ', 'SU Cumulative')

  print('Calculation Additional 25% Tariff...')
  output.loc[output['Metric'] == 'Additional 25% Tariff ', 'FM Cumulative'] = util.getSumGivenColumn('FM', DATA, "Add'l Total ")
  output.loc[output['Metric'] == 'Additional 25% Tariff ', 'SU Cumulative'] = util.getSumGivenColumn('SU', DATA, "Add'l Total ")
  output.loc[output['Metric'] == 'Additional 25% Tariff ', 'FM Per Unit'] = getPerUnit('Additional 25% Tariff ', 'FM Cumulative')
  output.loc[output['Metric'] == 'Additional 25% Tariff ', 'SU Per Unit'] = getPerUnit('Additional 25% Tariff ', 'SU Cumulative')

  print('Calculation Scrap Return Rate...')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'FM Cumulative'] = getScrapReturnRate('FM Cumulative')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'SU Cumulative'] = getScrapReturnRate('SU Cumulative')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'FM Per Unit'] = getPerUnit('Scrap Return Rate', 'FM Cumulative')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'SU Per Unit'] = getPerUnit('Scrap Return Rate', 'SU Cumulative')

  print('Calculation TOTAL VARIABLE COST...')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'FM Cumulative'] = getVariableCost('FM Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'SU Cumulative'] = getVariableCost('SU Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'FM Per Unit'] = getPerUnit('TOTAL VARIABLE COST', 'FM Cumulative')
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'SU Per Unit'] = getPerUnit('TOTAL VARIABLE COST', 'SU Cumulative')

  print('Calculation MARGIN...')
  output.loc[output['Metric'] == 'MARGIN', 'FM Cumulative'] = getMargin('FM Cumulative')
  output.loc[output['Metric'] == 'MARGIN', 'SU Cumulative'] = getMargin('SU Cumulative')
  output.loc[output['Metric'] == 'MARGIN', 'FM Per Unit'] = getPerUnit('MARGIN', 'FM Cumulative')
  output.loc[output['Metric'] == 'MARGIN', 'SU Per Unit'] = getPerUnit('MARGIN', 'SU Cumulative')

  print('Calculation MARGIN %...')
  output.loc[output['Metric'] == 'MARGIN %', 'FM Cumulative'] = getMarginPercent('FM Cumulative')
  output.loc[output['Metric'] == 'MARGIN %', 'SU Cumulative'] = getMarginPercent('SU Cumulative')

  print('Calculation Handling/Shipping...')
  output.loc[output['Metric'] == 'Handling/Shipping', 'FM Per Unit'] = ASSUMPTIONS['Handling/Shipping '].loc['FM']
  output.loc[output['Metric'] == 'Handling/Shipping', 'SU Per Unit'] = ASSUMPTIONS['Handling/Shipping '].loc['SU']
  output.loc[output['Metric'] == 'Handling/Shipping', 'FM Cumulative'] = getHandlingShipping('FM Cumulative')
  output.loc[output['Metric'] == 'Handling/Shipping', 'SU Cumulative'] = getHandlingShipping('SU Cumulative')

  print('Calculation Fill Rate Fines...')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'FM Cumulative'] = getFillRateFines('FM Cumulative')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'SU Cumulative'] = getFillRateFines('SU Cumulative')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'FM Per Unit'] = getPerUnit('Fill Rate Fines', 'FM Cumulative')
  output.loc[output['Metric'] == 'Fill Rate Fines', 'SU Per Unit'] = getPerUnit('Fill Rate Fines', 'SU Cumulative')

  print('Calculation Inspect Return...')
  output.loc[output['Metric'] == 'Inspect Return', 'FM Per Unit'] = ASSUMPTIONS['Handling/Shipping '].loc['FM']
  output.loc[output['Metric'] == 'Inspect Return', 'SU Per Unit'] = ASSUMPTIONS['Handling/Shipping '].loc['SU']
  output.loc[output['Metric'] == 'Inspect Return', 'FM Cumulative'] = getInspectReturn('FM Cumulative')
  output.loc[output['Metric'] == 'Inspect Return', 'SU Cumulative'] = getInspectReturn('SU Cumulative')

  print('Calculation Return Allowance Put Away/Rebox...')
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'FM Per Unit'] = ASSUMPTIONS['Handling/Shipping '].loc['FM'] + ASSUMPTIONS['Return Allowance Put Away/Rebox Add Amount'].loc['FM']
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'SU Per Unit'] = ASSUMPTIONS['Handling/Shipping '].loc['SU'] + ASSUMPTIONS['Return Allowance Put Away/Rebox Add Amount'].loc['SU']
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'FM Cumulative'] = getRebox('FM Cumulative')
  output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'SU Cumulative'] = getRebox('SU Cumulative')

  print('Calculation Pallets / Wrapping...')
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'FM Per Unit'] = ASSUMPTIONS['Pallets / Wrapping'].loc['FM']
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'SU Per Unit'] = ASSUMPTIONS['Pallets / Wrapping'].loc['SU']
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'FM Cumulative'] = ASSUMPTIONS['Pallets / Wrapping'].loc['FM'] * output.loc[output['Metric'] == 'QTY Gross', 'FM Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Pallets / Wrapping', 'SU Cumulative'] = ASSUMPTIONS['Pallets / Wrapping'].loc['SU'] * output.loc[output['Metric'] == 'QTY Gross', 'SU Cumulative'].iloc[0]

  print('Calculation Delivery...')
  output.loc[output['Metric'] == 'Delivery', 'FM Per Unit'] = ASSUMPTIONS['Delivery Cost '].loc['FM']
  output.loc[output['Metric'] == 'Delivery', 'SU Per Unit'] = ASSUMPTIONS['Delivery Cost '].loc['SU']
  output.loc[output['Metric'] == 'Delivery', 'FM Cumulative'] = ASSUMPTIONS['Delivery Cost '].loc['FM'] * output.loc[output['Metric'] == 'QTY Gross', 'FM Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Delivery', 'SU Cumulative'] = ASSUMPTIONS['Delivery Cost '].loc['SU'] * output.loc[output['Metric'] == 'QTY Gross', 'SU Cumulative'].iloc[0]

  print('Calculation Additional Marketing...')
  output.loc[output['Metric'] == 'Additional Marketing ', 'FM Per Unit'] = ASSUMPTIONS['Additional Marketing Per Unit'].loc['FM']
  output.loc[output['Metric'] == 'Additional Marketing ', 'SU Per Unit'] = ASSUMPTIONS['Additional Marketing Per Unit'].loc['SU']
  output.loc[output['Metric'] == 'Additional Marketing ', 'FM Cumulative'] = ASSUMPTIONS['Additional Marketing Cumulative'].loc['FM']
  output.loc[output['Metric'] == 'Additional Marketing ', 'SU Cumulative'] = ASSUMPTIONS['Additional Marketing Cumulative'].loc['SU']

  print('Calculation SG&A...')
  output.loc[output['Metric'] == 'SG&A', 'FM Cumulative'] = getSGandA('FM Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'SU Cumulative'] = getSGandA('SU Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'FM Per Unit'] = getPerUnit('SG&A', 'FM Cumulative')
  output.loc[output['Metric'] == 'SG&A', 'SU Per Unit'] = getPerUnit('SG&A', 'SU Cumulative')


  print('Calculation FACTORING %...')
  output.loc[output['Metric'] == 'FACTORING %', 'FM Cumulative'] = FACTORING_PERCENT * output.loc[output['Metric'] == 'NET SALES', 'FM Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'FACTORING %', 'SU Cumulative'] = FACTORING_PERCENT * output.loc[output['Metric'] == 'NET SALES', 'SU Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'FACTORING %', 'FM Per Unit'] = getPerUnit('FACTORING %', 'FM Cumulative')
  output.loc[output['Metric'] == 'FACTORING %', 'SU Per Unit'] = getPerUnit('FACTORING %', 'SU Cumulative')

  print('Calculation Contribution Margin...')
  output.loc[output['Metric'] == 'Contribution Margin', 'FM Cumulative'] = getContributionMargin('FM Cumulative')
  output.loc[output['Metric'] == 'Contribution Margin', 'SU Cumulative'] = getContributionMargin('SU Cumulative')
  output.loc[output['Metric'] == 'Contribution Margin', 'FM Per Unit'] = getPerUnit('Contribution Margin', 'FM Cumulative')
  output.loc[output['Metric'] == 'Contribution Margin', 'SU Per Unit'] = getPerUnit('Contribution Margin', 'SU Cumulative')

  print('Calculation Contribution Margin %...')
  output.loc[output['Metric'] == 'Contribution Margin %', 'FM Cumulative'] = getContributionMarginPercent('FM Cumulative')
  output.loc[output['Metric'] == 'Contribution Margin %', 'SU Cumulative'] = getContributionMarginPercent('SU Cumulative')
  print(ASSUMPTIONS)
  assumptions_list = {}
  for column in ASSUMPTIONS.columns:
    for index in ASSUMPTIONS.index:
      if 'Unnamed' in column:
        continue
      parameter_name = f"{index} {column}"
      value = ASSUMPTIONS.loc[index, column]
      assumptions_list[parameter_name] = round(float(value),4)
  
  defaults_list = {}
  for column in DEFAULTS.columns:
    for index in DEFAULTS.index:
      if 'Unnamed' in column:
        continue
      parameter_name = f"{column}"
      value = DEFAULTS.loc[index, column]
      defaults_list[parameter_name] = round(float(value),4)
  
  return output, assumptions_list, defaults_list
  