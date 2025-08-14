import pandas as pd
import utility_library as util

def getSummary(file, user_defaults_df=None, volume=0.0):

  def getSales(Pline, bulk=False):
    if Pline == 'All':
      if bulk:
        sales_total = DATA_V2['Sales (Modified)'].sum()
        sales_total = sales_total+(sales_total*volume)
        column_name = 'All Lines Bulk By Cumulative'
      else:
        sales_total = DATA_V2['Total Sales'].sum()
        sales_total = sales_total+(sales_total*volume)
        column_name = 'All Lines Cumulative'
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      sales_total = filtered_data['Total Sales'].sum()
      sales_total = sales_total+(sales_total*volume)
      column_name = f'{Pline} Cumulative'
    
    output.loc[output['Metric'] == 'Sales', column_name] = sales_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Sales', per_unit_column] = sales_total / qty_total
    else:
      output.loc[output['Metric'] == 'Sales', per_unit_column] = 0
    
    return sales_total

  def getRebatePlusAllowance(Pline, bulk=False):
    if bulk:
      rebate_col = 'Rebate.1' 
      payment_col = 'Payment.1'
      column_name = 'All Lines Bulk By Cumulative'
    else:
      rebate_col = 'Rebate'
      payment_col = 'Payment'
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      rebate_payment_total = DATA_V2[rebate_col].sum()*-1 + DATA_V2[payment_col].sum()*-1
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      rebate_payment_total = filtered_data[rebate_col].sum()*-1 + filtered_data[payment_col].sum()*-1
    output.loc[output['Metric'] == 'Rebate (10%) + Payment (2%)', column_name] = rebate_payment_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    if qty_total > 0:
      output.loc[output['Metric'] == 'Rebate (10%) + Payment (2%)', per_unit_column] = rebate_payment_total / qty_total
    else:
      output.loc[output['Metric'] == 'Rebate (10%) + Payment (2%)', per_unit_column] = 0
    return rebate_payment_total

  def getFreightAllowance(Pline, bulk=False):
    if bulk:
      freight_col = 'Freight Allowance.1'
      column_name = 'All Lines Bulk By Cumulative'
    else:
      freight_col = 'Freight Allowance'
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      freight_allowance_total = DATA_V2[freight_col].sum()*-1
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      freight_allowance_total = filtered_data[freight_col].sum()*-1
    output.loc[output['Metric'] == 'Freight Allowance', column_name] = freight_allowance_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    if qty_total > 0:
      output.loc[output['Metric'] == 'Freight Allowance', per_unit_column] = freight_allowance_total / qty_total
    else:
      output.loc[output['Metric'] == 'Freight Allowance', per_unit_column] = 0
    return freight_allowance_total

  def getAgencyRep(Pline, bulk=False):
    if bulk:
      agency_col = 'Rep Commission '
      column_name = 'All Lines Bulk By Cumulative'
    else:
      agency_col = 'Rep Commission'
      if Pline == 'All':
        agency_col = 'Rep Commission for All Lines'
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      if agency_col == 'Rep Commission for All Lines':
        total_sale = DATA_V2['Total Sales'].sum()
        total_defect = DATA_V2['Defect'].sum()*-1
        total_rebate = DATA_V2['Rebate'].sum()*-1
        total_payment = DATA_V2['Payment'].sum()*-1
        total_freight = DATA_V2['Freight Allowance'].sum()*-1
        total_return_allowance = DATA_V2['Return Allowance'].sum()*-1
        total_lieu_of_returns = DATA_V2['Defect Rebate in Lieu of Returns'].sum()*-1
        agency_rep_total = (total_sale + total_defect + total_rebate + total_payment + total_freight + total_return_allowance + total_lieu_of_returns)*REP_COMMISSION*-1
      else:
        agency_rep_total = DATA_V2[agency_col].sum()*-1
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      agency_rep_total = filtered_data[agency_col].sum()*-1
    output.loc[output['Metric'] == 'Agency Rep (RPS)', column_name] = agency_rep_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    if qty_total > 0:
      output.loc[output['Metric'] == 'Agency Rep (RPS)', per_unit_column] = agency_rep_total / qty_total
    else:
      output.loc[output['Metric'] == 'Agency Rep (RPS)', per_unit_column] = 0
    return agency_rep_total

  def getReturnAllowance(Pline, bulk=False):
    if bulk:
      return_col = 'Return Allowance.1'
      column_name = 'All Lines Bulk By Cumulative'
    else:
      return_col = 'Return Allowance'
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      return_allowance_total = DATA_V2[return_col].sum()*-1
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      return_allowance_total = filtered_data[return_col].sum()*-1
    output.loc[output['Metric'] == 'Return Allowance', column_name] = return_allowance_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    if qty_total > 0:
      output.loc[output['Metric'] == 'Return Allowance', per_unit_column] = return_allowance_total / qty_total
    else:
      output.loc[output['Metric'] == 'Return Allowance', per_unit_column] = 0
    return return_allowance_total

  def getDefectRebate(Pline, bulk=False):
    if bulk:
      defect_col = 'Defect Rebate in Lieu of Returns.1'
      column_name = 'All Lines Bulk By Cumulative'
    else:
      defect_col = 'Defect Rebate in Lieu of Returns'
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      defect_rebate_total = DATA_V2[defect_col].sum()*-1
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      defect_rebate_total = filtered_data[defect_col].sum()*-1
    output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', column_name] = defect_rebate_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    if qty_total > 0:
      output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', per_unit_column] = defect_rebate_total / qty_total
    else:
      output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', per_unit_column] = 0
    return defect_rebate_total

  def getDefect(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      defect_total = CUMULLATIVE_DEFECT 
    else:
      sales_value = output.loc[output['Metric'] == 'Sales', column_name].iloc[0]
      defect_percent = output.loc[output['Metric'] == 'Defect %', column_name].iloc[0]
      defect_total = sales_value * (defect_percent)
    
    output.loc[output['Metric'] == 'Defect', column_name] = defect_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Defect', per_unit_column] = defect_total / qty_total
    else:
      output.loc[output['Metric'] == 'Defect', per_unit_column] = 0
    
    return defect_total

  def getCost(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    
    if Pline == 'All':
      cost_total = DATA_V2['Total PO Cost'].sum()
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      cost_total = filtered_data['Total PO Cost'].sum()
    
    output.loc[output['Metric'] == 'Cost', column_name] = cost_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Cost', per_unit_column] = cost_total / qty_total
    else:
      output.loc[output['Metric'] == 'Cost', per_unit_column] = 0
    
    return cost_total

  def getScrapReturnRate(Pline, bulk=False):
    if bulk:
      scrap_col = 'Scrap Return Rate.1'
      column_name = 'All Lines Bulk By Cumulative'
    else:
      scrap_col = 'Scrap Return Rate'
      if Pline == 'All':
        scrap_col = 'Scrap Return Rate for All Lines'
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    if Pline == 'All':
      if scrap_col == 'Scrap Return Rate for All Lines':
        scrap_return_rate = getTotalScrapReturnRate()
      else:
        scrap_return_rate = DATA_V2[scrap_col].sum()*-1
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      scrap_return_rate = filtered_data[scrap_col].sum()*-1
    
    output.loc[output['Metric'] == 'Scrap Return Rate', column_name] = scrap_return_rate
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Scrap Return Rate', per_unit_column] = scrap_return_rate / qty_total
    else:
      output.loc[output['Metric'] == 'Scrap Return Rate', per_unit_column] = 0
    
    return scrap_return_rate

  def getHandlingOnFab(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    
    if Pline == 'All':
      handling_on_fab_total = DATA_V2['Total Fab Handling'].sum()
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      handling_on_fab_total = filtered_data['Total Fab Handling'].sum()
    
    output.loc[output['Metric'] == 'Handling on Fab', column_name] = handling_on_fab_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Handling on Fab', per_unit_column] = handling_on_fab_total / qty_total
    else:
      output.loc[output['Metric'] == 'Handling on Fab', per_unit_column] = 0
    
    return handling_on_fab_total

  def getDuty(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    
    if Pline == 'All':
      duty_total = DATA_V2['Total Duty'].sum()
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      duty_total = filtered_data['Total Duty'].sum()
    
    output.loc[output['Metric'] == 'Duty', column_name] = duty_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Duty', per_unit_column] = duty_total / qty_total
    else:
      output.loc[output['Metric'] == 'Duty', per_unit_column] = 0
    
    return duty_total

  def getBaseTariff(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    
    if Pline == 'All':
      base_tariff_total = DATA_V2['Total Tariff'].sum()
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      base_tariff_total = filtered_data['Total Tariff'].sum()
    
    output.loc[output['Metric'] == 'Base Tariff', column_name] = base_tariff_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Base Tariff', per_unit_column] = base_tariff_total / qty_total
    else:
      output.loc[output['Metric'] == 'Base Tariff', per_unit_column] = 0
    
    return base_tariff_total

  def getAddedTariff(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    
    if Pline == 'All':
      added_tariff_total = DATA_V2["Add'l Tariff Total"].sum()
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      added_tariff_total = filtered_data["Add'l Tariff Total"].sum()
    
    output.loc[output['Metric'] == 'Added Tariff', column_name] = added_tariff_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      output.loc[output['Metric'] == 'Added Tariff', per_unit_column] = added_tariff_total / qty_total
    else:
      output.loc[output['Metric'] == 'Added Tariff', per_unit_column] = 0
    
    return added_tariff_total

  def getFreightInterco(Pline, bulk=False):
    if bulk:
      column_name = 'All Lines Bulk By Cumulative'
    else:
      if Pline == 'All':
        column_name = 'All Lines Cumulative'
      else:
        column_name = f'{Pline} Cumulative'
    
    if Pline == 'All':
      freight_interco_total = DATA_V2['Total Interco'].sum()
    else:
      filtered_data = DATA_V2[DATA_V2['Pline'] == Pline]
      freight_interco_total = filtered_data['Total Interco'].sum()
    
    output.loc[output['Metric'] == 'Freight interco', column_name] = freight_interco_total
    qty_total = output.loc[output['Metric'] == 'QTY Total', column_name].iloc[0]
    per_unit_column = column_name.replace('Cumulative', 'Per Unit')
    
    if qty_total > 0:
      if 'All Lines' in column_name:
        output.loc[output['Metric'] == 'Freight interco', per_unit_column] = FREIGHT_INTERCO
      else:
        output.loc[output['Metric'] == 'Freight interco', per_unit_column] = freight_interco_total / qty_total
    else:
      output.loc[output['Metric'] == 'Freight interco', per_unit_column] = 0
    
    return freight_interco_total

  def getNETSales():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    metrics_to_sum = [
      'Sales',
      'Rebate (10%) + Payment (2%)',
      'Freight Allowance', 
      'Agency Rep (RPS)',
      'Return Allowance',
      'Defect Rebate in Lieu of Returns',
      'Defect'
    ]
    for column in cumulative_columns:
      net_sales_total = 0
      for metric in metrics_to_sum:
        value = output.loc[output['Metric'] == metric, column].iloc[0]
        net_sales_total += value
      output.loc[output['Metric'] == 'NET SALES', column] = net_sales_total
      qty_total = output.loc[output['Metric'] == 'QTY Total', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'NET SALES', per_unit_column] = net_sales_total / qty_total
      else:
        output.loc[output['Metric'] == 'NET SALES', per_unit_column] = 0

  def setDefaulted():
    for column in output.columns:
      if 'Cumulative' in column:
        output.loc[output['Metric'] == 'QTY Defect', column] = QTY_DEFECT
        output.loc[output['Metric'] == 'Defect %', column] = 0
        output.loc[output['Metric'] == 'Handling/Shipping', column] = SHIPPING_CUMULATIVE
        if 'All Lines' in column:
          output.loc[output['Metric'] == 'Marketing + Sponsored Ads', column] = SPONSORED_ADS_CUMULATIVE
        else:
          output.loc[output['Metric'] == 'Marketing + Sponsored Ads', column] = 0
      elif(column != 'Metric'):
        output.loc[output['Metric'] == 'Handling/Shipping', column] = SHIPPING_PER_UNIT
        output.loc[output['Metric'] == 'Marketing + Sponsored Ads', column] = SPONSORED_ADS_PER_UNIT

  def getTotalVariableCost():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    metrics_to_sum = [
      'Cost',
      'Scrap Return Rate',
      'Handling on Fab',
      'Duty',
      'Base Tariff',
      'Added Tariff',
      'Freight interco'
    ]
    for column in cumulative_columns:
      total_variable_cost = 0
      for metric in metrics_to_sum:
        value = output.loc[output['Metric'] == metric, column].iloc[0]
        total_variable_cost += value
      output.loc[output['Metric'] == 'TOTAL VARIABLE COST', column] = total_variable_cost
      qty_total = output.loc[output['Metric'] == 'QTY Total', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'TOTAL VARIABLE COST', per_unit_column] = total_variable_cost / qty_total
      else:
        output.loc[output['Metric'] == 'TOTAL VARIABLE COST', per_unit_column] = 0

  def getMargin():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      net_sales = output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
      total_variable_cost = output.loc[output['Metric'] == 'TOTAL VARIABLE COST', column].iloc[0]
      margin = net_sales - total_variable_cost
      output.loc[output['Metric'] == 'MARGIN', column] = margin
      qty_total = output.loc[output['Metric'] == 'QTY Total', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'MARGIN', per_unit_column] = margin / qty_total
      else:
        output.loc[output['Metric'] == 'MARGIN', per_unit_column] = 0
      margin_percent = (margin / net_sales) if net_sales != 0 else 0
      output.loc[output['Metric'] == 'MARGIN %', column] = margin_percent

  def getFillRateFines():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      fill_rate_fines = FILL_RATE_FINES * output.loc[output['Metric'] == 'QTY Total', column].iloc[0]
      output.loc[output['Metric'] == 'Fill Rate Fines', column] = fill_rate_fines
      qty_total = output.loc[output['Metric'] == 'QTY Total', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'Fill Rate Fines', per_unit_column] = fill_rate_fines / qty_total
      else:
        output.loc[output['Metric'] == 'Fill Rate Fines', per_unit_column] = 0

  def getInspectReturn():
    per_unit_columns = [col for col in output.columns if 'Per Unit' in col and col != 'Metric']
    for column in per_unit_columns:
      if 'All Lines' in column:
        output.loc[output['Metric'] == 'Inspect Return', column] = INPECT_RETURN_PER_UNIT
      else:
        pline = column.split(' ')[0]
        inspect_value = INSPECT_RETURN.loc['Value', pline]
        output.loc[output['Metric'] == 'Inspect Return', column] = inspect_value
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    total_all_lines = 0
    
    for column in cumulative_columns:
      if 'All Lines' in column:
        continue
      else:
        pline = column.replace(' Cumulative', '')
        returnAllowance = output.loc[output['Metric'] == 'Return Allowance', f'{pline} Cumulative'].iloc[0]
        netSalePerUnit = output.loc[output['Metric'] == 'NET SALES', f'{pline} Per Unit'].iloc[0]
        inspectReturnPerUnit = output.loc[output['Metric'] == 'Inspect Return', f'{pline} Per Unit'].iloc[0]
        cumulative_value = returnAllowance/netSalePerUnit*inspectReturnPerUnit*-1
        output.loc[output['Metric'] == 'Inspect Return', column] = cumulative_value
        total_all_lines += cumulative_value
    output.loc[output['Metric'] == 'Inspect Return', 'All Lines Cumulative'] = total_all_lines
    output.loc[output['Metric'] == 'Inspect Return', 'All Lines Bulk By Cumulative'] = total_all_lines

  def getReturnAllowancePutAwayRebox():
    total_all_lines = 0
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      if 'All Lines' in column:
        continue
      else:
        pline = column.replace(' Cumulative', '')
        output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{pline} Per Unit'] = RETURN_ALLOWANCE.loc['Per Unit', pline]
        returnAllowance = output.loc[output['Metric'] == 'Return Allowance', f'{pline} Cumulative'].iloc[0]
        netSalePerUnit = output.loc[output['Metric'] == 'NET SALES', f'{pline} Per Unit'].iloc[0]
        ReturnPerUnit = output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{pline} Per Unit'].iloc[0]
        output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{pline} Cumulative'] = returnAllowance/netSalePerUnit*-1*ReturnPerUnit*(1-SCRAP_RETURN_RATE)
        output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{pline} Per Unit'] = RETURN_ALLOWANCE.loc['Per Unit', pline]
        total_all_lines += output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', f'{pline} Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'All Lines Cumulative'] = total_all_lines
    output.loc[output['Metric'] == 'Return Allowance Put Away/Rebox', 'All Lines Bulk By Cumulative'] = total_all_lines

  def getPalletsWrapping():
    total_all_lines = 0
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col and 'All Lines' not in col]
    for column in cumulative_columns:
      pline = column.replace(' Cumulative', '')
      try :
        output.loc[output['Metric'] == 'Pallets / Wrapping', f'{pline} Per Unit'] = PALLETS_WRAPPING.loc['Cost', pline]
      except (KeyError, IndexError, ValueError) as e:
        print(f'Pallets / Wrapping for {pline} not found in Pallets Wrapping sheet. Setting to 0.')
        output.loc[output['Metric'] == 'Pallets / Wrapping', f'{pline} Per Unit'] = 0
      quantityGross = output.loc[output['Metric'] == 'QTY Gross', f'{pline} Cumulative'].iloc[0]
      output.loc[output['Metric'] == 'Pallets / Wrapping', f'{pline} Cumulative'] = quantityGross*output.loc[output['Metric'] == 'Pallets / Wrapping', f'{pline} Per Unit'].iloc[0]
      total_all_lines += output.loc[output['Metric'] == 'Pallets / Wrapping', f'{pline} Cumulative'].iloc[0]

    output.loc[output['Metric'] == 'Pallets / Wrapping', 'All Lines Cumulative'] = total_all_lines
    output.loc[output['Metric'] == 'Pallets / Wrapping', 'All Lines Bulk By Cumulative'] = total_all_lines
    output.loc[output['Metric'] == 'Pallets / Wrapping', 'All Lines Per Unit'] = total_all_lines/output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Pallets / Wrapping', 'All Lines Bulk By Per Unit'] = total_all_lines/output.loc[output['Metric'] == 'QTY Gross', 'All Lines Bulk By Cumulative'].iloc[0]

  def getDelivery():
    total_all_lines = 0
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col and 'All Lines' not in col]
    for column in cumulative_columns:
      pline = column.replace(' Cumulative', '')
      output.loc[output['Metric'] == 'Delivery', f'{pline} Per Unit'] = DELIVERY_PER_UNIT
      quantityGross = output.loc[output['Metric'] == 'QTY Gross', f'{pline} Cumulative'].iloc[0]
      output.loc[output['Metric'] == 'Delivery', f'{pline} Cumulative'] = quantityGross*output.loc[output['Metric'] == 'Delivery', f'{pline} Per Unit'].iloc[0]
      total_all_lines += output.loc[output['Metric'] == 'Delivery', f'{pline} Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Delivery', 'All Lines Cumulative'] = total_all_lines
    output.loc[output['Metric'] == 'Delivery', 'All Lines Bulk By Cumulative'] = total_all_lines
    output.loc[output['Metric'] == 'Delivery', 'All Lines Per Unit'] = DELIVERY_PER_UNIT
    output.loc[output['Metric'] == 'Delivery', 'All Lines Bulk By Per Unit'] = DELIVERY_PER_UNIT

  def getSGandA():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    metrics_to_sum = [
      'Handling/Shipping',
      'Fill Rate Fines',
      'Inspect Return', 
      'Return Allowance Put Away/Rebox',
      'Pallets / Wrapping',
      'Delivery',
      'Marketing + Sponsored Ads'
    ]
    for column in cumulative_columns:
      SGA = 0.0
      for metric in metrics_to_sum:
        value = output.loc[output['Metric'] == metric, column].iloc[0]
        SGA += value
      
      output.loc[output['Metric'] == 'SG&A', column] = SGA

      qty_total = output.loc[output['Metric'] == 'QTY Total', 'All Lines Cumulative'].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      
      if qty_total > 0:
        output.loc[output['Metric'] == 'SG&A', per_unit_column] = SGA / qty_total
      else:
        output.loc[output['Metric'] == 'SG&A', per_unit_column] = 0

  def getFactoring():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      factoring_value = FACTORING * output.loc[output['Metric'] == 'NET SALES', column].iloc[0]
      output.loc[output['Metric'] == 'FACTORING %', column] = factoring_value
      column = column.replace('Cumulative', 'Per Unit')
      output.loc[output['Metric'] == 'FACTORING %', column] = factoring_value / output.loc[output['Metric'] == 'QTY Total', column].iloc[0] if output.loc[output['Metric'] == 'QTY Total', column].iloc[0] > 0 else 0

  def getContributionMargin():
    cumulative_columns = [col for col in output.columns if 'Cumulative' in col]
    for column in cumulative_columns:
      contribution_margin = output.loc[output['Metric'] == 'MARGIN', column].iloc[0] - output.loc[output['Metric'] == 'FACTORING %', column].iloc[0] - output.loc[output['Metric'] == 'SG&A', column].iloc[0]
      output.loc[output['Metric'] == 'Contribution Margin', column] = contribution_margin
      qty_total = output.loc[output['Metric'] == 'QTY Total', column].iloc[0]
      per_unit_column = column.replace('Cumulative', 'Per Unit')
      if qty_total > 0:
        output.loc[output['Metric'] == 'Contribution Margin', per_unit_column] = contribution_margin / qty_total
      else:
        output.loc[output['Metric'] == 'Contribution Margin', per_unit_column] = 0
      contribution_margin_percent = (contribution_margin / output.loc[output['Metric'] == 'NET SALES', column].iloc[0]) if output.loc[output['Metric'] == 'NET SALES', column].iloc[0] != 0 else 0
      output.loc[output['Metric'] == 'Contribution Margin %', column] = contribution_margin_percent

  def getTotalScrapReturnRate():
    all_rep_commission = REP_COMMISSION * (DATA_V2.loc[:, 'Total Sales']-DATA_V2.loc[:, 'Defect']-DATA_V2.loc[:, 'Rebate']-DATA_V2.loc[:, 'Payment']-DATA_V2.loc[:, 'Freight Allowance']-DATA_V2.loc[:, 'Return Allowance']-DATA_V2.loc[:, 'Defect Rebate in Lieu of Returns'])
    all_net_sales = DATA_V2.loc[:, 'Total Sales'] - DATA_V2.loc[:, 'Defect'] - DATA_V2.loc[:, 'Rebate'] - DATA_V2.loc[:, 'Payment'] - DATA_V2.loc[:, 'Freight Allowance'] - DATA_V2.loc[:, 'Return Allowance'] - DATA_V2.loc[:, 'Defect Rebate in Lieu of Returns'] - all_rep_commission
    scrap_return = (1-SCRAP_RETURN_RATE)*(DATA_V2.loc[:, 'Return Allowance']/(all_net_sales/DATA_V2.loc[:, 'L12 Shipped']))*(DATA_V2.loc[:, 'Total PO Cost'] + DATA_V2.loc[:, 'Total Fab Handling']+ DATA_V2.loc[:, 'Total Duty'] + DATA_V2.loc[:, 'Total Tariff'] + DATA_V2.loc[:, "Add'l Tariff Total"])/DATA_V2.loc[:, 'L12 Shipped']
    return scrap_return.sum()*-1

  print('Reading AMZ DATA sheet...')
  DATA_V2 = pd.read_excel(file, sheet_name='DATA V2', header=1)
  print('Reading Assumptions...')
  if user_defaults_df is not None:
    ASSUMPTIONS = pd.DataFrame({'Values': user_defaults_df}).T
  else:
    ASSUMPTIONS = pd.read_excel(file, sheet_name='Defaults & Assumptions', header=4)
    print("Default Assumptions not provided, using file defaults.")
    ASSUMPTIONS = ASSUMPTIONS.iloc[:, 1:]
  print('Reading Defaults...')
  if user_defaults_df is not None:
    DEFAULTS = pd.DataFrame({'Values': user_defaults_df}).T
  else:
    DEFAULTS = pd.read_excel(file, sheet_name='Defaults & Assumptions', nrows=2)
    DEFAULTS = DEFAULTS.iloc[:, 1:]
  print('Reading Inspect Return Sheet...')
  INSPECT_RETURN = pd.read_excel(file, sheet_name='Inspect Return')
  print('Reading Return Allowance Sheet...')
  RETURN_ALLOWANCE = pd.read_excel(file, sheet_name='Return Allowance Put Away Rebox')
  print('Reading Pallet Wrap Sheet...')
  PALLETS_WRAPPING = pd.read_excel(file, sheet_name='Pallet Wrap', header=1)

  INSPECT_RETURN = INSPECT_RETURN.set_index('Plines').T
  RETURN_ALLOWANCE = RETURN_ALLOWANCE.set_index('Plines').T
  PALLETS_WRAPPING = PALLETS_WRAPPING.set_index('P Line').T

  DELIVERY_PER_UNIT = ASSUMPTIONS['Delivery Per Unit'][0]
  INPECT_RETURN_PER_UNIT = ASSUMPTIONS['Inspect Return Per Unit'][0]
  QTY_DEFECT = ASSUMPTIONS['QTY Defect'][0]
  CUMULLATIVE_DEFECT = ASSUMPTIONS['Cumulative Defect'][0]
  SHIPPING_PER_UNIT = ASSUMPTIONS['Handling/Shipping Per Unit'][0]
  SHIPPING_CUMULATIVE = ASSUMPTIONS['Handling/Shipping Cumulative'][0]
  FILL_RATE_FINES = DEFAULTS['Fill Rate Fines'][0]
  SCRAP_RETURN_RATE = DEFAULTS['Scrap Return Rate'][0]
  SPONSORED_ADS_CUMULATIVE = ASSUMPTIONS['All Lines Marketing + Sponsored Ads Cumulative'][0]
  SPONSORED_ADS_PER_UNIT = ASSUMPTIONS['All Lines Marketing + Sponsored Ads Per Unit'][0]
  FACTORING = DEFAULTS['FACTORING %'][0]
  REP_COMMISSION = DEFAULTS['Rep Commission'][0]
  FREIGHT_INTERCO = 0

  print('Preparing output structure...')
  METRICS = [
    'QTY Gross',
    'QTY Defect',
    'QTY Total',
    'Defect %',
    'Sales',
    'Rebate (10%) + Payment (2%)',
    'Freight Allowance',
    'Agency Rep (RPS)',
    'Return Allowance',
    'Defect Rebate in Lieu of Returns',
    'Defect',
    'NET SALES',
    'Cost',
    'Scrap Return Rate',
    'Handling on Fab',
    'Duty',
    'Base Tariff',
    'Added Tariff',
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
    'Marketing + Sponsored Ads',
    'SG&A',
    'FACTORING %',
    'Contribution Margin',
    'Contribution Margin %'
  ]
  COLUMNS = [
    'Metric',
    'All Lines Cumulative',
    'All Lines Per Unit',
    'All Lines Bulk By Cumulative',
    'All Lines Bulk By Per Unit'
  ]
  Pline = DATA_V2['Pline'].unique()

  for line in Pline:
    COLUMNS.append(f'{line} Cumulative')
    COLUMNS.append(f'{line} Per Unit')

  output = pd.DataFrame(columns= COLUMNS)
  output['Metric'] = METRICS

  setDefaulted()

  print('Calculating QTY values...')
  gross = util.getSumGivenColumn('All', DATA_V2, 'L12 Shipped')
  gross_all_lines = util.getSumGivenColumn('All', DATA_V2, 'L12 Shipped')
  print(gross, gross + (gross * volume))
  output.loc[output['Metric'] == 'QTY Gross', 'All Lines Cumulative'] = gross + (gross * volume)
  output.loc[output['Metric'] == 'QTY Gross', 'All Lines Bulk By Cumulative'] = gross_all_lines + (gross_all_lines * volume)
  for line in Pline:
    gross_pline = util.getSumGivenColumn(line, DATA_V2, 'L12 Shipped')
    output.loc[output['Metric'] == 'QTY Gross', f'{line} Cumulative'] = gross_pline + (gross_pline * volume)
  output = util.getQTYTotalAndDefect(output)

  print('Calculating Sales values...')
  output.loc[output['Metric'] == 'Sales', 'All Lines Cumulative'] = getSales('All')
  output.loc[output['Metric'] == 'Sales', 'All Lines Bulk By Cumulative'] = getSales('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Sales', f'{line} Cumulative'] = getSales(line)

  print('Calculating Rebate + Payment values...')
  output.loc[output['Metric'] == 'Rebate (10%) + Payment (2%)', 'All Lines Cumulative'] = getRebatePlusAllowance('All')
  output.loc[output['Metric'] == 'Rebate (10%) + Payment (2%)', 'All Lines Bulk By Cumulative'] = getRebatePlusAllowance('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Rebate (10%) + Payment (2%)', f'{line} Cumulative'] = getRebatePlusAllowance(line)

  print('Calculating Freight Allowance values...')
  output.loc[output['Metric'] == 'Freight Allowance', 'All Lines Cumulative'] = getFreightAllowance('All')
  output.loc[output['Metric'] == 'Freight Allowance', 'All Lines Bulk By Cumulative'] = getFreightAllowance('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Freight Allowance', f'{line} Cumulative'] = getFreightAllowance(line)

  print('Calculating Agency Rep values...')
  output.loc[output['Metric'] == 'Agency Rep (RPS)', 'All Lines Cumulative'] = getAgencyRep('All')
  output.loc[output['Metric'] == 'Agency Rep (RPS)', 'All Lines Bulk By Cumulative'] = getAgencyRep('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Agency Rep (RPS)', f'{line} Cumulative'] = getAgencyRep(line)

  print('Calculating Return Allowance values...')
  output.loc[output['Metric'] == 'Return Allowance', 'All Lines Cumulative'] = getReturnAllowance('All')
  output.loc[output['Metric'] == 'Return Allowance', 'All Lines Bulk By Cumulative'] = getReturnAllowance('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Return Allowance', f'{line} Cumulative'] = getReturnAllowance(line)

  print('Calculating Defect Rebate values...')
  output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', 'All Lines Cumulative'] = getDefectRebate('All')
  output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', 'All Lines Bulk By Cumulative'] = getDefectRebate('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Defect Rebate in Lieu of Returns', f'{line} Cumulative'] = getDefectRebate(line)

  print('Calculating Defect values...')
  output.loc[output['Metric'] == 'Defect', 'All Lines Cumulative'] = getDefect('All')
  output.loc[output['Metric'] == 'Defect', 'All Lines Bulk By Cumulative'] = getDefect('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Defect', f'{line} Cumulative'] = getDefect(line)

  print('Calculating NET Sales values...')
  getNETSales()

  print('Calculating Cost values...')
  output.loc[output['Metric'] == 'Cost', 'All Lines Cumulative'] = getCost('All')
  output.loc[output['Metric'] == 'Cost', 'All Lines Bulk By Cumulative'] = getCost('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Cost', f'{line} Cumulative'] = getCost(line)

  print('Calculating Scrap Return Rate values...')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'All Lines Cumulative'] = getScrapReturnRate('All')
  output.loc[output['Metric'] == 'Scrap Return Rate', 'All Lines Bulk By Cumulative'] = getScrapReturnRate('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Scrap Return Rate', f'{line} Cumulative'] = getScrapReturnRate(line)

  print('Calculating Handling on Fab values...')
  output.loc[output['Metric'] == 'Handling on Fab', 'All Lines Cumulative'] = getHandlingOnFab('All')
  output.loc[output['Metric'] == 'Handling on Fab', 'All Lines Bulk By Cumulative'] = getHandlingOnFab('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Handling on Fab', f'{line} Cumulative'] = getHandlingOnFab(line)

  print('Calculating Duty values...')
  output.loc[output['Metric'] == 'Duty', 'All Lines Cumulative'] = getDuty('All')
  output.loc[output['Metric'] == 'Duty', 'All Lines Bulk By Cumulative'] = getDuty('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Duty', f'{line} Cumulative'] = getDuty(line)

  print('Calculating Base Tariff values...')
  output.loc[output['Metric'] == 'Base Tariff', 'All Lines Cumulative'] = getBaseTariff('All')
  output.loc[output['Metric'] == 'Base Tariff', 'All Lines Bulk By Cumulative'] = getBaseTariff('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Base Tariff', f'{line} Cumulative'] = getBaseTariff(line)

  print('Calculating Added Tariff values...')
  output.loc[output['Metric'] == 'Added Tariff', 'All Lines Cumulative'] = getAddedTariff('All')
  output.loc[output['Metric'] == 'Added Tariff', 'All Lines Bulk By Cumulative'] = getAddedTariff('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Added Tariff', f'{line} Cumulative'] = getAddedTariff(line)

  print('Calculating Freight Interco values...')
  output.loc[output['Metric'] == 'Freight interco', 'All Lines Cumulative'] = getFreightInterco('All')
  output.loc[output['Metric'] == 'Freight interco', 'All Lines Bulk By Cumulative'] = getFreightInterco('All', bulk=True)
  for line in Pline:
    output.loc[output['Metric'] == 'Freight interco', f'{line} Cumulative'] = getFreightInterco(line)

  print('Calculating Total Variable Cost values...')
  getTotalVariableCost()

  print('Calculating Margin values...')
  getMargin()

  print('Calculating Fill Rate Fines values...')
  getFillRateFines()

  print('Calculating Inspect Return values...')
  getInspectReturn()

  print('Calculating Return Allowance Put Away/Rebox values...')
  getReturnAllowancePutAwayRebox()

  print('Calculating Pallets / Wrapping values...')
  getPalletsWrapping()

  print('Calculating Delivery values...')
  getDelivery()

  print('Calculating SG&A values...')
  getSGandA()

  print('Calculating Factoring values...')
  getFactoring()

  print('Calculating Contribution Margin values...')
  getContributionMargin()

  assumptions_list = {}
  for column in ASSUMPTIONS.columns:
    if 'Unnamed' in column:
      continue
    for index in ASSUMPTIONS.index:
      parameter_name = f"{column}"
      value = ASSUMPTIONS.loc[index, column]
      assumptions_list[parameter_name] = round(float(value),4)
  
  defaults_list = {}
  for column in DEFAULTS.columns:
    if 'Unnamed' in column:
      continue
    for index in DEFAULTS.index:
      parameter_name = f"{column}"
      value = DEFAULTS.loc[index, column]
      defaults_list[parameter_name] = round(float(value),4)
  
  return output, assumptions_list, defaults_list