import pandas as pd
import utility_library as util

def getSummary(file, user_defaults_df=None):

  def getPerUnit(row, column):
    return output.loc[output['Metric'] == row, column].iloc[0] / output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]

  METRICS = [
    'QTY Gross',
    'QTY Defect',
    'QTY Total',
    'Defect %',

    'Sales',

    'Current',
    'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)',
    'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)',
    'Volume rebate 20% (CMP, CKP, MAF, IGN)',
    'Exchange Rate Program 5% COM',
    'Loyalty Rebate 1% ALL',
    'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)',
    'Volume Rebate Additional 7.5% (EVA, ACC, CO, CON, CGT, SUC, COM, CMP, CKP, MAP, MAF, OXS, ETB, IGN, VTP, VVT, FP, FPA, FPD, SF, MFP, FIP, FM, SU, CHD, CIN, CIA, CAC, INT, EMC, FC, BPA, OPN, TPN, FNH, LO, GTP, FNA, FN, FRA, ST, GTN, DFC)',
    'Marketing Rebate 7.5% (FP, SF, MFP, FM)',
    'Enhancement Rebate 8.3% (FP, SF, MFP, FM, FTA)',
    'Fuel Delivery Marketing 5% (FP, FPA, SF, MFP, FM)',
    'Heater Loyalty Rebate 9% (HEA, HDH)',
    'Alliance Volume Rebate 8% ALL',
    'Alliance Fuel Delivery 6% (SUC, FP, FPA, FPD, SF, MFP, FIP, FM, SU)',

    'New',
    'Support Rebate 1% ALL',
    'Volume Rebate 1.5% (CON, CHD)',
    'Business Development Fund 2% (CMP, CKP, ETB, IGN, MAP, MAF, OXS, VTP, VVT)',
    'Off Invoice 0.5% VTP',
    'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)',
    'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)',
    'Off Invoice 6% CHD',
    'Off Invoice 6.25% (CIA, CAC, CIN)',
    'Off Invoice 6.875% (CON, COM, INT)',
    'Off Invoice 7.125% (HEA, HDH)',
    'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)',
    'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)',
    'Off Invoice 20.125% (OPN, TPN, DFC)',
    'Off Invoice 21.125% ST',
    'Off Invoice 26.125% SU',
    'Off Invoice 26.625% (FN, FNA, FNH)',
    'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)',
    '-',
    'Agency fees',
    'Return Allowance',
    'Defect rate',
    'NET SALES',

    'FG PO cost',
    'Variable - Overhead',
    'Labor',
    'Duty',
    'Scrap Return % (Resellable %)',
    'Tariffs',
    'FG Freight',
    'Freight interco',
    'TOTAL VARIABLE COST',
    'MARGIN',
    'MARGIN %',

    'Handling/Shipping',
    'Fill rate fines',
    'Inspect of Return',
    'Return Allowance Put Away Costs, for usable product',
    'Special Marketing (we control)',
    'Pallets / Wrapping',
    'Delivery',
    'SG&A',

    'Payment term',
    'Contribution Margin $s',
    'Contribution Margin %'
  ]

  excel_file = pd.ExcelFile(file)
  plines = excel_file.sheet_names[3:]
  PALLETS_AND_WRAPPING = pd.read_excel(file, sheet_name='Pallets&Wrapping 202402')
  BOX = pd.read_excel(file, sheet_name='Box', header=5)

  COLUMNS = [
    'Metric',
    'Total Vast Current Cumulative',
    'Total Vast Current Per Unit',
    'Total Vast New Cumulative',
    'Total Vast New Per Unit'
  ]

  for x in plines:
    COLUMNS.append(f'{x} Current Vast Cumulative')
    COLUMNS.append(f'{x} Current Vast Per Unit')
    COLUMNS.append(f'{x} New Vast Cumulative')
    COLUMNS.append(f'{x} New Vast Per Unit')

  DEFECT = {
    'ACC': 0.01,
    'CGT': 0.03,
    'CKP': 0.015,
    'CMP': 0.015,
    'CO': 0.025,
    'COM': 0.05,
    'CON': 0.02,
    'EMC': 0.12,
    'ETB': 0.15,
    'EVA': 0.03,
    'FIP': 0.22,
    'FM': 0.1,
    'FN': 0.035,
    'FNH': 0.02,
    'FP': 0.055,
    'FPD': 0.16,
    'GTN': 0.045,
    'HDH': 0.015,
    'HEA': 0.015,
    'IGN': 0.05,
    'INT': 0.06,
    'LO': 0.01,
    'MAF': 0.1,
    'MAP': 0.01,
    'MFP': 0.025,
    'OPN': 0.03,
    'SF': 0.005,
    'ST': 0.02,
    'SU': 0.015,
    'SUC': 0.04,
    'TPN': 0.005,
  }

  output = pd.DataFrame(columns=COLUMNS)
  output['Metric'] = METRICS

  print('Calculating QTY Gross...')
  total = 0
  for pline in plines:
    data = pd.read_excel(file, sheet_name=pline, header=1)
    value = data['Vast volume'].sum()
    output.loc[output['Metric'] == 'QTY Gross', f'{pline} Current Vast Cumulative'] = value
    output.loc[output['Metric'] == 'QTY Gross', f'{pline} New Vast Cumulative'] = value
    total += value
  output.loc[output['Metric'] == 'QTY Gross', 'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'QTY Gross', 'Total Vast New Cumulative'] = total

  print('Calculating Defect % ...')
  for pline in plines:
    defect_rate = DEFECT.get(pline)
    output.loc[output['Metric'] == 'Defect %', f'{pline} Current Vast Cumulative'] = defect_rate
    output.loc[output['Metric'] == 'Defect %', f'{pline} New Vast Cumulative'] = defect_rate
  output.loc[output['Metric'] == 'Defect %', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Defect %', 'Total Vast New Cumulative'] = 0

  print('Calculating QTY Defect...')
  total = 0
  for pline in plines:
    qty_gross = output.loc[output['Metric'] == 'QTY Gross', f'{pline} Current Vast Cumulative'].values[0]
    defect_rate = output.loc[output['Metric'] == 'Defect %', f'{pline} Current Vast Cumulative'].values[0]
    value = qty_gross * defect_rate * -1
    output.loc[output['Metric'] == 'QTY Defect', f'{pline} Current Vast Cumulative'] = value
    output.loc[output['Metric'] == 'QTY Defect', f'{pline} New Vast Cumulative'] = value
    total += value
  output.loc[output['Metric'] == 'QTY Defect', 'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'QTY Defect', 'Total Vast New Cumulative'] = total

  print('Calculating QTY Total...')
  total = 0
  for pline in plines:
    qty_gross = output.loc[output['Metric'] == 'QTY Gross', f'{pline} Current Vast Cumulative'].values[0]
    qty_defect = output.loc[output['Metric'] == 'QTY Defect', f'{pline} Current Vast Cumulative'].values[0]
    value = qty_gross + qty_defect
    output.loc[output['Metric'] == 'QTY Total', f'{pline} Current Vast Cumulative'] = value
    output.loc[output['Metric'] == 'QTY Total', f'{pline} New Vast Cumulative'] = value
    total += value
  output.loc[output['Metric'] == 'QTY Total', 'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'QTY Total', 'Total Vast New Cumulative'] = total

  print('Calculating Sales...')
  total_cumulative = 0
  total_per_unit = 0
  for pline in plines:
    data = pd.read_excel(file, sheet_name=pline, header=1)
    current_sales = (data['Current invoice']*data['Vast volume']).sum()
    new_sales = (data['NEW Invoice']*data['Vast volume']).sum()
    output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'] = current_sales
    output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'] = new_sales
    total_cumulative += current_sales
    total_per_unit += new_sales
  output.loc[output['Metric'] == 'Sales', 'Total Vast Current Cumulative'] = total_cumulative
  output.loc[output['Metric'] == 'Sales', 'Total Vast New Cumulative'] = total_per_unit
  for columns in output.columns:
    if 'Cumulative' in columns:
      if 'Total' in columns:
        column = 'TPN Current Vast Cumulative'
        output.loc[output['Metric'] == 'Sales', columns.replace('Cumulative', 'Per Unit')] = total_cumulative / output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
      elif 'Totoal' in columns:
        column = 'TPN New Vast Cumulative'
        output.loc[output['Metric'] == 'Sales', columns.replace('Cumulative', 'Per Unit')] = total_per_unit / output.loc[output['Metric'] == 'QTY Gross', column].iloc[0]
      else:
        output.loc[output['Metric'] == 'Sales', columns.replace('Cumulative', 'Per Unit')] = getPerUnit('Sales', columns)

  print('Calculating Volume Rabates...')
  total_14 = 0
  total_5 = 0
  total_20 = 0
  for pline in plines:
    if pline in 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)':
      if pline == 'OPN' or pline == 'TPN':
        output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Per Unit'] = 0
        output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} New Vast Cumulative'] = 0
        output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Per Unit'] = 0
        output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} New Vast Cumulative'] = 0
        output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = 0
        output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} New Vast Cumulative'] = 0
        output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} New Vast Per Unit'] = 0
        continue
      value = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].values[0] * -0.14
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'] = value
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Per Unit'] = getPerUnit('Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = 0
      total_14 += value
    elif pline in 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)':
      value = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].values[0] * -0.05
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'] = value
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Per Unit'] = getPerUnit('Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = 0
      total_5 += value
    elif pline in 'Volume rebate 20% (CMP, CKP, MAF, IGN)':
      value = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].values[0] * -0.2
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = value
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = getPerUnit('Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative')
      total_20 += value
    else:
      output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = 0
    output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} New Vast Cumulative'] = 0
    output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} New Vast Per Unit'] = 0
    output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} New Vast Cumulative'] = 0
    output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} New Vast Per Unit'] = 0
    output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} New Vast Cumulative'] = 0
    output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} New Vast Per Unit'] = 0
  output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', 'Total Vast New Cumulative'] = 0
  output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', 'Total Vast New Cumulative'] = 0
  output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', 'Total Vast New Cumulative'] = 0
  output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', 'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', 'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', 'Total Vast New Per Unit'] = 0
  divisor = output.loc[output['Metric'] == 'QTY Gross', 'TPN Current Vast Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', 'Total Vast Current Cumulative'] = total_14
  output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', 'Total Vast Current Per Unit'] = total_14/divisor
  output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', 'Total Vast Current Per Unit'] = total_5/divisor
  output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', 'Total Vast Current Cumulative'] = total_5
  output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', 'Total Vast Current Cumulative'] = total_20
  output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', 'Total Vast Current Per Unit'] = total_20/divisor

  print('Calculating Exchange Rate Program...')
  total = 0
  for pline in plines:
    if pline == 'COM':
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      value = sales * -0.05
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} Current Vast Cumulative'] = value
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} Current Vast Per Unit'] = getPerUnit('Exchange Rate Program 5% COM', f'{pline} New Vast Cumulative')
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} New Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} New Vast Per Unit'] = 0
      total += value
    else:
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} New Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', f'{pline} New Vast Per Unit'] = 0
  output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', 'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', 'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', 'Total Vast New Cumulative'] = 0
  output.loc[output['Metric'] == 'Exchange Rate Program 5% COM', 'Total Vast New Per Unit'] = 0

  print('Calculating Loyalty Revate...')
  total_1 = 0
  total_5 = 0
  for pline in plines:
    sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
    if pline == 'COM' or pline == 'CMP' or pline == 'CKP' or pline == 'MAF' or pline == 'IGN':
      output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = sales * -0.05
      output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = getPerUnit('Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative')
      total_5 += sales * -0.05
    else:
      output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} Current Vast Per Unit'] = 0
    output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} New Vast Per Unit'] = 0
    output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', f'{pline} New Vast Cumulative'] = 0
    output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', f'{pline} New Vast Per Unit'] = 0
    output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', f'{pline} New Vast Cumulative'] = 0
    output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', f'{pline} Current Vast Cumulative'] = sales * -0.01
    output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', f'{pline} Current Vast Per Unit'] = getPerUnit('Loyalty Rebate 1% ALL', f'{pline} Current Vast Cumulative')
    total_1 += sales * -0.01
  output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', 'Total Vast Current Cumulative'] = total_5
  output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', 'Total Vast Current Per Unit'] = total_5/divisor
  output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', 'Total Vast Current Cumulative'] = total_1
  output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', 'Total Vast Current Per Unit'] = total_1/divisor
  output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', 'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL', 'Total Vast New Cumulative'] = 0
  output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', 'Total Vast New Cumulative'] = 0
  output.loc[output['Metric'] == 'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)', 'Total Vast New Per Unit'] = 0

  print('Calculating Volume Rebate Additional...')
  total = 0
  for pline in plines:
    row = 'Volume Rebate Additional 7.5% (EVA, ACC, CO, CON, CGT, SUC, COM, CMP, CKP, MAP, MAF, OXS, ETB, IGN, VTP, VVT, FP, FPA, FPD, SF, MFP, FIP, FM, SU, CHD, CIN, CIA, CAC, INT, EMC, FC, BPA, OPN, TPN, FNH, LO, GTP, FNA, FN, FRA, ST, GTN, DFC)'
    if pline in row:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      value = sales * -0.075
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = value
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
      total += value
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
    output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Marketing Rebate...')
  total = 0
  for pline in plines:
    row = 'Marketing Rebate 7.5% (FP, SF, MFP, FM)'
    if pline in row:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = sales * -0.075
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      total += sales * -0.075
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Enhancement Rebate...')
  total = 0
  for pline in plines:
    row = 'Enhancement Rebate 8.3% (FP, SF, MFP, FM, FTA)'
    if pline in row:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = sales * -0.083
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      total += sales * -0.083
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Fuel Delivery Marketing...')
  total = 0
  for pline in plines:
    row = 'Fuel Delivery Marketing 5% (FP, FPA, SF, MFP, FM)'
    if pline in row:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = sales * -0.05
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      total += sales * -0.05
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Heater Loyalty Rebate...')
  total = 0
  for pline in plines:
    row = 'Heater Loyalty Rebate 9% (HEA, HDH)'
    if pline in row:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = sales * -0.09
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      total += sales * -0.09
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Alliance Volume Rebate')
  total = 0
  row = 'Alliance Volume Rebate 8% ALL'
  for pline in plines:
    sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = sales * -0.08
    output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
    output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
    total += sales * -0.08
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Alliance Fuel Delivery...')
  total = 0
  row = 'Alliance Fuel Delivery 6% (SUC, FP, FPA, FPD, SF, MFP, FIP, FM, SU)'
  for pline in plines:
    if pline in row:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = sales * -0.06
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = getPerUnit(row, f'{pline} Current Vast Cumulative')
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      total += sales * -0.06
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = 0

  print('Calculating Support Rebate...')
  total = 0
  row = 'Support Rebate 1% ALL'
  for pline in plines:
    sales = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
    output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.01
    output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
    total += sales * -0.01
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = total

  print('Calculating Volume Rebate...')
  total = 0
  row = 'Volume Rebate 1.5% (CON, CHD)'
  for pline in plines:
    if pline == 'CON' or pline == 'CHD':
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.015
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
      total += sales * -0.015
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = total/divisor
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = total

  print('Calculating Business Development Fund...')
  total = 0
  row = 'Business Development Fund 2% (CMP, CKP, ETB, IGN, MAP, MAF, OXS, VTP, VVT)'
  for pline in plines:
    if pline in ['CMP', 'CKP', 'ETB', 'IGN', 'MAP', 'MAF', 'OXS', 'VTP', 'VVT']:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.02
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
      total += sales * -0.02
    else:
      output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
      output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == row, f'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == row, f'Total Vast New Per Unit'] = total/divisor 
  output.loc[output['Metric'] == row, f'Total Vast New Cumulative'] = total

  print('Calculating Off Invoice...')
  total_p5 = 0
  total_p625 = 0
  total_1p5 = 0
  total_6 = 0
  total_6p25 = 0
  total_6p875 = 0
  total_7p125 = 0
  total_14p125 = 0
  total_19p125 = 0
  total_20p125 = 0
  total_21p125 = 0
  total_26p125 = 0
  total_26p625 = 0
  total_28 = 0
  rows = [
    'Off Invoice 0.5% VTP',
    'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)',
    'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)',
    'Off Invoice 6% CHD',
    'Off Invoice 6.25% (CIA, CAC, CIN)',
    'Off Invoice 6.875% (CON, COM, INT)',
    'Off Invoice 7.125% (HEA, HDH)',
    'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)',
    'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)',
    'Off Invoice 20.125% (OPN, TPN, DFC)',
    'Off Invoice 21.125% ST',
    'Off Invoice 26.125% SU',
    'Off Invoice 26.625% (FN, FNA, FNH)',
    'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)'
  ]
  for pline in plines:
    sales = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
    if pline in ['VTP']:
      for row in rows:
        if row == 'Off Invoice 0.5% VTP':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.005
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_p5 += sales * -0.005
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['MAP', 'FPD', 'BPA', 'EMC', 'FC']:
      for row in rows:
        if row == 'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.00625
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_p625 += sales * -0.00625
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['ACC', 'CO', 'EVA', 'OXS', 'GTP']:
      for row in rows:
        if row == 'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.015
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_1p5 += sales * -0.015
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['CHD']:
      for row in rows:
        if row == 'Off Invoice 6% CHD':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.06
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_6 += sales * -0.06
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['CIA', 'CAC', 'CIN']:
      for row in rows:
        if row == 'Off Invoice 6.25% (CIA, CAC, CIN)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.0625
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_6p25 += sales * -0.0625
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['CON', 'COM', 'INT']:
      for row in rows:
        if row == 'Off Invoice 6.875% (CON, COM, INT)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.06875
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_6p875 += sales * -0.06875
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['HEA', 'HDH']:
      for row in rows:
        if row == 'Off Invoice 7.125% (HEA, HDH)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.07125
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_7p125 += sales * -0.07125
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['SUC', 'FTA', 'LO', 'GTN', 'CGT']:
      for row in rows:
        if row == 'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.14125
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_14p125 += sales * -0.14125
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['VVT', 'FP', 'FIP', 'FPA', 'FM', 'SF', 'MFP']:
      for row in rows:
        if row == 'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.19125
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_19p125 += sales * -0.19125
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['OPN', 'TPN', 'DFC']:
      for row in rows:
        if row == 'Off Invoice 20.125% (OPN, TPN, DFC)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.20125
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_20p125 += sales * -0.20125
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['ST']:
      for row in rows:
        if row == 'Off Invoice 21.125% ST':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.21125
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_21p125 += sales * -0.21125
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['SU']:
      for row in rows:
        if row == 'Off Invoice 26.125% SU':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.26125
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_26p125 += sales * -0.26125
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['FN', 'FNA', 'FNH']:
      for row in rows:
        if row == 'Off Invoice 26.625% (FN, FNA, FNH)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.26625
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_26p625 += sales * -0.26625
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    elif pline in ['CMP', 'CKP', 'IGN', 'MAF']:
      for row in rows:
        if row == 'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)':
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = sales * -0.28
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = getPerUnit(row, f'{pline} New Vast Cumulative')
          total_28 += sales * -0.28
        else:
          output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
          output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
    else:
      for row in rows:
        output.loc[output['Metric'] == row, f'{pline} Current Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} Current Vast Per Unit'] = 0
        output.loc[output['Metric'] == row, f'{pline} New Vast Cumulative'] = 0
        output.loc[output['Metric'] == row, f'{pline} New Vast Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 0.5% VTP', 'Total Vast New Cumulative'] = total_p5
  output.loc[output['Metric'] == 'Off Invoice 0.5% VTP', 'Total Vast New Per Unit'] = total_p5/divisor
  output.loc[output['Metric'] == 'Off Invoice 0.5% VTP', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 0.5% VTP', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)', 'Total Vast New Cumulative'] = total_p625
  output.loc[output['Metric'] == 'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)', 'Total Vast New Per Unit'] = total_p625/divisor
  output.loc[output['Metric'] == 'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)', 'Total Vast New Cumulative'] = total_1p5
  output.loc[output['Metric'] == 'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)', 'Total Vast New Per Unit'] = total_1p5/divisor
  output.loc[output['Metric'] == 'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 6% CHD', 'Total Vast New Cumulative'] = total_6
  output.loc[output['Metric'] == 'Off Invoice 6% CHD', 'Total Vast New Per Unit'] = total_6/divisor
  output.loc[output['Metric'] == 'Off Invoice 6% CHD', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 6% CHD', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 6.25% (CIA, CAC, CIN)', 'Total Vast New Cumulative'] = total_6p25
  output.loc[output['Metric'] == 'Off Invoice 6.25% (CIA, CAC, CIN)', 'Total Vast New Per Unit'] = total_6p25/divisor
  output.loc[output['Metric'] == 'Off Invoice 6.25% (CIA, CAC, CIN)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 6.25% (CIA, CAC, CIN)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 6.875% (CON, COM, INT)', 'Total Vast New Cumulative'] = total_6p875
  output.loc[output['Metric'] == 'Off Invoice 6.875% (CON, COM, INT)', 'Total Vast New Per Unit'] = total_6p875/divisor
  output.loc[output['Metric'] == 'Off Invoice 6.875% (CON, COM, INT)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 6.875% (CON, COM, INT)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 7.125% (HEA, HDH)', 'Total Vast New Cumulative'] = total_7p125
  output.loc[output['Metric'] == 'Off Invoice 7.125% (HEA, HDH)', 'Total Vast New Per Unit'] = total_7p125/divisor
  output.loc[output['Metric'] == 'Off Invoice 7.125% (HEA, HDH)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 7.125% (HEA, HDH)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)', 'Total Vast New Cumulative'] = total_14p125
  output.loc[output['Metric'] == 'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)', 'Total Vast New Per Unit'] = total_14p125/divisor
  output.loc[output['Metric'] == 'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)', 'Total Vast New Cumulative'] = total_19p125
  output.loc[output['Metric'] == 'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)', 'Total Vast New Per Unit'] = total_19p125/divisor
  output.loc[output['Metric'] == 'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 20.125% (OPN, TPN, DFC)', 'Total Vast New Cumulative'] = total_20p125
  output.loc[output['Metric'] == 'Off Invoice 20.125% (OPN, TPN, DFC)', 'Total Vast New Per Unit'] = total_20p125/divisor
  output.loc[output['Metric'] == 'Off Invoice 20.125% (OPN, TPN, DFC)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 20.125% (OPN, TPN, DFC)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 21.125% ST', 'Total Vast New Cumulative'] = total_21p125
  output.loc[output['Metric'] == 'Off Invoice 21.125% ST', 'Total Vast New Per Unit'] = total_21p125/divisor
  output.loc[output['Metric'] == 'Off Invoice 21.125% ST', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 21.125% ST', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 26.125% SU', 'Total Vast New Cumulative'] = total_26p125
  output.loc[output['Metric'] == 'Off Invoice 26.125% SU', 'Total Vast New Per Unit'] = total_26p125/divisor
  output.loc[output['Metric'] == 'Off Invoice 26.125% SU', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 26.125% SU', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 26.625% (FN, FNA, FNH)', 'Total Vast New Cumulative'] = total_26p625
  output.loc[output['Metric'] == 'Off Invoice 26.625% (FN, FNA, FNH)', 'Total Vast New Per Unit'] = total_26p625/divisor
  output.loc[output['Metric'] == 'Off Invoice 26.625% (FN, FNA, FNH)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 26.625% (FN, FNA, FNH)', 'Total Vast Current Cumulative'] = 0
  output.loc[output['Metric'] == 'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)', 'Total Vast New Cumulative'] = total_28
  output.loc[output['Metric'] == 'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)', 'Total Vast New Per Unit'] = total_28/divisor
  output.loc[output['Metric'] == 'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)', 'Total Vast Current Per Unit'] = 0
  output.loc[output['Metric'] == 'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)', 'Total Vast Current Cumulative'] = 0

  print("Calculating Blank Expense...")
  blank_expense = {
    'CKP': 0.1,
    'CMP': 0.1,
    'COM': 0.12,
    'CON': 0.12,
    'EMC': 0.1,
    'ETB': 0.12,
    'EVA': 0.05,
    'FIP': 0.2,
    'FM': 0.12,
    'FN': 0.05,
    'FNH': 0.02,
    'FP': 0.08,
    'GTN': 0.02,
    'HDH': 0.12,
    'HEA': 0.12,
    'IGN': 0.1,
    'INT': 0.05,
    'LO': 0.05,
    'MAF': 0.12,
    'MAP': 0.2,
    'MFP': 0.05,
    'OPN': 0.1,
    'ST': 0.2,
    'SU': 0.25,
    'SUC': 0.05,
    'TPN': 0.05
  }
  total = 0
  for pline in plines:
    if pline in blank_expense:
      sales = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == '-', f'{pline} New Vast Cumulative'] = blank_expense[pline] * sales * -1
      output.loc[output['Metric'] == '-', f'{pline} New Vast Per Unit'] = getPerUnit('-', f'{pline} New Vast Cumulative')
      total += output.loc[output['Metric'] == '-', f'{pline} New Vast Cumulative'].iloc[0]
  output.loc[output['Metric'] == '-', 'Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == '-', 'Total Vast New Per Unit'] = total/divisor

  current_rows = [
    'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)',
    'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)',
    'Volume rebate 20% (CMP, CKP, MAF, IGN)',
    'Exchange Rate Program 5% COM',
    'Loyalty Rebate 1% ALL',
    'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)',
    'Volume Rebate Additional 7.5% (EVA, ACC, CO, CON, CGT, SUC, COM, CMP, CKP, MAP, MAF, OXS, ETB, IGN, VTP, VVT, FP, FPA, FPD, SF, MFP, FIP, FM, SU, CHD, CIN, CIA, CAC, INT, EMC, FC, BPA, OPN, TPN, FNH, LO, GTP, FNA, FN, FRA, ST, GTN, DFC)',
    'Marketing Rebate 7.5% (FP, SF, MFP, FM)',
    'Enhancement Rebate 8.3% (FP, SF, MFP, FM, FTA)',
    'Fuel Delivery Marketing 5% (FP, FPA, SF, MFP, FM)',
    'Heater Loyalty Rebate 9% (HEA, HDH)',
    'Alliance Volume Rebate 8% ALL',
    'Alliance Fuel Delivery 6% (SUC, FP, FPA, FPD, SF, MFP, FIP, FM, SU)'
  ]

  new_rows = [
    'Support Rebate 1% ALL',
    'Volume Rebate 1.5% (CON, CHD)',
    'Business Development Fund 2% (CMP, CKP, ETB, IGN, MAP, MAF, OXS, VTP, VVT)',
    'Off Invoice 0.5% VTP',
    'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)',
    'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)',
    'Off Invoice 6% CHD',
    'Off Invoice 6.25% (CIA, CAC, CIN)',
    'Off Invoice 6.875% (CON, COM, INT)',
    'Off Invoice 7.125% (HEA, HDH)',
    'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)',
    'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)',
    'Off Invoice 20.125% (OPN, TPN, DFC)',
    'Off Invoice 21.125% ST',
    'Off Invoice 26.125% SU',
    'Off Invoice 26.625% (FN, FNA, FNH)',
    'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)'
  ]

  print('Calculation Agency Fee...')
  total_current = 0
  total_new = 0
  for pline in plines:
    agency_current = 0
    agency_new = 0
    sales_current = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
    for rows in current_rows:
      agency_current += output.loc[output['Metric'] == rows, f'{pline} Current Vast Cumulative'].iloc[0]
    for rows in new_rows:
      agency_new += output.loc[output['Metric'] == rows, f'{pline} New Vast Cumulative'].iloc[0]
    if pline != 'FNH':
      if pline == "IGN":
        agency_current -= output.loc[output['Metric'] == 'Volume rebate 20% (CMP, CKP, MAF, IGN)', f'{pline} Current Vast Cumulative'].iloc[0]
      elif pline == "MFP":
        agency_current -= output.loc[output['Metric'] == 'Volume Rebate Additional 7.5% (EVA, ACC, CO, CON, CGT, SUC, COM, CMP, CKP, MAP, MAF, OXS, ETB, IGN, VTP, VVT, FP, FPA, FPD, SF, MFP, FIP, FM, SU, CHD, CIN, CIA, CAC, INT, EMC, FC, BPA, OPN, TPN, FNH, LO, GTP, FNA, FN, FRA, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'].iloc[0]
        agency_current -= output.loc[output['Metric'] == 'Marketing Rebate 7.5% (FP, SF, MFP, FM)', f'{pline} Current Vast Cumulative'].iloc[0]
        agency_current -= output.loc[output['Metric'] == 'Enhancement Rebate 8.3% (FP, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'].iloc[0]
        agency_current -= output.loc[output['Metric'] == 'Fuel Delivery Marketing 5% (FP, FPA, SF, MFP, FM)', f'{pline} Current Vast Cumulative'].iloc[0]
        agency_current -= output.loc[output['Metric'] == 'Alliance Volume Rebate 8% ALL', f'{pline} Current Vast Cumulative'].iloc[0]
        agency_current -= output.loc[output['Metric'] == 'Alliance Fuel Delivery 6% (SUC, FP, FPA, FPD, SF, MFP, FIP, FM, SU)', f'{pline} Current Vast Cumulative'].iloc[0]
        agency_current -= output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'].iloc[0]
      elif pline == "ST":
        agency_current -= output.loc[output['Metric'] == 'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)', f'{pline} Current Vast Cumulative'].iloc[0]
      elif pline == "SF":
        agency_current -= output.loc[output['Metric'] == 'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)', f'{pline} Current Vast Cumulative'].iloc[0]
      output.loc[output['Metric'] == 'Agency fees', f'{pline} Current Vast Cumulative'] = (agency_current + sales_current) * -0.025
      output.loc[output['Metric'] == 'Agency fees', f'{pline} Current Vast Per Unit'] = getPerUnit('Agency fees', f'{pline} Current Vast Cumulative')
      total_current += (agency_current + sales_current) * -0.025
    else:
      output.loc[output['Metric'] == 'Agency fees', f'{pline} Current Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Agency fees', f'{pline} Current Vast Per Unit'] = getPerUnit('Agency fees', f'{pline} Current Vast Cumulative')
    if pline != 'FPD':
      output.loc[output['Metric'] == 'Agency fees', f'{pline} New Vast Cumulative'] = (agency_new + sales_new) * -0.025
      output.loc[output['Metric'] == 'Agency fees', f'{pline} New Vast Per Unit'] = getPerUnit('Agency fees', f'{pline} New Vast Cumulative')
      total_new += (agency_new + sales_new) * -0.025
    else:
      output.loc[output['Metric'] == 'Agency fees', f'{pline} New Vast Cumulative'] = 0
      output.loc[output['Metric'] == 'Agency fees', f'{pline} New Vast Per Unit'] = getPerUnit('Agency fees', f'{pline} New Vast Cumulative')
  output.loc[output['Metric'] == 'Agency fees', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Agency fees', 'Total Vast Current Per Unit'] = total_current/divisor
  output.loc[output['Metric'] == 'Agency fees', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Agency fees', 'Total Vast New Per Unit'] = total_new/divisor

  print("Calculating Return Allowance...")
  total_current = 0
  total_new = 0
  for pline in plines:
    sales_current = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Return Allowance', f'{pline} Current Vast Cumulative'] = sales_current*-0.05
    output.loc[output['Metric'] == 'Return Allowance', f'{pline} Current Vast Per Unit'] = getPerUnit('Return Allowance', f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Return Allowance', f'{pline} New Vast Cumulative'] = sales_new*-0.05
    output.loc[output['Metric'] == 'Return Allowance', f'{pline} New Vast Per Unit'] = getPerUnit('Return Allowance', f'{pline} New Vast Cumulative')
    total_current += sales_current*-0.05
    total_new += sales_new*-0.05
  output.loc[output['Metric'] == 'Return Allowance', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Return Allowance', 'Total Vast Current Per Unit'] = total_current/divisor
  output.loc[output['Metric'] == 'Return Allowance', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Return Allowance', 'Total Vast New Per Unit'] = total_new/divisor

  print("Calculating Defect Rate...")
  total_current = 0
  total_new = 0
  for pline in plines:
    sales_current = output.loc[output['Metric'] == 'Sales', f'{pline} Current Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'Sales', f'{pline} New Vast Cumulative'].iloc[0]
    defect_current = output.loc[output['Metric'] == 'Defect %', f'{pline} Current Vast Cumulative'].iloc[0]
    defect_new = output.loc[output['Metric'] == 'Defect %', f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Defect rate', f'{pline} Current Vast Cumulative'] = sales_current * -1 * defect_current
    output.loc[output['Metric'] == 'Defect rate', f'{pline} Current Vast Per Unit'] = getPerUnit('Defect rate', f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Defect rate', f'{pline} New Vast Cumulative'] = sales_new * -1 * defect_new
    output.loc[output['Metric'] == 'Defect rate', f'{pline} New Vast Per Unit'] = getPerUnit('Defect rate', f'{pline} New Vast Cumulative')
    total_current += sales_current * -1 * defect_current
    total_new += sales_new * -1 * defect_new
  output.loc[output['Metric'] == 'Defect rate', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Defect rate', 'Total Vast Current Per Unit'] = total_current/divisor
  output.loc[output['Metric'] == 'Defect rate', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Defect rate', 'Total Vast New Per Unit'] = total_new/divisor

  net_sale_rows = [
    'Sales',
    'Volume rebate 14% ( ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)',
    'Volume rebate 5% (COM, FP, FPA, SF, MFP, FM, FTA)',
    'Volume rebate 20% (CMP, CKP, MAF, IGN)',
    'Exchange Rate Program 5% COM',
    'Loyalty Rebate 1% ALL',
    'Loyalty Rebate 5% (COM, CMP, CKP, MAF, IGN)',
    'Volume Rebate Additional 7.5% (EVA, ACC, CO, CON, CGT, SUC, COM, CMP, CKP, MAP, MAF, OXS, ETB, IGN, VTP, VVT, FP, FPA, FPD, SF, MFP, FIP, FM, SU, CHD, CIN, CIA, CAC, INT, EMC, FC, BPA, OPN, TPN, FNH, LO, GTP, FNA, FN, FRA, ST, GTN, DFC)',
    'Marketing Rebate 7.5% (FP, SF, MFP, FM)',
    'Enhancement Rebate 8.3% (FP, SF, MFP, FM, FTA)',
    'Fuel Delivery Marketing 5% (FP, FPA, SF, MFP, FM)',
    'Heater Loyalty Rebate 9% (HEA, HDH)',
    'Alliance Volume Rebate 8% ALL',
    'Alliance Fuel Delivery 6% (SUC, FP, FPA, FPD, SF, MFP, FIP, FM, SU)',
    'Support Rebate 1% ALL',
    'Volume Rebate 1.5% (CON, CHD)',
    'Business Development Fund 2% (CMP, CKP, ETB, IGN, MAP, MAF, OXS, VTP, VVT)',
    'Off Invoice 0.5% VTP',
    'Off Invoice 0.625% (MAP, FPD, BPA, EMC, FC)',
    'Off Invoice 1.5% (ACC, CO, EVA, OXS, GTP)',
    'Off Invoice 6% CHD',
    'Off Invoice 6.25% (CIA, CAC, CIN)',
    'Off Invoice 6.875% (CON, COM, INT)',
    'Off Invoice 7.125% (HEA, HDH)',
    'Off Invoice 14.125% (SUC, FTA, LO, GTN, CGT)',
    'Off Invoice 19.125% (VVT, FP, FIP, FPA, FM, SF, MFP)',
    'Off Invoice 20.125% (OPN, TPN, DFC)',
    'Off Invoice 21.125% ST',
    'Off Invoice 26.125% SU',
    'Off Invoice 26.625% (FN, FNA, FNH)',
    'Off Invoice 28% (CMP, CKP, ETB, IGN, MAF)',
    '-',
    'Agency fees'
  ]

  print("Calculating NET SALES...")
  for pline in plines:
    net_sales_current = 0
    net_sales_new = 0
    if pline == 'ACC' or pline == 'CGT' or pline == 'CKP' or pline == 'CMP':
      net_sales_current += output.loc[output['Metric'] == 'Return Allowance', f'{pline} Current Vast Cumulative'].iloc[0]
      net_sales_current += output.loc[output['Metric'] == 'Defect rate', f'{pline} Current Vast Cumulative'].iloc[0]
      net_sales_new += output.loc[output['Metric'] == 'Return Allowance', f'{pline} New Vast Cumulative'].iloc[0]
      net_sales_new += output.loc[output['Metric'] == 'Defect rate', f'{pline} New Vast Cumulative'].iloc[0]
    for metric in net_sale_rows:
      if pd.notna(output.loc[output['Metric'] == metric, f'{pline} Current Vast Cumulative'].iloc[0]):
        net_sales_current += output.loc[output['Metric'] == metric, f'{pline} Current Vast Cumulative'].iloc[0]
      if pd.notna(output.loc[output['Metric'] == metric, f'{pline} New Vast Cumulative'].iloc[0]):
        net_sales_new += output.loc[output['Metric'] == metric, f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'NET SALES', f'{pline} Current Vast Cumulative'] = net_sales_current
    output.loc[output['Metric'] == 'NET SALES', f'{pline} Current Vast Per Unit'] = getPerUnit('NET SALES', f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'NET SALES', f'{pline} New Vast Cumulative'] = net_sales_new
    output.loc[output['Metric'] == 'NET SALES', f'{pline} New Vast Per Unit'] = getPerUnit('NET SALES', f'{pline} New Vast Cumulative')
    if pline == "EVA":
      ACC_gross = output.loc[output['Metric'] == 'QTY Gross', 'ACC Current Vast Cumulative'].iloc[0]
      sales = output.loc[output['Metric'] == 'NET SALES', f'{pline} Current Vast Cumulative'].iloc[0]
      print(ACC_gross, sales)
      output.loc[output['Metric'] == 'NET SALES', f'{pline} Current Vast Per Unit'] = sales/ACC_gross
  net_sales_current = 0
  net_sales_new = 0
  for metric in net_sale_rows:
    if pd.notna(output.loc[output['Metric'] == metric, 'Total Vast Current Cumulative'].iloc[0]):
      net_sales_current += output.loc[output['Metric'] == metric, 'Total Vast Current Cumulative'].iloc[0]
    if pd.notna(output.loc[output['Metric'] == metric, 'Total Vast New Cumulative'].iloc[0]):
      net_sales_new += output.loc[output['Metric'] == metric, 'Total Vast New Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'NET SALES', 'Total Vast Current Cumulative'] = net_sales_current
  output.loc[output['Metric'] == 'NET SALES', 'Total Vast Current Per Unit'] = net_sales_current/divisor
  output.loc[output['Metric'] == 'NET SALES', 'Total Vast New Cumulative'] = net_sales_new
  output.loc[output['Metric'] == 'NET SALES', 'Total Vast New Per Unit'] = net_sales_new/divisor

  print('Calculating PO cost...')
  total = 0
  for pline in plines:
    data = pd.read_excel(file, sheet_name=pline, header=1)
    cost = (data['V42.2 PO cost']*data['Vast volume']).sum()
    output.loc[output['Metric'] == 'FG PO cost',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'FG PO cost',f'{pline} Current Vast Per Unit'] = getPerUnit('FG PO cost',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'FG PO cost',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'FG PO cost',f'{pline} New Vast Per Unit'] = getPerUnit('FG PO cost',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'FG PO cost','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'FG PO cost','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'FG PO cost','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'FG PO cost','Total Vast New Per Unit'] = total/divisor

  print('Calculating Variable - Overhead...')
  total = 0
  for pline in plines:
    cost = 0
    output.loc[output['Metric'] == 'Variable - Overhead',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Variable - Overhead',f'{pline} Current Vast Per Unit'] = getPerUnit('Variable - Overhead',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Variable - Overhead',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Variable - Overhead',f'{pline} New Vast Per Unit'] = getPerUnit('Variable - Overhead',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'Variable - Overhead','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Variable - Overhead','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Variable - Overhead','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Variable - Overhead','Total Vast New Per Unit'] = total/divisor

  print('Calculating Labor...')
  total = 0
  for pline in plines:
    cost = 0
    output.loc[output['Metric'] == 'Labor',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Labor',f'{pline} Current Vast Per Unit'] = getPerUnit('Labor',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Labor',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Labor',f'{pline} New Vast Per Unit'] = getPerUnit('Labor',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'Labor','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Labor','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Labor','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Labor','Total Vast New Per Unit'] = total/divisor

  print('Calculating Duty...')
  total = 0
  for pline in plines:
    data = pd.read_excel(file, sheet_name=pline, header=1)
    cost = (data['Duty']*data['Vast volume']).sum()
    output.loc[output['Metric'] == 'Duty',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Duty',f'{pline} Current Vast Per Unit'] = getPerUnit('Duty',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Duty',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Duty',f'{pline} New Vast Per Unit'] = getPerUnit('Duty',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'Duty','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Duty','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Duty','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Duty','Total Vast New Per Unit'] = total/divisor

  print('Calculating Tariffs...')
  total = 0
  for pline in plines:
    cost = 0
    output.loc[output['Metric'] == 'Tariffs',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Tariffs',f'{pline} Current Vast Per Unit'] = getPerUnit('Tariffs',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Tariffs',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Tariffs',f'{pline} New Vast Per Unit'] = getPerUnit('Tariffs',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'Tariffs','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Tariffs','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Tariffs','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Tariffs','Total Vast New Per Unit'] = total/divisor

  print('Calculating Duty...')
  total = 0
  for pline in plines:
    data = pd.read_excel(file, sheet_name=pline, header=1)
    cost = (data['Freight $']*data['Vast volume']).sum()
    output.loc[output['Metric'] == 'FG Freight',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'FG Freight',f'{pline} Current Vast Per Unit'] = getPerUnit('FG Freight',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'FG Freight',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'FG Freight',f'{pline} New Vast Per Unit'] = getPerUnit('FG Freight',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'FG Freight','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'FG Freight','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'FG Freight','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'FG Freight','Total Vast New Per Unit'] = total/divisor

  print('Calculating Freight interco...')
  total = 0
  for pline in plines:
    cost = 0
    output.loc[output['Metric'] == 'Freight interco',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Freight interco',f'{pline} Current Vast Per Unit'] = getPerUnit('Freight interco',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Freight interco',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Freight interco',f'{pline} New Vast Per Unit'] = getPerUnit('Freight interco',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'Freight interco','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Freight interco','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Freight interco','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Freight interco','Total Vast New Per Unit'] = total/divisor

  print('Calculating Scrap Return...')
  total_current = 0
  total_new = 0
  for pline in plines:
    return_allowance_current = output.loc[output['Metric'] == 'Return Allowance',f'{pline} Current Vast Cumulative'].iloc[0]
    sales_current = output.loc[output['Metric'] == 'NET SALES',f'{pline} Current Vast Per Unit'].iloc[0]
    PO_cost_current = output.loc[output['Metric'] == 'FG PO cost',f'{pline} Current Vast Per Unit'].iloc[0]
    variable_overhead_current = output.loc[output['Metric'] == 'Variable - Overhead',f'{pline} Current Vast Per Unit'].iloc[0]
    labor_current = output.loc[output['Metric'] == 'Labor',f'{pline} Current Vast Per Unit'].iloc[0]
    duty_current = output.loc[output['Metric'] == 'Duty',f'{pline} Current Vast Per Unit'].iloc[0]
    FG_freight_current = output.loc[output['Metric'] == 'FG Freight',f'{pline} Current Vast Per Unit'].iloc[0]
    tariffs_current = output.loc[output['Metric'] == 'Tariffs',f'{pline} Current Vast Per Unit'].iloc[0]

    return_allowance_new = output.loc[output['Metric'] == 'Return Allowance',f'{pline} New Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'NET SALES',f'{pline} New Vast Per Unit'].iloc[0]
    PO_cost_new = output.loc[output['Metric'] == 'FG PO cost',f'{pline} New Vast Per Unit'].iloc[0]
    variable_overhead_new = output.loc[output['Metric'] == 'Variable - Overhead',f'{pline} New Vast Per Unit'].iloc[0]
    labor_new = output.loc[output['Metric'] == 'Labor',f'{pline} New Vast Per Unit'].iloc[0]
    duty_new = output.loc[output['Metric'] == 'Duty',f'{pline} New Vast Per Unit'].iloc[0]
    if pline == 'ACC':
      FG_freight_new = 0
    else:
      FG_freight_new = output.loc[output['Metric'] == 'FG Freight',f'{pline} New Vast Per Unit'].iloc[0]
    tariffs_new = output.loc[output['Metric'] == 'Tariffs',f'{pline} New Vast Per Unit'].iloc[0]
    value_current = (1-0.75)*(return_allowance_current/sales_current)*(PO_cost_current + variable_overhead_current + labor_current + duty_current + FG_freight_current + tariffs_current)
    value_new = (1-0.75)*(return_allowance_new/sales_new)*(PO_cost_new + variable_overhead_new + labor_new + duty_new + FG_freight_new + tariffs_new)

    output.loc[output['Metric'] == 'Scrap Return % (Resellable %)',f'{pline} Current Vast Cumulative'] = value_current
    output.loc[output['Metric'] == 'Scrap Return % (Resellable %)',f'{pline} Current Vast Per Unit'] = getPerUnit('Scrap Return % (Resellable %)',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Scrap Return % (Resellable %)',f'{pline} New Vast Cumulative'] = value_new
    output.loc[output['Metric'] == 'Scrap Return % (Resellable %)',f'{pline} New Vast Per Unit'] = getPerUnit('Scrap Return % (Resellable %)',f'{pline} New Vast Cumulative')
    total_current += value_current
    total_new += value_new
  output.loc[output['Metric'] == 'Scrap Return % (Resellable %)','Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Scrap Return % (Resellable %)','Total Vast Current Per Unit'] = total_current/divisor
  output.loc[output['Metric'] == 'Scrap Return % (Resellable %)','Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Scrap Return % (Resellable %)','Total Vast New Per Unit'] = total_new/divisor

  total_variable_rows = [
    'FG PO cost',
    'Variable - Overhead',
    'Labor',
    'Duty',
    'Scrap Return % (Resellable %)',
    'Tariffs',
    'FG Freight',
    'Freight interco'
  ]

  print('Calculating Total Variable...')
  for pline in plines:
    total_new = 0
    total_current = 0
    for row in total_variable_rows:
      total_current += output.loc[output['Metric'] == row,f'{pline} Current Vast Cumulative'].iloc[0]
      total_new += output.loc[output['Metric'] == row,f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'TOTAL VARIABLE COST',f'{pline} Current Vast Cumulative'] = total_current
    output.loc[output['Metric'] == 'TOTAL VARIABLE COST',f'{pline} Current Vast Per Unit'] = getPerUnit('TOTAL VARIABLE COST',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'TOTAL VARIABLE COST',f'{pline} New Vast Cumulative'] = total_new
    output.loc[output['Metric'] == 'TOTAL VARIABLE COST',f'{pline} New Vast Per Unit'] = getPerUnit('TOTAL VARIABLE COST',f'{pline} New Vast Cumulative')
  total_new = 0
  total_current = 0
  for metric in total_variable_rows:
    if pd.notna(output.loc[output['Metric'] == metric, 'Total Vast Current Cumulative'].iloc[0]):
      total_current += output.loc[output['Metric'] == metric, 'Total Vast Current Cumulative'].iloc[0]
    if pd.notna(output.loc[output['Metric'] == metric, 'Total Vast New Cumulative'].iloc[0]):
      total_new += output.loc[output['Metric'] == metric, 'Total Vast New Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'Total Vast Current Per Unit'] = total_current/divisor
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'Total Vast New Per Unit'] = total_new/divisor

  print('Calculating Margin...')
  for pline in plines:
    sales_current = output.loc[output['Metric'] == 'NET SALES',f'{pline} Current Vast Cumulative'].iloc[0]
    total_variable_current = output.loc[output['Metric'] == 'TOTAL VARIABLE COST',f'{pline} Current Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'NET SALES',f'{pline} New Vast Cumulative'].iloc[0]
    total_variable_new = output.loc[output['Metric'] == 'TOTAL VARIABLE COST',f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'MARGIN',f'{pline} Current Vast Cumulative'] = sales_current - total_variable_current
    output.loc[output['Metric'] == 'MARGIN',f'{pline} New Vast Cumulative'] = sales_new - total_variable_new
    output.loc[output['Metric'] == 'MARGIN',f'{pline} Current Vast Per Unit'] = getPerUnit('MARGIN',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'MARGIN',f'{pline} New Vast Per Unit'] = getPerUnit('MARGIN',f'{pline} New Vast Cumulative')
  total_sales_current = output.loc[output['Metric'] == 'NET SALES', 'Total Vast Current Cumulative'].iloc[0]
  total_variable_current = output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'Total Vast Current Cumulative'].iloc[0]
  total_sales_new = output.loc[output['Metric'] == 'NET SALES', 'Total Vast New Cumulative'].iloc[0]
  total_variable_new = output.loc[output['Metric'] == 'TOTAL VARIABLE COST', 'Total Vast New Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'MARGIN', 'Total Vast Current Cumulative'] = total_sales_current - total_variable_current
  output.loc[output['Metric'] == 'MARGIN', 'Total Vast New Cumulative'] = total_sales_new - total_variable_new
  output.loc[output['Metric'] == 'MARGIN', 'Total Vast Current Per Unit'] = total_sales_current - total_variable_current
  output.loc[output['Metric'] == 'MARGIN', 'Total Vast New Per Unit'] = total_sales_new - total_variable_new

  for columns in output.columns:
    if 'Cumulative' in columns:
      margin = output.loc[output['Metric'] == 'MARGIN', columns].iloc[0]
      sales = output.loc[output['Metric'] == 'NET SALES', columns].iloc[0]
      output.loc[output['Metric'] == 'MARGIN %', columns] = margin / sales

  print('Calculating Handling...')
  total = 0
  for pline in plines:
    output.loc[output['Metric'] == 'Handling/Shipping',f'{pline} Current Vast Per Unit'] = 1.85
    output.loc[output['Metric'] == 'Handling/Shipping',f'{pline} New Vast Per Unit'] = 1.85
    gross = output.loc[output['Metric'] == 'QTY Gross',f'{pline} Current Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Handling/Shipping',f'{pline} Current Vast Cumulative'] = 1.85 * gross
    output.loc[output['Metric'] == 'Handling/Shipping',f'{pline} New Vast Cumulative'] = 1.85 * gross
    total += 1.85 * gross
  output.loc[output['Metric'] == 'Handling/Shipping', 'Total Vast Current Per Unit'] = 1.85
  output.loc[output['Metric'] == 'Handling/Shipping', 'Total Vast New Per Unit'] = 1.85
  output.loc[output['Metric'] == 'Handling/Shipping', 'Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Handling/Shipping', 'Total Vast New Cumulative'] = total

  print('Calculating Fill Rates...')
  total_current = 0
  total_new = 0
  for pline in plines:
    sales_current = output.loc[output['Metric'] == 'Sales',f'{pline} Current Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'Sales',f'{pline} New Vast Cumulative'].iloc[0]
    agency_current = output.loc[output['Metric'] == 'Agency fees',f'{pline} Current Vast Cumulative'].iloc[0]
    agency_new = output.loc[output['Metric'] == 'Agency fees',f'{pline} New Vast Cumulative'].iloc[0]
    print(sales_current, agency_current, (sales_current + agency_current) * 0.01)
    output.loc[output['Metric'] == 'Fill rate fines',f'{pline} Current Vast Cumulative'] = (sales_current + agency_current) * 0.01
    output.loc[output['Metric'] == 'Fill rate fines',f'{pline} Current Vast Per Unit'] = getPerUnit('Fill rate fines',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Fill rate fines',f'{pline} New Vast Cumulative'] = (sales_new + agency_new) * 0.01
    output.loc[output['Metric'] == 'Fill rate fines',f'{pline} New Vast Per Unit'] = getPerUnit('Fill rate fines',f'{pline} New Vast Cumulative')
    total_current += (sales_current + agency_current) * 0.01
    total_new += (sales_new + agency_new) * 0.01
  output.loc[output['Metric'] == 'Fill rate fines', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Fill rate fines', 'Total Vast Current Per Unit'] = getPerUnit('Fill rate fines', 'Total Vast Current Cumulative')
  output.loc[output['Metric'] == 'Fill rate fines', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Fill rate fines', 'Total Vast New Per Unit'] = getPerUnit('Fill rate fines', 'Total Vast New Cumulative')

  print('Calculating Inspect of Return...')
  total_current = 0
  total_new = 0
  for pline in plines:
    output.loc[output['Metric'] == 'Inspect of Return',f'{pline} Current Vast Per Unit'] = 1.85
    output.loc[output['Metric'] == 'Inspect of Return',f'{pline} New Vast Per Unit'] = 1.85
    return_allowance_current = output.loc[output['Metric'] == 'Return Allowance',f'{pline} Current Vast Cumulative'].iloc[0]
    if pline == 'FP':
      return_allowance_current = output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL',f'{pline} Current Vast Cumulative'].iloc[0]
    return_allowance_new = output.loc[output['Metric'] == 'Return Allowance',f'{pline} New Vast Cumulative'].iloc[0]
    sales_current = output.loc[output['Metric'] == 'NET SALES',f'{pline} Current Vast Per Unit'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'NET SALES',f'{pline} New Vast Per Unit'].iloc[0]
    output.loc[output['Metric'] == 'Inspect of Return',f'{pline} Current Vast Cumulative'] = 1.85 * return_allowance_current/sales_current*-1
    output.loc[output['Metric'] == 'Inspect of Return',f'{pline} New Vast Cumulative'] = 1.85 * return_allowance_new/sales_new*-1
    total_current += 1.85 * return_allowance_current/sales_current*-1
    total_new += 1.85 * return_allowance_new/sales_new*-1
  output.loc[output['Metric'] == 'Inspect of Return', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Inspect of Return', 'Total Vast Current Per Unit'] = 1.85
  output.loc[output['Metric'] == 'Inspect of Return', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Inspect of Return', 'Total Vast New Per Unit'] = 1.85

  print("Calculating Put Away Box...")
  total_current = 0
  total_new = 0
  box = pd.read_excel(file, sheet_name='Box', header=4)
  box = box.set_index('Pline').T
  for pline in plines:
    if pline == 'TPN':
      value = 1.34
    elif pline in box.columns:
      value = box[f'{pline}'].iloc[0]
    else:
      value = 0
    output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product',f'{pline} Current Vast Per Unit'] = value + 1.85
    output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product',f'{pline} New Vast Per Unit'] = value + 1.85
    if pline == 'ETB' or pline == 'EVA' or pline == 'FN':
      output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product',f'{pline} New Vast Per Unit'] = 1.38 + 1.85
    return_allowance_current = output.loc[output['Metric'] == 'Return Allowance',f'{pline} Current Vast Cumulative'].iloc[0]
    if pline == 'CGT' or pline == 'CKP' or pline == 'CMP' or pline == 'FP' or pline == 'MAP' or pline == 'OPN':
      return_allowance_current = output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL',f'{pline} Current Vast Cumulative'].iloc[0]
    return_allowance_new = output.loc[output['Metric'] == 'Return Allowance',f'{pline} New Vast Cumulative'].iloc[0]
    if pline == 'CKP' or pline == 'CMP':
      return_allowance_new = output.loc[output['Metric'] == 'Loyalty Rebate 1% ALL',f'{pline} New Vast Cumulative'].iloc[0]
    sales_current = output.loc[output['Metric'] == 'NET SALES',f'{pline} Current Vast Per Unit'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'NET SALES',f'{pline} New Vast Per Unit'].iloc[0]
    output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product',f'{pline} Current Vast Cumulative'] = (value + 1.85) * return_allowance_current/sales_current*-1*(1-0.75)
    output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product',f'{pline} New Vast Cumulative'] = (value + 1.85) * return_allowance_new/sales_new*-1*(1-0.75)
    if pline == 'ETB' or pline == 'EVA' or pline == 'FN':
      output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product',f'{pline} New Vast Cumulative'] = (1.38 + 1.85) * return_allowance_new/sales_new*-1*(1-0.75)
    total_current += (value + 1.85) * return_allowance_current/sales_current*-1*(1-0.75)
    if pline == 'ETB' or pline == 'EVA' or pline == 'FN':
      total_new += (1.38 + 1.85) * return_allowance_new/sales_new*-1*(1-0.75)
    else:
      total_new += (value + 1.85) * return_allowance_new/sales_new*-1*(1-0.75)
  output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', 'Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', 'Total Vast Current Per Unit'] = 1.34 + 1.85
  output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', 'Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Return Allowance Put Away Costs, for usable product', 'Total Vast New Per Unit'] = 1.34 + 1.85
  print('Done and exporting to Excel...')

  print('Calculating Special Marketing (we control)...')
  total = 0
  for pline in plines:
    cost = 0
    output.loc[output['Metric'] == 'Special Marketing (we control)',f'{pline} Current Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Special Marketing (we control)',f'{pline} Current Vast Per Unit'] = getPerUnit('Special Marketing (we control)',f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Special Marketing (we control)',f'{pline} New Vast Cumulative'] = cost
    output.loc[output['Metric'] == 'Special Marketing (we control)',f'{pline} New Vast Per Unit'] = getPerUnit('Special Marketing (we control)',f'{pline} Current Vast Cumulative')
    total += cost
  output.loc[output['Metric'] == 'Special Marketing (we control)','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Special Marketing (we control)','Total Vast Current Per Unit'] = total/divisor
  output.loc[output['Metric'] == 'Special Marketing (we control)','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Special Marketing (we control)','Total Vast New Per Unit'] = total/divisor

  pallet = {
    'ACC': 0.06,
    'CGT': 0.84,
    'CKP': 0.01,
    'CMP': 0.01,
    'CO': 0.07,
    'COM': 0.4,
    'CON': 0.3,
    'EMC': 0.23,
    'ETB': 0.03,
    'EVA': 0.1,
    'FIP': 0.02,
    'FM': 0.09,
    'FN': 0.17,
    'FNH': 0.03,
    'FP': 0.01,
    'FPD': 0.01,
    'GTN': 0.84,
    'HDH': 0.07,
    'HEA': 0.07,
    'IGN': 0.01,
    'INT': 0.54,
    'LO': 0.00,
    'MAF': 0.01,
    'MAP': 0.00,
    'MFP': 0.02,
    'OPN': 0.15,
    'SF': 0.00,
    'ST': 0.07,
    'SU': 0.08,
    'SUC': 0.08,
    'TPN': 0.15
  }
  total = 0
  for pline in plines:
    output.loc[output['Metric'] == 'Pallets / Wrapping',f'{pline} Current Vast Per Unit'] = pallet[pline]
    output.loc[output['Metric'] == 'Pallets / Wrapping',f'{pline} New Vast Per Unit'] = pallet[pline]
    gross = output.loc[output['Metric'] == 'QTY Gross',f'{pline} Current Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Pallets / Wrapping',f'{pline} Current Vast Cumulative'] = gross*pallet[pline]
    output.loc[output['Metric'] == 'Pallets / Wrapping',f'{pline} New Vast Cumulative'] = gross*pallet[pline]
    total += output.loc[output['Metric'] == 'Pallets / Wrapping',f'{pline} Current Vast Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Pallets / Wrapping','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Pallets / Wrapping','Total Vast Current Per Unit'] = 0.15
  output.loc[output['Metric'] == 'Pallets / Wrapping','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Pallets / Wrapping','Total Vast New Per Unit'] = 0.15

  print('Calculatung Delivery...')
  delivery = {
    'ACC': 2,
    'CGT': 5,
    'CKP': 1,
    'CMP': 1,
    'CO': 2,
    'COM': 5,
    'CON': 4,
    'EMC': 2,
    'ETB': 2,
    'EVA': 3,
    'FIP': 2,
    'FM': 3,
    'FN': 4,
    'FNH': 2,
    'FP': 1,
    'FPD': 1,
    'GTN': 9,
    'HDH': 3,
    'HEA': 3,
    'IGN': 1,
    'INT': 5,
    'LO': 0.5,
    'MAF': 2,
    'MAP': 0.5,
    'MFP': 2,
    'OPN': 4,
    'SF': 0.5,
    'ST': 2,
    'SU': 3,
    'SUC': 3,
    'TPN': 4
  }
  total = 0
  for pline in plines:
    output.loc[output['Metric'] == 'Delivery',f'{pline} Current Vast Per Unit'] = delivery[pline]
    output.loc[output['Metric'] == 'Delivery',f'{pline} New Vast Per Unit'] = delivery[pline]
    gross = output.loc[output['Metric'] == 'QTY Gross',f'{pline} Current Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Delivery',f'{pline} Current Vast Cumulative'] = gross*delivery[pline]
    output.loc[output['Metric'] == 'Delivery',f'{pline} New Vast Cumulative'] = gross*delivery[pline]
    total += output.loc[output['Metric'] == 'Delivery',f'{pline} Current Vast Cumulative'].iloc[0]
  output.loc[output['Metric'] == 'Delivery','Total Vast Current Cumulative'] = total
  output.loc[output['Metric'] == 'Delivery','Total Vast Current Per Unit'] = 4
  output.loc[output['Metric'] == 'Delivery','Total Vast New Cumulative'] = total
  output.loc[output['Metric'] == 'Delivery','Total Vast New Per Unit'] = 4

  SGA_rows = [
    'Handling/Shipping',
    'Fill rate fines',
    'Inspect of Return',
    'Return Allowance Put Away Costs, for usable product',
    'Special Marketing (we control)',
    'Pallets / Wrapping',
    'Delivery'
  ]

  print('Calculating SG and A...')
  for columns in output.columns:
    if 'Cumulative' in columns:
      total = 0
      for row in SGA_rows:
        total += output.loc[output['Metric'] == row, columns].iloc[0]
      output.loc[output['Metric'] == 'SG&A', columns] = total
      if 'Total' in columns:
        over = output.loc[output['Metric'] == 'QTY Gross', 'TPN Current Vast Cumulative'].iloc[0]
        output.loc[output['Metric'] == 'SG&A', column.replace('Cumulative', 'Per Unit')] = total / over
      else:
        output.loc[output['Metric'] == 'SG&A', columns.replace('Cumulative', 'Per Unit')] = getPerUnit('SG&A', columns)

  print('calculating for payment term...')
  total_current = 0
  total_new = 0
  for pline in plines:
    sales_current = output.loc[output['Metric'] == 'NET SALES',f'{pline} Current Vast Cumulative'].iloc[0]
    sales_new = output.loc[output['Metric'] == 'NET SALES',f'{pline} New Vast Cumulative'].iloc[0]
    output.loc[output['Metric'] == 'Payment term',f'{pline} Current Vast Cumulative'] = sales_current * 0.03
    output.loc[output['Metric'] == 'Payment term',f'{pline} Current Vast Per Unit'] = getPerUnit('Payment term', f'{pline} Current Vast Cumulative')
    output.loc[output['Metric'] == 'Payment term',f'{pline} New Vast Cumulative'] = sales_new * 0.03
    output.loc[output['Metric'] == 'Payment term',f'{pline} New Vast Per Unit'] = getPerUnit('Payment term', f'{pline} New Vast Cumulative')
    total_current += sales_current * 0.03
    total_new += sales_new * 0.03
  output.loc[output['Metric'] == 'Payment term','Total Vast Current Cumulative'] = total_current
  output.loc[output['Metric'] == 'Payment term','Total Vast Current Per Unit'] = total_current / divisor
  output.loc[output['Metric'] == 'Payment term','Total Vast New Cumulative'] = total_new
  output.loc[output['Metric'] == 'Payment term','Total Vast New Per Unit'] = total_new / divisor

  print('Calculating contribution margin...')
  for columns in output.columns:
    if 'Cumulative' in columns:
      margin = output.loc[output['Metric'] == 'MARGIN', columns].iloc[0]
      sga = output.loc[output['Metric'] == 'SG&A', columns].iloc[0]
      payment_term = output.loc[output['Metric'] == 'Payment term', columns].iloc[0]
      net_sales = output.loc[output['Metric'] == 'NET SALES', columns].iloc[0]
      contrib_margin = margin - sga - payment_term
      output.loc[output['Metric'] == 'Contribution Margin $s', columns] = contrib_margin
      output.loc[output['Metric'] == 'Contribution Margin %', columns] = contrib_margin / net_sales
      if 'Total' in columns:
        over = output.loc[output['Metric'] == 'QTY Gross', 'TPN Current Vast Cumulative'].iloc[0]
        print(contrib_margin, over, column.replace('Cumulative', 'Per Unit'), contrib_margin / over)
        output.loc[output['Metric'] == 'Contribution Margin $s', columns.replace('Cumulative', 'Per Unit')] = contrib_margin / over
      else:
        output.loc[output['Metric'] == 'Contribution Margin $s', columns.replace('Cumulative', 'Per Unit')] = getPerUnit('Contribution Margin $s', columns)

  psuedo_defaults = {
    'Current': {
      'Volume rebate (ACC, EVA, CO, CON, CGT, SUC, HEA, HDH, OXS, VVT, SU, CHD, OPN, TPN, FN, ST, GTN, DFC)': 0.14,
      'Volume rebate (COM, FP, FPA, SF, MFP, FM, FTA)': 0.05,
      'Volume rebate (CMP, CKP, MAF, IGN)': 0.2,
      'Exchange Rate Program (COM)': 0.05,
      'Loyalty Rebate (ALL)': 0.01,
      'Loyalty Rebate (COM, CMP, CKP, MAF, IGN)': 0.05,
      'Volume Rebate Additional (EVA, ACC, CO, CON, CGT, SUC, COM, CMP, CKP, MAP, MAF, OXS, ETB, IGN, VTP, VVT, FP, FPA, FPD, SF, MFP, FIP, FM, SU, CHD, CIN, CIA, CAC, INT, EMC, FC, BPA, OPN, TPN, FNH, LO, GTP, FNA, FN, FRA, ST, GTN, DFC)': 0.075,
      'Marketing Rebate (FP, SF, MFP, FM)': 0.075,
      'Enhancement Rebate (FP, SF, MFP, FM, FTA)': 0.083,
      'Fuel Delivery Marketing (FP, FPA, SF, MFP, FM)': 0.05,
      'Heater Loyalty Rebate (HEA, HDH)': 0.09,
      'Alliance Volume Rebate (ALL)': 0.08,
      'Alliance Fuel Delivery (SUC, FP, FPA, FPD, SF, MFP, FIP, FM, SU)': 0.06
    },
    'New': {
      'Support Rebate (ALL)': 0.01,
      'Volume Rebate (CON, CHD)': 0.015,
      'Business Development Fund (CMP, CKP, ETB, IGN, MAP, MAF, OXS, VTP, VVT)': 0.02,
      'Off Invoice (VTP)': 0.005,
      'Off Invoice (MAP, FPD, BPA, EMC, FC)': 0.00625,
      'Off Invoice (ACC, CO, EVA, OXS, GTP)': 0.015,
      'Off Invoice (CHD)': 0.06,
      'Off Invoice (CIA, CAC, CIN)': 0.0625,
      'Off Invoice (CON, COM, INT)': 0.06875,
      'Off Invoice (HEA, HDH)': 0.075,
      'Off Invoice (SUC, FTA, LO, GTN, CGT)': 0.14,
      'Off Invoice (VVT, FP, FIP, FPA, FM, SF, MFP)': 0.19,
      'Off Invoice (OPN, TPN, DFC)': 0.20,
      'Off Invoice (ST)': 0.21125,
      'Off Invoice (SU)': 0.26125,
      'Off Invoice (FN, FNA, FNH)': 0.26625,
      'Off Invoice (CMP, CKP, ETB, IGN, MAF)': 0.28
    }
  }

  psuedo_assumptions = {
    'Agency fees': 0.025,
    'Return Allowance': 0.05,
    'Scrap Return % (Resellable %)': 0.75,
    'Fill rate fines': 0.05,
    'Payment term': 0.05
  }

  return output, psuedo_defaults, psuedo_assumptions