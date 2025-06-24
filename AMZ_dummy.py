import pandas as pd

def getSummary(file):
  try:
    df = pd.read_excel(file, sheet_name='DATA V2')
    return df
  except Exception as e:
    return {"error": str(e)}