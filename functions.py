def formatter(name, value, defaults=False):
  if 'per unit' in name.lower() or 'cumulative' in name.lower() or 'big' in name.lower() or 'shipping' in name.lower() or 'return allowance' in name.lower() or 'pallets' in name.lower() or 'delivery' in name.lower() or 'inspect return' in name.lower() or 'rebox' in name.lower() or 'labor' in name.lower() or 'overhead' in name.lower() or 'special marketing' in name.lower():
    if defaults:
      value = f"% {value*100:.3f}"
    elif 'alliance' in name.lower():
      value = f"% {value*100:.3f}"
    else:
      value = f"$ {value:.2f}"
  elif 'qty' in name.lower():
    value = f"{value}"
  else:
    value = f"% {value*100:.3f}"
  return f'**{value}**'