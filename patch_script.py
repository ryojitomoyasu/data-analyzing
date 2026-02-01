import os

file_path = r"c:\Users\ryoji\OneDrive\GSF\00. Antigravity\戦略考案・相談アプリ2026\pages\1_📊_Sales_Dashboard.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Target block 1: YoY calculation logic
target1 = """                          # Correctly handle columns if unstack resulted in year columns
                          if str(sec_curr_year) in df_trend_totals.columns and str(sec_prev_year) in df_trend_totals.columns:
                              df_trend_totals['YoY %'] = df_trend_totals.apply(lambda r: ((r[str(sec_curr_year)] - r[str(sec_prev_year)]) / r[str(sec_prev_year)] * 100) if r[str(sec_prev_year)] != 0 else np.nan, axis=1)
                          else:
                              df_trend_totals['YoY %'] = np.nan"""

replacement1 = """                          # Robust Year Logic
                          curr_key = str(sec_curr_year)
                          prev_key = str(sec_prev_year)
                          if curr_key not in df_trend_totals.columns: df_trend_totals[curr_key] = 0
                          if prev_key not in df_trend_totals.columns: df_trend_totals[prev_key] = 0

                          df_trend_totals['YoY %'] = df_trend_totals.apply(
                              lambda r: ((r[curr_key] - r[prev_key]) / r[prev_key] * 100) if (prev_key in r and r[prev_key] != 0) else 0.0, 
                              axis=1
                          )"""

# Target block 2: Legend name
target2 = 'name=f"YoY % vs {sec_prev_year}",'
replacement2 = 'name=f"YoY % vs {prev_key}",'

# Target block 3: Title
target3 = 'title=f"{t[\'trend_chart_title\']} ({sec_curr_year}) & YoY Growth (%)",'
replacement3 = 'title=f"{t[\'trend_chart_title\']} ({curr_key}) & YoY Growth (%)",'

# Target block 4: Legend markers+lines
target4 = "mode='lines+markers',"
replacement4 = "mode='markers+lines',"

# Standardize whitespace slightly for target1 to increase match chance if it still fails
# But I'll try literal first.

# Check if target1 is in content
if target1 in content:
    print("Found target 1")
    content = content.replace(target1, replacement1)
else:
    print("Target 1 NOT FOUND")
    # Try with potentially different line endings or whitespace
    # (Just a fallback)
    target1_alt = target1.replace('\r\n', '\n')
    if target1_alt in content.replace('\r\n', '\n'):
         print("Found target 1 (Line ending mismatch)")
         content = content.replace('\r\n', '\n').replace(target1_alt, replacement1.replace('\r\n', '\n'))

if target2 in content:
    print("Found target 2")
    content = content.replace(target2, replacement2)

if target3 in content:
    print("Found target 3")
    content = content.replace(target3, replacement3)
    
if target4 in content:
    print("Found target 4")
    content = content.replace(target4, replacement4)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement complete.")
