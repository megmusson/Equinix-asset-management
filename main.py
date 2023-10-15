import openpyxl
import pandas as pd
import re

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Excel からインポートするデータ。
# df_pcru_sheet　は　Power Circuit Report Usage Sheet から　Cabinet, UPS を抜粋したデータ。必要な Cabinet しか含まれてない
# df_p_sheet は　Power Sheet から　Cabinet、UPS、Voltage を抜粋したデータ。全ての Cabinets が含まれてる
# import from excel and turn into dataframe
df_pcru_sheet = pd.read_excel('power_circuit_report_usage.xlsx',
                              sheet_name='Sheet1')
df_p_sheet = pd.read_excel('Power_sheet_extracted_data.xlsx',
                           sheet_name='Sheet1')

# Power Circuit Report Usage データ　を　Cabinet名でグループする
# Group by 'Cabinet' and aggregate the unique UPS names
pcru_cab_group = df_pcru_sheet.groupby(
    'Cabinet',
    sort=False)['UPS'].apply(lambda x: ', '.join(sorted(x.unique())))

# 一つの　UPS　にしか通船がない　Cabinet　を取り消す
# Filter out racks connected to only one unique UPS
pcru_cab_mult_ups = pcru_cab_group[pcru_cab_group.apply(
    lambda x: len(set(x.split(', '))) > 1)]

# Create a new DataFrame with the filtered results
df_pcru_cabinet = pd.DataFrame({
    'Cabinet': pcru_cab_mult_ups.index,
    'All UPS': pcru_cab_mult_ups.values
})

# Power　データから　Power Circuit Report Usage データに乗ってない　Cabinet　を取り消す
# Extract unique cabinets from df1
unique_cabinets = df_pcru_cabinet['Cabinet'].unique()
p_cab_filtered = df_p_sheet[df_p_sheet['Cabinet'].isin(unique_cabinets)]

# 二つのデータを一つにまとめる
merged_df = p_cab_filtered.merge(df_pcru_cabinet, on=['Cabinet'], how='inner')
print(merged_df.iloc[:50])

# 配線しているけど、接続していないコネクションを消す
# もしこのコネクションをEXCELに含みたかったらこの列を消す
# Filter rows where the UPS value is in the 'All UPS' list
merged_df = merged_df[merged_df.apply(
    lambda row: row['UPS'].split()[-1] in row['All UPS'], axis=1)]

# インデックスを並べ替えす
# Reset index without adding duplicate columns
merged_df = merged_df.reset_index(drop=True)


# ファンクション
# Modify the function to return only the relevant columns
def get_ups_and_voltage(x):
  unique_ups = ','.join(pd.unique(
      x['UPS'].str.split(' ').str[1]))  # Extract the number part
  unique_voltage_105 = ','.join(
      pd.unique(x['UPS'][x['Voltage'] == 105].str.split(' ').str[1]))
  unique_voltage_210 = ','.join(
      pd.unique(x['UPS'][x['Voltage'] == 210].str.split(' ').str[1]))
  return pd.Series([unique_ups, unique_voltage_105, unique_voltage_210],
                   index=['Connected UPS', '105V', '210V'])


# Apply the function without the 'Cabinet' column
grouped = merged_df.groupby('Cabinet').apply(get_ups_and_voltage).reset_index(
    level='Cabinet')

# Voltage を使ってない空きセルの入力希望
# 今は何もなし
# もし入れたかったら上と下に囲んでいる三つの引用符を消し、　'N/A'　に入れたい文字と交換
# そのまま　'N/A'　でもオッケー
# Replace empty cells with N/A for Voltage columns
"""
grouped['105V'] = grouped['105V'].replace('', 'N/A')
grouped['210V'] = grouped['210V'].replace('', 'N/A')
"""

# 出来上がったデータを　Excel　に入れる
# Save the results to Excel
grouped.to_excel('Power_Filtered.xlsx', index=False, engine='openpyxl')

print(
    'Excel sheet is ready. Please download Power_Filtered from the left hand side'
)
