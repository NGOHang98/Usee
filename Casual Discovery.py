import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import japanize_matplotlib
from datetime import datetime, timedelta
import os
from scipy.interpolate import make_interp_spline
import numpy as np
import matplotlib as mpl

# Import Data
file_paths = [
    'C:/Users/HangNTT192/Desktop/Python/202206-202211.xlsx',
    'C:/Users/HangNTT192/Desktop/Python/202212-202306.xlsx'
]
data_frames = [
    pd.read_excel(file_path, sheet_name=sheet_name)
    for file_path in file_paths
    for sheet_name in pd.ExcelFile(file_path).sheet_names[:-1]
]
data = pd.concat(data_frames, ignore_index=True)

# Data Preprocessing
data.dropna(inplace=True)
data['SAGYO_UNYO_DATE'] = pd.to_datetime(data['SAGYO_UNYO_DATE'], format='%Y%m%d')
data['SGY_TM'] = pd.to_datetime(data['SGY_TM'], format='%Y/%m/%d %H:%M:%S')

# Data Analysis
def sales_comparison_by_periods(sales_df: pd.DataFrame, run_date: datetime, type: str, feats):
    if type == 'prev_day':
        history_day = run_date - pd.DateOffset(days=1)
    elif type == 'prev_week':
        history_day = run_date - pd.DateOffset(weeks=1)  
    elif type == 'prev_month':
        history_day = run_date - pd.DateOffset(months=1)
    elif type == 'prev_quarter':
        history_day = run_date - pd.DateOffset(months=3)
    elif type == 'prev_year':
        history_day = run_date - pd.DateOffset(years=1)
    else:
        raise ValueError('Invalid input')
    
    result = []
    for feat in feats:
        history_df = sales_df[sales_df['SAGYO_UNYO_DATE'] == history_day]
        curr_df = sales_df[sales_df['SAGYO_UNYO_DATE'] == run_date]
        history_total = history_df.groupby(feat, as_index=False)['SGY_JSK_QTY'].sum()
        curr_total = curr_df.groupby(feat, as_index=False)['SGY_JSK_QTY'].sum()
        df = pd.merge(curr_total, history_total, on=feat)
        df.columns = [feat] + ['{}_{}'.format('SGY_JSK_QTY', run_date.strftime('%Y-%m-%d'))] + \
                     ['{}_{}'.format('SGY_JSK_QTY', history_day.strftime('%Y-%m-%d'))]
        df['CHANGE_AMOUNT'] = df[f"SGY_JSK_QTY_{run_date.strftime('%Y-%m-%d')}"] - \
                              df[f"SGY_JSK_QTY_{history_day.strftime('%Y-%m-%d')}"]
        df['CHANGE(%)'] = (df['CHANGE_AMOUNT'] / df[f"SGY_JSK_QTY_{run_date.strftime('%Y-%m-%d')}"]) * 100
        df['CHANGE(%)'] = df['CHANGE(%)'].round(2)
        df = df.sort_values('CHANGE(%)')
        df_neg = df.query('`CHANGE(%)` < 0')
        df_neg['CATEGORY'] = feat
        df_neg.rename(columns={feat: 'TITLE'}, inplace=True)
        result.append(df_neg)
        df_neg.to_csv(f"output/{feat}_{run_date.strftime('%Y-%m-%d')}.csv", index=False)
        top5 = sales_comparison_by_periods(sales_df=data,run_date=datetime(2023,4,15), type='prev_month',feats=['CATE', 'NIUKE_NM', 'NIOKURI_NM', 'JAN_CD'])
        cols_used = ['TITLE',f"SGY_JSK_QTY_{run_date.strftime('%Y-%m-%d')}", f"SGY_JSK_QTY_{history_day.strftime('%Y-%m-%d')}",'CHANGE_AMOUNT','CHANGE(%)','CATEGORY']
        top5.sort_values('CHANGE(%)').groupby('CATEGORY')[cols_used].head(1)
    result = pd.concat(result)
    return result
def top5_categories(sales_df: pd.DataFrame, run_date: datetime):
    if type == 'prev_day':
        history_day = run_date - pd.DateOffset(days=1)
    elif type == 'prev_week':
        history_day = run_date - pd.DateOffset(weeks=1)  
    elif type == 'prev_month':
        history_day = run_date - pd.DateOffset(months=1)
    elif type == 'prev_quarter':
        history_day = run_date - pd.DateOffset(months=3)
    elif type == 'prev_year':
        history_day = run_date - pd.DateOffset(years=1)
    else:
        raise ValueError('Invalid input')
    history_df = sales_df[(sales_df['SAGYO_UNYO_DATE'] == history_day) & (sales_df['CATE'] == '生活')]
    curr_df = sales_df[(sales_df['SAGYO_UNYO_DATE'] == run_date) & (sales_df['CATE'] == '生活')]
    history_total = history_df.groupby('NIOKURI_NM', as_index=False)['SGY_JSK_QTY'].sum()
    curr_total = curr_df.groupby('NIOKURI_NM', as_index=False)['SGY_JSK_QTY'].sum()
    df = pd.merge(curr_total, history_total, on='NIOKURI_NM', suffixes=('', '_prev')) 
    df['CHANGE_AMOUNT'] = df['SGY_JSK_QTY'] - df['SGY_JSK_QTY_prev']
    df['CHANGE(%)'] = (df['CHANGE_AMOUNT'] / df['SGY_JSK_QTY']) * 100
    df['CHANGE(%)'] = df['CHANGE(%)'].round(2)
    df = df.sort_values('CHANGE(%)')
    top5_categories = df.nsmallest(5, 'CHANGE(%)')
    return top5_categories
print(top5_categories)
# Visualization
# Total quantity chart from 2023-03-01 to 2023-04-30
df_total = data[(data['SAGYO_UNYO_DATE'] >= "2023-03-01") & (data['SAGYO_UNYO_DATE'] <= "2023-04-30")]
x = mpl.dates.date2num(df_total['SAGYO_UNYO_DATE'])
y = df_total['SGY_JSK_QTY']
x_new = np.linspace(x.min(), x.max(), 300)
a_BSpline = make_interp_spline(x, y)
y_new = a_BSpline(x_new)

plt.figure(figsize=(16, 6))
plt.plot(mpl.dates.num2date(x_new), y_new)
plt.xticks(rotation=45)
plt.xlabel('Date')
plt.ylabel('Total Quantity')
plt.title('Total Quantity over Time')
plt.show()

# Total quantity chart from 2021-06-01 to 2023-06-15
df = data.query('SAGYO_UNYO_DATE >= "2021-06-01" & SAGYO_UNYO_DATE <= "2023-06-15"')
x = mpl.dates.date2num(df_total['SAGYO_UNYO_DATE'])
y = df_total['SGY_JSK_QTY']
x_new = np.linspace(x.min(), x.max(), 300)
a_BSpline = make_interp_spline(x, y)
y_new = a_BSpline(x_new)

plt.figure(figsize=(20, 6))
plt.plot(mpl.dates.num2date(x_new), y_new)
plt.xticks(rotation=45)
plt.xlabel('The entire period of time')
plt.ylabel('Total Quantity')
plt.title('Total Quantity for the entire period of time')
plt.show()

# Top 10 NIOKURI_NM by Quantity
grouped_data = data.groupby('NIOKURI_NM')['SGY_JSK_QTY'].sum().nlargest(10).reset_index()
plt.figure(figsize=(10, 7))
sns.barplot(data=grouped_data, x='SGY_JSK_QTY', y='NIOKURI_NM')
plt.xlabel('Quantity')
plt.ylabel('NIOKURI_NM')
plt.title('Top 10 NIOKURI_NM by Quantity')
plt.show()

# Top 10 CATE by Quantity
grouped_data = data.groupby('CATE')['SGY_JSK_QTY'].sum().nlargest(10).reset_index()
plt.figure(figsize=(10, 6))
sns.barplot(data=grouped_data, x='SGY_JSK_QTY', y='CATE')
plt.xlabel('Quantity')
plt.ylabel('CATE')
plt.title('Top 10 CATE by Quantity')
plt.show()
