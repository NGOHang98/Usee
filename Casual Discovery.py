### Funtions ###
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import japanize_matplotlib
from datetime import datetime, timedelta
import os
from scipy.interpolate import make_interp_spline
import numpy as np
import matplotlib.dates as mpl_dates
import matplotlib as mpl


### Import Data ###
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


### Data Preprocessing ###
data.dropna(inplace=True)
data['SAGYO_UNYO_DATE'] = pd.to_datetime(data['SAGYO_UNYO_DATE'], format='%Y%m%d')
data['SGY_TM'] = pd.to_datetime(data['SGY_TM'], format='%Y/%m/%d %H:%M:%S')


### Data Analysis ###

##Top Most Likely Causes ##
def logistic_comparison_by_periods(logistic_df: pd.DataFrame, run_date: datetime, type: str, feats):
    if type=='prev_day':
        history_day = run_date - pd.DateOffset(days=1)
    elif type=='prev_week':
        history_day = run_date - pd.DateOffset(weeks=1)
    elif type=='prev_month':
        history_day = run_date - pd.DateOffset(months=1)
    elif type=='prev_quarter':
        history_day = run_date - pd.DateOffset(months=3)
    elif type=='prev_year':
        history_day = run_date - pd.DateOffset(years=1)
    else: 
        raise ValueError('Invalid input')  
    result = []
    for feat in feats: 
        history_df = logistic_df[logistic_df['SAGYO_UNYO_DATE']==history_day]
        curr_df = logistic_df[logistic_df['SAGYO_UNYO_DATE']==run_date]
        history_total = history_df.groupby(feat, as_index=False)['SGY_JSK_QTY'].sum()
        curr_total = curr_df.groupby(feat, as_index=False)['SGY_JSK_QTY'].sum()
        df = pd.merge(curr_total, history_total, on=feat)
        df.columns = [feat] + ['{}_{}'.format('SGY_JSK_QTY', run_date.strftime('%Y-%m-%d'))] + ['{}_{}'.format('SGY_JSK_QTY', history_day.strftime('%Y-%m-%d'))]
        df['CHANGE_AMOUNT'] = df[f"SGY_JSK_QTY_{run_date.strftime('%Y-%m-%d')}"] - df[f"SGY_JSK_QTY_{history_day.strftime('%Y-%m-%d')}"] 
        df['CHANGE(%)'] = (df['CHANGE_AMOUNT'] / df[f"SGY_JSK_QTY_{run_date.strftime('%Y-%m-%d')}"]) * 100   
        df['CHANGE(%)'] = df['CHANGE(%)'].round(2)
        df = df.sort_values('CHANGE(%)')
        df_neg = df.query('`CHANGE(%)` < 0')
        df_neg['CATEGORY'] = feat
        df_neg.rename(columns={feat:'TITLE'}, inplace=True)
        result.append(df_neg)
        df_neg.to_csv(f"output/{feat}_{run_date.strftime('%Y-%m-%d')}.csv", index=False)
        cols_used = ['TITLE',f"SGY_JSK_QTY_{run_date.strftime('%Y-%m-%d')}", f"SGY_JSK_QTY_{history_day.strftime('%Y-%m-%d')}",'CHANGE_AMOUNT','CHANGE(%)','CATEGORY']
    top5 = pd.concat(result)
    sorted_top5 = top5.sort_values('CHANGE(%)').groupby('CATEGORY')[cols_used].head(1)
    print(sorted_top5)
    return sorted_top5
top5 = logistic_comparison_by_periods(logistic_df=data, run_date=datetime(2023, 4, 15), type='prev_month', feats=['CATE', 'NIUKE_NM', 'NIOKURI_NM'])

## Exploring 生活 in CATE by NIUKE_NM ##
def logistic_comparison_by_periods(logistic_df: pd.DataFrame, run_date: datetime, type: str, feats):
    if type=='prev_day':
        history_day = run_date - pd.DateOffset(days=1)
    elif type=='prev_week':
        history_day = run_date - pd.DateOffset(weeks=1)
    elif type=='prev_month':
        history_day = run_date - pd.DateOffset(months=1)
    elif type=='prev_quarter':
        history_day = run_date - pd.DateOffset(months=3)
    elif type=='prev_year':
        history_day = run_date - pd.DateOffset(years=1)
    else: 
        raise ValueError('Invalid input')  
    result = []
    for feat in feats: 
        history_df = logistic_df[(logistic_df['SAGYO_UNYO_DATE'] == history_day) & (logistic_df['NIOKURI_NM'] == feat)]
        curr_df = logistic_df[(logistic_df['SAGYO_UNYO_DATE'] == run_date) & (logistic_df['NIOKURI_NM'] == feat)]
        history_total = history_df.groupby('NIUKE_NM', as_index=False)['SGY_JSK_QTY'].sum()
        curr_total = curr_df.groupby('NIUKE_NM', as_index=False)['SGY_JSK_QTY'].sum()
        df = pd.merge(curr_total, history_total, on='NIUKE_NM', suffixes=('', '_prev')) 
        df['CHANGE_AMOUNT'] = df['SGY_JSK_QTY'] - df['SGY_JSK_QTY_prev']
        df['CHANGE(%)'] = (df['CHANGE_AMOUNT'] / df['SGY_JSK_QTY']) * 100
        df['CHANGE(%)'] = df['CHANGE(%)'].round(2)
        df = df.sort_values('CHANGE(%)')
        top5_categories = df.nsmallest(5, 'CHANGE(%)')
        result.append(top5_categories)
        top5_categories.to_csv(f"output/{feat}_{run_date.strftime('%Y-%m-%d')}.csv", index=False)  
    top5 = pd.concat(result)
    print(top5)
    return top5
# Example usage:
top5 = logistic_comparison_by_periods(logistic_df=data, run_date=datetime(2023, 4, 15), type='prev_month', feats=['近畿法人営業１課'])


### Visualization ###

# Total quantity chart from 2021-06-01 to 2023-06-15（Overview)
df = data.query('SAGYO_UNYO_DATE >= "2021-06-01" & SAGYO_UNYO_DATE <= "2023-06-15"')
df_total = data.groupby(['SAGYO_UNYO_DATE'], as_index=False)['SGY_JSK_QTY'].sum()
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
plt.savefig(f"output/total_quantity_overview_chart.jpg")
plt.close()  


# Total quantity chart for CATE
#Creat a time frame 
df_1 = data.query('SAGYO_UNYO_DATE=={}'.format('20230415'))
df_2 = data.query('SAGYO_UNYO_DATE=={}'.format('20230315'))
FEAT = ['CATE']
total2 = df_2.groupby(FEAT, as_index=False)['SGY_JSK_QTY'].sum()
total1 = df_1.groupby(FEAT, as_index=False)['SGY_JSK_QTY'].sum()
df = pd.merge(total1, total2, on=FEAT)
df.columns = FEAT + ['{}_{}'.format('SGY_JSK_QTY', '20230415')] + ['{}_{}'.format('SGY_JSK_QTY', '20230315')]
df['CHANGE AMOUNT'] = df['SGY_JSK_QTY_20230415'] - df['SGY_JSK_QTY_20230315']
df['CHANGE(%)'] = (df['CHANGE AMOUNT'] / df['SGY_JSK_QTY_20230415']) *100
df['CHANGE(%)'] = df['CHANGE(%)'].round(2)
df = df.sort_values('CHANGE(%)')
df1 = df.reset_index()
df_neg1 = df.query('`CHANGE(%)` < 0')
df_neg1['CATEGORY'] = 'CATE'
df_neg1.rename(columns={'CATE': 'TITLE'}, inplace=True)
df_neg1.head(5)
df_neg1.set_index('TITLE', inplace=True)
ax = df_neg1[['SGY_JSK_QTY_20230415', 'SGY_JSK_QTY_20230315']].plot(kind='bar', rot=0, figsize=(10, 6))
plt.xlabel('Category')
plt.ylabel('SGY_JSK_QTY')
plt.title('Comparison of SGY_JSK_QTY between 2023-04-15 and 2023-03-15 by Category')
plt.savefig(f"output/total_quantity_cate.jpg")
plt.close()  

# Total quantity chart for NIUKE_NM
grouped_data = data.groupby('NIUKE_NM')['SGY_JSK_QTY'].sum().reset_index()
sorted_data = grouped_data.sort_values('SGY_JSK_QTY', ascending=False).head(10)
plt.figure(figsize=(16,7))
sns.barplot(data=sorted_data, x='SGY_JSK_QTY', y='NIUKE_NM')
plt.xlabel('数量')
plt.ylabel('NIUKE_NM')
plt.savefig(f"output/total_quantity_NIUKE_NM.jpg")
plt.close()  

# Total quantity chart for the period of time
df = data[(data['SAGYO_UNYO_DATE']>="2023-03-01")&(data['SAGYO_UNYO_DATE']<="2023-04-30")]
df_total = df.groupby(['SAGYO_UNYO_DATE'], as_index=False)['SGY_JSK_QTY'].sum()
x = mpl.dates.date2num(df_total['SAGYO_UNYO_DATE'])  
y = df_total['SGY_JSK_QTY']
x_new = np.linspace(x.min(), x.max(), 300)
a_BSpline = make_interp_spline(x, y)
y_new = a_BSpline(x_new)
plt.plot(mpl.dates.num2date(x_new), y_new)
plt.xticks(rotation=45)
plt.figure(figsize=(16, 8))
plt.plot(df_total['SAGYO_UNYO_DATE'], df_total['SGY_JSK_QTY'], marker='o')
plt.xticks(rotation=45)
plt.savefig(f"output/total_quantity_period_of_time.jpg")
plt.close()  