import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import xlsxwriter
import openpyxl

tickers = pd.read_excel('summary.xlsx', sheet_name='price_history')

start = '2023-02-23'
end = datetime.today().strftime('%Y-%m-%d')

df_new = pd.DataFrame()

for stock in list(tickers.columns):
    df_new[stock] = yf.download(stock, start=start, end=end)['Adj Close']

summary_file = 'Summary.xlsx'

with pd.ExcelWriter(summary_file, engine='openpyxl', mode='a') as writer:
    df_new.to_excel(writer, sheet_name='price_history_new')
