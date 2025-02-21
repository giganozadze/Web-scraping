import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import xlsxwriter
import openpyxl

tickers = pd.read_excel('indexes_list.xlsx')

start = '2023-03-02'
end = datetime.today().strftime('%Y-%m-%d')

df_new = pd.DataFrame()

for index in list(tickers.columns):
    df_new[index] = yf.download(index, start=start, end=end)['Adj Close']

df_new.to_excel('indexes_new.xlsx')

# indexes_file = 'indexes.xlsx'
#
# with pd.ExcelWriter(indexes_file, engine='openpyxl', mode='a') as writer:
#     df_new.to_excel(writer, sheet_name='price_history_new_indexes')
