from os import write
from typing import final
import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math

#importing lists of stocks
stocks = pd.read_csv('sp_500_stocks.csv')

#importing api_key
from secrets import IEX_CLOUD_API_TOKEN
symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
first_data = requests.get(api_url).json() #converted the data retreived into a json object

#storing price and market_cap
price = first_data['latestPrice']
market_cap = first_data['marketCap']

#adding stocks to Pandas DataFrame
my_columns = ['Ticker', 'Stock Price', 'Market Capatilization', 'Number Of Shares To Buy']
# final_dataframe = pd.DataFrame(columns = my_columns)

# for stock in stocks['Ticker']:
#     api_url = f'https://sandbox.iexapis.com/stable/stock/{stock}/quote/?token={IEX_CLOUD_API_TOKEN}'
#     data = requests.get(api_url).json()
#     final_dataframe = final_dataframe.append(
#         pd.Series(
#             [
#                 stock,
#                 data['latestPrice'],
#                 data['marketCap'],
#                 'N/A'
#             ],
#         index=my_columns),
#     ignore_index=True
#     )
    
def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks['Ticker'],100))
symbol_strings = []
for i in range(0,len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

final_dataframe = pd.DataFrame(columns=my_columns)

for symbol_string in symbol_strings[:1]:
    batch_api_call = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['marketCap'],
                    'N/A'
                ],
                index=my_columns
            ),
            ignore_index=True
        )

#calculating the number of shares to buy
portfolio_size = input('Enter the value of your portfolio:')
try:
    val = float(portfolio_size)
except ValueError:
    print('Thats not a number! \nPlease try again:')
    portfolio_size = input('Enter the value of your portfolio:')
    val = float(portfolio_size)

position_size = val/len(final_dataframe.index)
for i in range(0,len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number Of Shares To Buy'] = math.floor(position_size/final_dataframe.loc[i,'Stock Price'])

#save data in an excel sheet --------
writer = pd.ExcelWriter('recommended trades.xlsx', engine='xlsxwriter')  #initialising the excel writer object
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

#write formats to store the items in the xlxs sheet
string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

#applying formats to the xlxs file
column_formats = {
    'A':['Ticker', string_format],
    'B':['Stock Price', dollar_format],
    'C':['Market Capitalization', dollar_format],
    'D':['Number Of Shares To Buy', integer_format],
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}',18,column_formats[column][1])
    #overwriting the column headings
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()