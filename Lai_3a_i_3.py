import sys
sys.path.append("c:/users/dilly/appdata/local/programs/python/python313/lib/site-packages")
import requests
from bs4 import BeautifulSoup
import json
import csv
import sys
from typing import Any, Dict
import pandas as pd
import os
def get_data(ticker_symbol):
    """
    Get stock data for a given ticker symbol from Yahoo Finance.
    Parameters:
    - ticker_symbol (Any): Ticker symbol of the stock.
    Returns:
    - Dict[str, Any]: Dictionary containing stock data.
    """
    print('Getting stock data of ', ticker_symbol)    
    url = f'https://finance.yahoo.com/quote/{ticker_symbol}'
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36',
        'Accept-Language': 'en-US,en;q=0.9',
        'Referer': 'https://www.google.com/',
    }

    session = requests.Session()
    session.headers.update(headers)

    r = session.get(url)
    if r.status_code == 200:
        soup = BeautifulSoup(r.text,'html.parser')

        head = soup.find('div',attrs={'class':'price yf-k4z9w'})
        regularMarketPrice = head.find(attrs={'data-testid':'qsp-price'}).text.strip()
        regularMarketChange = head.find(attrs={'data-testid':'qsp-price-change'}).text.strip()
        regularMarketChangePercent = head.find(attrs={'data-testid':'qsp-price-change-percent'}).text.strip()

        stats_labels = [label.text for label in soup.select('div[data-testid="quote-statistics"] ul li span:first-child')]
        stats_values = [value.text for value in soup.select('div[data-testid="quote-statistics"] ul li span:last-child')]

        financial_data = dict(zip(stats_labels,stats_values))

        stock = {
            'Stock': ticker_symbol ,
            'Price':regularMarketPrice,
            'Change':regularMarketChange,
            'ChangePercent':regularMarketChangePercent,
            #Summary
            **financial_data
        }

        return stock
    else:
        print(f"Unexpected status code: {response.status_code}")
        return response.status_code


def main():
    ticker_symbols = []
    while True:
        ticker_symbol = input("Enter the Stock Ticker Symbol(or type 'STOP' to finish): ").strip().upper()
        if ticker_symbol == 'STOP':
            break
        ticker_symbols.append(ticker_symbol)
    
    stockdata = [get_data(symbol) for symbol in ticker_symbols]

    # Writing stock data to a JSON file
    with open('Lai_stock_profile_data.json', 'w') as f:
        json.dump(stockdata, f,indent=4)
    # Writing stock data to a CSV file with aligned values
    df = pd.DataFrame(stockdata)
    df.to_csv('Lai_stock_profile_data.csv', index=False)
    # Writing stock data to an Excel file 
    df.to_excel('Lai_stock_profile_data.xlsx', index=False)

    path = os.getcwd()
    print("Download links:")
    print(f"JSON: {path}\\Lai_stock_profile_data.json")
    print(f"CSV: {path}\\Lai_stock_profile_data.csv")
    print(f"XLSX: {path}\\Lai_stock_profile_data.xlsx")

main()

