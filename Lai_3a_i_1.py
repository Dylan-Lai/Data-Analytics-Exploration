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
def get_data():
    """
    Get stock data for a given ticker symbol from Yahoo Finance.
    Parameters:
    - ticker_symbol (Any): Ticker symbol of the stock.
    Returns:
    - Dict[str, Any]: Dictionary containing stock data.
    """
    ticker_symbol = input("Enter the Stock Ticker Symbol: ").strip().upper()
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
        
        print("Price:",regularMarketPrice)
        print("Change:",regularMarketChange)
        print("Change (%)",regularMarketChangePercent)

        stats_labels = [label.text for label in soup.select('div[data-testid="quote-statistics"] ul li span:first-child')]
        stats_values = [value.text for value in soup.select('div[data-testid="quote-statistics"] ul li span:last-child')]

        financial_data = dict(zip(stats_labels,stats_values))

        for label,value in financial_data.items():
            print(f"{label}: {value}")

        stock = {
            'Stock': ticker_symbol ,
            'Price':regularMarketPrice,
            'Change':regularMarketChange,
            'ChangePercent':regularMarketChangePercent,
            #Summary
            **financial_data
        }
        # Writing stock data to a JSON file
        with open('Lai_stock_holder_data.json', 'w') as f:
            json.dump(stock, f,indent=4)
        # Writing stock data to a CSV file with aligned values
        df = pd.DataFrame(list(stock.items()),columns=['Data','Value'])
        df.to_csv('Lai_stock_holder_data.csv', index=False)
        # Writing stock data to an Excel file 
        df.to_excel('Lai_stock_holder_data.xlsx', index=False)

        path = os.getcwd()
        print("Download links:")
        print(f"JSON: {path}\\Lai_stock_holder_data.json")
        print(f"CSV: {path}\\Lai_stock_holder_data.csv")
        print(f"XLSX: {path}\\Lai_stock_holder_data.xlsx")

        return stock
    else:
        print(f"Unexpected status code: {response.status_code}")
        return response.status_code

data = get_data()

