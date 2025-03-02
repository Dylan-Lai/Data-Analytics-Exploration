import sys
sys.path.append("c:/users/dilly/appdata/local/programs/python/python313/lib/site-packages")
import requests
from bs4 import BeautifulSoup
import json
import csv
import sys
from typing import Any, Dict
import pandas as pd
def get_data(ticker_symbol: Any) -> Dict[str, Any]:
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
    soup = BeautifulSoup(r.text,'html.parser')

    head = soup.find('div',attrs={'class':'price yf-k4z9w'})
    regularMarketPrice = head.find(attrs={'data-testid':'qsp-price'}).text.strip()
    regularMarketChange = head.find(attrs={'data-testid':'qsp-price-change'}).text.strip()
    regularMarketChangePercent = head.find(attrs={'data-testid':'qsp-price-change-percent'}).text.strip()
    stock = {
        'ticker': ticker_symbol ,
        'stock_name': soup.find(attrs={'class':'yf-xxbei9'}).text.strip(),
        'regularMarketPrice':regularMarketPrice,
        'regularMarketChange':regularMarketChange,
        'regularMarketChangePercent':regularMarketChangePercent,
        'quote': head.find(attrs={'data-testid':'qsp-post-price'}).text.strip(),
        #Summary
        'previous_close': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[1].text.strip(),
        'open_value': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[3].text.strip(),
        'bid': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[5].text.strip(),
        'ask': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[7].text.strip(),
        'days_range': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[9].text.strip(),
        'week_range': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[11].text.strip(),
        'volume': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[13].text.strip(),
        'avg_volume': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[15].text.strip(),
        'market_cap': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[17].text.strip(),
        'beta': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[19].text.strip(),
        'pe_ratio': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[21].text.strip(),
        'eps': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[23].text.strip(),
        'earnings_date': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[25].text.strip(),
        'dividend_yield': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[27].text.strip(),
        'ex_dividend_date': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[29].text.strip(),
        'year_target_est': soup.find('div', {'data-testid':'quote-statistics'}).find_all('span')[31].text.strip()
    }
    return stock

# Check if ticker symbols are provided as command line arguments
if len(sys.argv) < 2:
    print("Usage: python script.py <ticker_symbol1> <ticker_symbol2> ...")
    sys.exit(1)
# Extract ticker symbols from command line arguments
ticker_symbols = sys.argv[1:]
# Get stock data for each ticker symbol
stockdata = [get_data(symbol) for symbol in ticker_symbols]

# Writing stock data to a JSON file
with open('stock_data.json', 'w', encoding='utf-8') as f:
    json.dump(stockdata, f)
# Writing stock data to a CSV file with aligned values
CSV_FILE_PATH = 'stock_data.csv'
with open(CSV_FILE_PATH, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = stockdata[0].keys()
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(stockdata)
# Writing stock data to an Excel file
EXCEL_FILE_PATH = 'stock_data.xlsx'
df = pd.DataFrame(stockdata)
df.to_excel(EXCEL_FILE_PATH, index=False)

print("DONE!")
