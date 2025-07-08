import yfinance as yf
import pandas as pd

# 🎯 Tickers to track
tickers = ['QUICKHEAL', 'MSFT', 'GOOGL', 'TSLA', 'NVDA', 'META', 'AMZN', 'AMD', 'NFLX', 'CRM']

# 📅 Date range
start_date = '2024-04-11'
end_date = '2024-05-11'

# 📊 Download data for all tickers
data = yf.download(tickers, start=start_date, end=end_date, auto_adjust=True)['Close']

# 🧼 Drop rows with all missing values
data.dropna(how='all', inplace=True)

# ✅ Display the table
print(f"\n📅 Daily closing prices from {start_date} to {end_date}")
print(data)
