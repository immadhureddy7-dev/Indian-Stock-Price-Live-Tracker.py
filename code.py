import yfinance as yf
import pandas as pd
import openpyxl
import time
from datetime import datetime


FILE_NAME = "STOCK DETAILS.xlsx"

def get_stock_price(symbol):
    stock = yf.Ticker(symbol)
    data = stock.history(period="1d")
    if not data.empty:
        return data['Close'][-1]
    return None

print("📈 Indian Stock Price Tracker (Yahoo Finance)")
print("-------------------------------------------------")

company = input("Enter NSE Symbol (e.g., TCS.NS, RELIANCE.NS): ").upper()
records = []

while True:
    price = get_stock_price(company)
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if price:
        print(f"💹 {company} → ₹{price} at {current_time}")

        records.append({
            "Company": company,
            "Price": price,
            "Time": current_time
        })

        df = pd.DataFrame(records)
        df.to_excel(FILE_NAME, index=False)
        print("📁 Saved to Excel!")
    else:
        print("❌ Could not fetch live price.")

    time.sleep(5)
