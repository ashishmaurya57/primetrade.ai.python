import requests
import pandas as pd
import time
from openpyxl import load_workbook

API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": "false"
}

def fetch_crypto_data():
    try:
        response = requests.get(API_URL, params=PARAMS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print("Error fetching data:", e)
        return None

def analyze_data(data):
    df = pd.DataFrame(data)[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]

    top_5 = df.nlargest(5, 'market_cap')
    avg_price = df['current_price'].mean()

    highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
    lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]

    return df, top_5, avg_price, highest_change, lowest_change

def update_excel(df, top_5, avg_price, highest_change, lowest_change):
    file_name = "crypto_data.xlsx"
    
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="Live Data", index=False)

            # Save analysis summary
            analysis_df = pd.DataFrame({
                "Metric": ["Average Price", "Highest % Change (24h)", "Lowest % Change (24h)"],
                "Value": [avg_price, highest_change['price_change_percentage_24h'], lowest_change['price_change_percentage_24h']],
                "Currency": ["USD", highest_change['name'], lowest_change['name']]
            })
            analysis_df.to_excel(writer, sheet_name="Analysis Summary", index=False)

            top_5.to_excel(writer, sheet_name="Top 5 Coins", index=False)

        print("Excel file updated successfully!")

    except FileNotFoundError:
        
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name="Live Data", index=False)
            analysis_df.to_excel(writer, sheet_name="Analysis Summary", index=False)
            top_5.to_excel(writer, sheet_name="Top 5 Coins", index=False)

        print("New Excel file created and updated successfully!")

while True:
    crypto_data = fetch_crypto_data()
    if crypto_data:
        df, top_5, avg_price, highest_change, lowest_change = analyze_data(crypto_data)
        update_excel(df, top_5, avg_price, highest_change, lowest_change)
    time.sleep(300) 
