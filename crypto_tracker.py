import requests
import pandas as pd
import time
from datetime import datetime
import os

class CryptoTracker:
    def __init__(self):
        self.base_url = "https://api.coingecko.com/api/v3"
        self.excel_file = "crypto_data.xlsx"
        
    def fetch_top_50_crypto(self):
        """Fetch top 50 cryptocurrencies data from CoinGecko"""
        try:
            endpoint = f"{self.base_url}/coins/markets"
            params = {
                'vs_currency': 'usd',
                'order': 'market_cap_desc',
                'per_page': 50,
                'page': 1,
                'sparkline': False
            }
            
            response = requests.get(endpoint, params=params)
            response.raise_for_status()
            return response.json()
            
        except requests.RequestException as e:
            print(f"Error fetching data: {e}")
            return None

    def process_crypto_data(self, data):
        """Process the raw cryptocurrency data into a pandas DataFrame"""
        if not data:
            return None
            
        df = pd.DataFrame(data)
        df = df[[
            'name',
            'symbol',
            'current_price',
            'market_cap',
            'total_volume',
            'price_change_percentage_24h'
        ]]
        
        df.columns = [
            'Cryptocurrency Name',
            'Symbol',
            'Current Price (USD)',
            'Market Capitalization',
            '24h Trading Volume',
            'Price Change 24h (%)'
        ]
        
        return df

    def analyze_data(self, df):
        """Perform analysis on the cryptocurrency data"""
        if df is None or df.empty:
            return None
            
        analysis = {
            'Top 5 by Market Cap': df.head(5)[['Cryptocurrency Name', 'Market Capitalization']],
            'Average Price': df['Current Price (USD)'].mean(),
            'Highest 24h Change': df.nlargest(1, 'Price Change 24h (%)')[['Cryptocurrency Name', 'Price Change 24h (%)']],
            'Lowest 24h Change': df.nsmallest(1, 'Price Change 24h (%)')[['Cryptocurrency Name', 'Price Change 24h (%)']],
        }
        
        return analysis

    def update_excel(self, df, analysis):
        """Update Excel file with latest data and analysis"""
        try:
            with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='w') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Live Crypto Data', index=False)
                
                # Write analysis
                analysis['Top 5 by Market Cap'].to_excel(writer, sheet_name='Analysis', startrow=1, index=False)
                pd.DataFrame({
                    'Average Price (USD)': [analysis['Average Price']]
                }).to_excel(writer, sheet_name='Analysis', startrow=8, index=False)
                
                analysis['Highest 24h Change'].to_excel(writer, sheet_name='Analysis', startrow=11, index=False)
                analysis['Lowest 24h Change'].to_excel(writer, sheet_name='Analysis', startrow=14, index=False)
                
            print(f"Excel file updated at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
        except Exception as e:
            print(f"Error updating Excel file: {e}")

    def run(self, update_interval=300):
        """Run the crypto tracker with specified update interval (in seconds)"""
        while True:
            # Fetch and process data
            raw_data = self.fetch_top_50_crypto()
            df = self.process_crypto_data(raw_data)
            
            if df is not None:
                # Perform analysis
                analysis = self.analyze_data(df)
                
                # Update Excel
                self.update_excel(df, analysis)
                
            # Wait for the specified interval
            time.sleep(update_interval)

if __name__ == "__main__":
    tracker = CryptoTracker()
    # Update every 5 minutes (300 seconds)
    tracker.run(update_interval=300) 