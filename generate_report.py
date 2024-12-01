from datetime import datetime
import pandas as pd

def generate_report(excel_file):
    """Generate a simple analysis report"""
    df = pd.read_excel(excel_file, sheet_name='Live Crypto Data')
    analysis_df = pd.read_excel(excel_file, sheet_name='Analysis')
    
    report = f"""
Cryptocurrency Market Analysis Report
Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

1. Market Overview
-------------------
Total cryptocurrencies analyzed: {len(df)}
Total market capitalization: ${df['Market Capitalization'].sum():,.2f}
Average cryptocurrency price: ${df['Current Price (USD)'].mean():,.2f}

2. Top 5 Cryptocurrencies by Market Cap
--------------------------------------
{df.head(5)[['Cryptocurrency Name', 'Market Capitalization']].to_string()}

3. Price Changes (24h)
---------------------
Highest price change: {df['Price Change 24h (%)'].max():.2f}%
Lowest price change: {df['Price Change 24h (%)'].min():.2f}%
Average price change: {df['Price Change 24h (%)'].mean():.2f}%

4. Trading Volume
----------------
Total 24h trading volume: ${df['24h Trading Volume'].sum():,.2f}
Average trading volume: ${df['24h Trading Volume'].mean():,.2f}
"""
    
    with open('crypto_analysis_report.txt', 'w') as f:
        f.write(report)

if __name__ == "__main__":
    generate_report('crypto_data.xlsx') 