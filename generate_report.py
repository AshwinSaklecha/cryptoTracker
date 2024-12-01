from datetime import datetime
import pandas as pd

def generate_report(excel_file):
    """Generate an enhanced analysis report"""
    df = pd.read_excel(excel_file, sheet_name='Live Crypto Data')
    analysis_df = pd.read_excel(excel_file, sheet_name='Analysis')
    
    # Calculate additional metrics
    total_market_cap = df['Market Capitalization'].sum()
    top_5_market_share = (df.head(5)['Market Capitalization'].sum() / total_market_cap) * 100
    
    report = f"""
{'='*80}
                    CRYPTOCURRENCY MARKET ANALYSIS REPORT
{'='*80}

GENERATED ON: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

{'*'*80}
1. MARKET OVERVIEW
{'*'*80}
Total Cryptocurrencies Analyzed: {len(df)}
Total Market Capitalization   : $ {df['Market Capitalization'].sum():,.2f}
Average Cryptocurrency Price  : $ {df['Current Price (USD)'].mean():,.2f}
Top 5 Market Share           : {top_5_market_share:.2f}%

{'*'*80}
2. TOP 5 CRYPTOCURRENCIES BY MARKET CAP
{'*'*80}
{df.head(5)[['Cryptocurrency Name', 'Symbol', 'Market Capitalization', 'Price Change 24h (%)']].to_string(index=False)}

{'*'*80}
3. PRICE MOVEMENT ANALYSIS (24H)
{'*'*80}
Highest Price Change: {df['Price Change 24h (%)'].max():>8.2f}% ({df.loc[df['Price Change 24h (%)'].idxmax(), 'Symbol']})
Lowest Price Change : {df['Price Change 24h (%)'].min():>8.2f}% ({df.loc[df['Price Change 24h (%)'].idxmin(), 'Symbol']})
Average Change     : {df['Price Change 24h (%)'].mean():>8.2f}%

{'*'*80}
4. TRADING VOLUME ANALYSIS
{'*'*80}
Total 24h Trading Volume : $ {df['24h Trading Volume'].sum():,.2f}
Average Trading Volume   : $ {df['24h Trading Volume'].mean():,.2f}
Highest Volume          : $ {df['24h Trading Volume'].max():,.2f} ({df.loc[df['24h Trading Volume'].idxmax(), 'Symbol']})

{'='*80}
Note: All prices are in USD. Data sourced from CoinGecko API.
{'='*80}
"""
    
    with open('crypto_analysis_report.txt', 'w') as f:
        f.write(report)

if __name__ == "__main__":
    generate_report('crypto_data.xlsx') 