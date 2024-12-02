# Cryptocurrency Market Tracker 📈

A real-time cryptocurrency tracking system that monitors the top 50 cryptocurrencies, performs market analysis, and generates detailed reports.

## Features 🚀

- **Live Data Tracking**
  - Fetches real-time data for the top 50 cryptocurrencies
  - Updates every 5 minutes automatically
  - Includes price, market cap, volume, and 24h changes

- **Professional Excel Output**
  - Beautifully formatted spreadsheet
  - Color-coded price changes (green for gains, red for losses)
  - Auto-adjusted column widths
  - Professional header styling
  - Separate analysis sheet

- **Comprehensive Analysis**
  - Market overview statistics
  - Top 5 cryptocurrencies by market cap
  - Price movement analysis
  - Trading volume insights
  - Market share calculations

- **Detailed Reporting**
  - Generates comprehensive analysis reports
  - Includes key market metrics
  - Time-stamped data points
  - Easy-to-read formatted output

## Prerequisites 📋

- Python 3.7 or higher
- pip (Python package installer)

## Installation 🔧

1. Clone the repository:
   ```bash
   git clone https://github.com/AshwinSaklecha/cryptoTracker.git
   cd cryptoTracker

2. Install required packages:
   ```bash
   pip install -r requirements.txt

## Usage 💻

1. Start the cryptocurrency tracker (keeps running and updating Excel):
   ```bash
   python crypto_tracker.py

2. In a separate terminal, generate an analysis report:
   ```bash
   python generate_report.py

## Output Files 📁

- **crypto_data.xlsx**: Live-updating Excel file with current market data
- **crypto_analysis_report.txt**: Detailed market analysis report

## Data Sources 🌐

- CoinGecko API (Free tier)
- Updates every 5 minutes
- Top 50 cryptocurrencies by market capitalization

## Technical Details 🔧

The project uses:
- `pandas` for data manipulation
- `openpyxl` for Excel formatting
- `requests` for API calls
- Real-time data processing
- Automated report generation


## License 📄

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments 🙏

- CoinGecko API for providing cryptocurrency data
- Python community for excellent libraries

## Author ✨

Ashwin Saklecha

---

*Note*: This project is for educational purposes and should not be used as financial advice.
