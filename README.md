# ETF & Stock Comparator

A Python tool that analyzes and compares ETFs and individual stocks using live data from Yahoo Finance. Automatically generates a formatted Excel report with charts.

## Modes
- **ETF Mode** — analyze and compare any ETFs
- **Stock Mode** — analyze and compare any individual stocks

## ETF Mode features
- Pulls live top 10 holdings for any ETFs you choose
- Calculates overlap between each pair of ETFs by stock count and weight
- Breaks down sector allocation for each ETF
- Generates a bar chart comparing sector allocations across all ETFs
- 1 year price performance chart normalized to 100
- Exports a formatted Excel report with tabs: Summary, Holdings, Overlap, Sectors, Charts

## Stock Mode features
- Key metrics — price, market cap, 52 week high/low, beta, dividend yield
- Valuation — P/E, forward P/E, PEG, price to book, EV/EBITDA
- Financials — revenue, margins, EPS, debt to equity, return on equity
- Analyst ratings breakdown with stacked bar chart
- 1 year price performance chart normalized to 100

## How to run it
1. Clone the repo
2. Install dependencies:
   pip install pandas openpyxl yfinance matplotlib
3. Run the script:
   python3 etf.py
4. Choose ETF or Stock mode when prompted
5. Enter tickers separated by commas (e.g. SPY,QQQ or AAPL,NVDA,JPM)
6. Open the generated Excel file to view the report

## Tech used
- Python, pandas, yfinance, openpyxl, matplotlib
