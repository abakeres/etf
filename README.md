# ETF Comparator

A Python tool that analyzes and compares ETF holdings, overlap, and sector allocation using live data from Yahoo Finance. Automatically generates a formatted Excel report with charts.

## What it does
- Pulls live top 10 holdings for any ETFs you choose
- Calculates overlap between each pair of ETFs by stock count and weight
- Breaks down sector allocation for each ETF
- Generates a bar chart comparing sector allocations across all ETFs
- Exports a formatted Excel report with 4 tabs: Summary, Holdings, Overlap, Sectors, and Charts

## How to run it

1. Clone the repo
2. Install dependencies:
   pip install pandas openpyxl yfinance matplotlib
3. Run the script:
   python3 etf.py
4. Enter ETF tickers when prompted (e.g. SPY,QQQ,VTI)
5. Open etf_comparison.xlsx to view the report

## Tech used
- python
- pandas
- yfinance
- openpyxl
