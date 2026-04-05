import pandas as pd
import yfinance as yf

user_input = input("Enter ETF tickers seperated by commas (e.g. SPY,QQQ,VTI):")

etfs = user_input.upper().split(",")

all_holdings = {}

for etf in etfs:
    ticker = yf.Ticker(etf)
    holdings = ticker.funds_data.top_holdings
    holdings = holdings.reset_index()
    holdings.columns = ["Ticker", "Name", "Weight"]
    all_holdings[etf] = holdings
    print(f"{etf} loaded - {len(holdings)} holdings found")
    for _, row in holdings.iterrows():
        print(f"    {row['Ticker']} - {row['Name']}: {row['Weight']:.1%}")

print("\nALL ETFS loaded successfully!")

print("\nOverlap Analysis:")

etf_list = list(all_holdings.keys())

for i in range(len(etf_list)):
    for j in range(i+1, len(etf_list)):
        etf1 = etf_list[i]
        etf2 = etf_list[j]

        tickers1 = set(all_holdings[etf1]["Ticker"])
        tickers2 = set(all_holdings[etf2]["Ticker"])

        overlap = tickers1.intersection(tickers2)
        overlap_count = len(overlap)

        weight1 = all_holdings[etf1][all_holdings[etf1]["Ticker"].isin(overlap)]["Weight"].sum()
        weight2 = all_holdings[etf2][all_holdings[etf2]["Ticker"].isin(overlap)]["Weight"].sum()

        print(f"{etf1} vs {etf2}: {overlap_count} stocks in common")
        print(f"    Overlap weight in {etf1}: {weight1:.1%}")
        print(f"    Overlap weigth in {etf2}: {weight2:.1%}")

print("\nSector Allocation:")

sector_data = {}

for etf in etfs:
    ticker = yf.Ticker(etf)
    sectors = ticker.funds_data.sector_weightings
    
    cleaned = {k.replace("_", " ").title(): round(v * 100, 2) for k, v in sectors.items()}
    sector_data[etf] = cleaned

    print(f"\n{etf} Sector Breakdown:")
    for sector, weight in sorted(cleaned.items(), key=lambda x: x[1], reverse=True):
        print(f"    {sector}: {weight}%")

print("\nExporting to Excel...")

with pd.ExcelWriter("etf_comparison.xlsx", engine="openpyxl") as writer:

    holdings_combined = pd.DataFrame()
    for etf in etfs:
        df = all_holdings[etf].copy()
        df["ETF"] = etf
        holdings_combined = pd.concat([holdings_combined, df])
    holdings_combined.to_excel(writer, sheet_name="Holdings", index=False)

    overlap_rows = []
    for i in range(len(etf_list)):
        for j in range(i+1, len(etf_list)):
            etf1 = etf_list[i]
            etf2 = etf_list[j]
            tickers1 = set(all_holdings[etf1]["Ticker"])
            tickers2 = set(all_holdings[etf2]["Ticker"])
            overlap = tickers1.intersection(tickers2)
            weight1 = all_holdings[etf1][all_holdings[etf1]["Ticker"].isin(overlap)]["Weight"].sum()
            weight2 = all_holdings[etf2][all_holdings[etf2]["Ticker"].isin(overlap)]["Weight"].sum()
            overlap_rows.append({
                "ETF Pair": f"{etf1} vs {etf2}",
                "Stocks in Common": len(overlap),
                f"Overlap Weight in {etf1}": f"{weight1:.1%}",
                f"Overlap Weight in {etf2}": f"{weight2:.1%}",
                "Shared Tickers": ", ".join(sorted(overlap))
            })
    overlap_df = pd.DataFrame(overlap_rows)
    overlap_df.to_excel(writer, sheet_name="Overlap", index=False)

    sector_df = pd.DataFrame(sector_data).fillna(0)
    sector_df.index.name = "Sector"
    sector_df.to_excel(writer, sheet_name="Sectors")

print("Done! File saved as etf_comparison.xlsx")
