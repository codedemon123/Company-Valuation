import yfinance as yf
import pandas as pd

# Define companies and tickers
companies = {
    "TCS.NS": "TCS",
    "INFY.NS": "Infosys",
    "HCLTECH.NS": "HCL Technologies",
    "WIPRO.NS": "Wipro",
    "LTIM.NS": "LTIMindtree",
    "TECHM.NS": "Tech Mahindra",
    "OFSS.NS": "Oracle Fin.Serv.",
    "PERSISTENT.NS": "Persistent Sys",
    "POLICYBZR.NS": "PB Fintech",
    "LTTS.NS": "L&T Technology"
}

# Store results
company_data = []

# Extract from yfinance
for ticker, name in companies.items():
    stock = yf.Ticker(ticker)
    info = stock.info
    
    try:
        market_cap = info.get("marketCap", 0)
        enterprise_value = info.get("enterpriseValue", 0)
        shares_outstanding = info.get("sharesOutstanding", 1)
        revenue = info.get("totalRevenue", 0)
        ebitda = info.get("ebitda", 0)
        net_income = info.get("netIncomeToCommon", 0)
        share_price = info.get("currentPrice", 0)

        # Derived metrics
        net_debt = enterprise_value - market_cap
        ev_rev = enterprise_value / revenue if revenue else None
        ev_ebitda = enterprise_value / ebitda if ebitda else None
        pe_ratio = share_price / (net_income / shares_outstanding) if net_income and shares_outstanding else None

        company_data.append({
            "Company": name,
            "Ticker": ticker,
            "Share Price": share_price,
            "Shares Outstanding (Cr)": round(shares_outstanding / 1e7, 2),
            "Market Cap (Cr)": round(market_cap / 1e7, 2),
            "Enterprise Value (Cr)": round(enterprise_value / 1e7, 2),
            "Net Debt (Cr)": round(net_debt / 1e7, 2),
            "Revenue (Cr)": round(revenue / 1e7, 2),
            "EBITDA (Cr)": round(ebitda / 1e7, 2),
            "Net Income (Cr)": round(net_income / 1e7, 2),
            "EV/Revenue": round(ev_rev, 2) if ev_rev else None,
            "EV/EBITDA": round(ev_ebitda, 2) if ev_ebitda else None,
            "P/E": round(pe_ratio, 2) if pe_ratio else None
        })

    except Exception as e:
        print(f"Error for {ticker}: {e}")

# Create and save Excel file
df = pd.DataFrame(company_data)
df.to_excel("Comparable_Company_Valuation.xlsx", index=False)
print("✅ Excel file with Net Debt created successfully.")
