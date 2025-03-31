import yfinance as yf
import pandas as pd
import os
import subprocess

# Define the company ticker (CPB for Campbell Soup, change to the company of choice)
ticker_symbol = "CPB"

# Fetch company data
company = yf.Ticker(ticker_symbol)

# Fetch historical financial data
income_statement = company.financials
balance_sheet = company.balance_sheet
cash_flow = company.cashflow

# Define the years to filter
years = ["2020", "2021", "2022", "2023", "2024"]


# Define a function to filter financial data by years
def filter_years(df, years):
    """Filter columns to match the given years."""
    if df is not None and not df.empty:
        # Ensure the index is in a datetime format
        df = df.T  # Transpose to have years as index
        df.index = pd.to_datetime(df.index, errors='coerce').strftime('%Y')  # Convert index to just year format
        # Ensure years exist in the index
        available_years = [year for year in years if year in df.index]
        if available_years:
            df = df.loc[available_years]  # Filter the dataframe to include only available years
        return df.T  # Transpose back to original format
    else:
        print("No data found for this dataset.")
        return pd.DataFrame()  # Return an empty DataFrame if no data exists


# Filter the financial data
income_statement_filtered = filter_years(income_statement, years)
balance_sheet_filtered = filter_years(balance_sheet, years)
cash_flow_filtered = filter_years(cash_flow, years)

# Additional components (simulated placeholders, can be filled with data)
# Add placeholders for company overview, ratios, CAPM, and valuation calculations
company_info = {
    "Company Name": ["Campbell Soup Company"],
    "Ticker Symbol": ["CPB"],
    "Industry": ["Food Products"],
    "Headquarters": ["Camden, NJ, USA"],
    "Fiscal Year End": ["July 31"]
}
company_info_df = pd.DataFrame(company_info)

# Simulated financial ratios (to be replaced with real data from Net Advantage)
ratios = {
    "Year": years,
    "Profit Margin": [5.2, 6.0, 6.5, 5.8, 6.1],
    "Cash Ratio": [0.3, 0.4, 0.5, 0.3, 0.4],
    "Avg. Collection Period (Days)": [32, 31, 34, 33, 30],
    "Days Sales in Inventory (Days)": [45, 50, 49, 48, 45],
    "Debt to Total Assets": [0.45, 0.47, 0.48, 0.46, 0.44],
    "EBIT/Interest Exp.": [8.5, 9.2, 8.9, 8.8, 9.0]
}
ratios_df = pd.DataFrame(ratios)

# Placeholder for CAPM & WACC (will use sample data for now)
capm_wacc = {
    "Company Beta": [1.2],
    "Risk-Free Rate": [3.2],  # Example: 10-year US Treasury bond yield
    "Market Risk Premium": [5.0],  # Assumed market risk premium
    "WACC": [8.0]  # Example: WACC estimated from the company
}
capm_wacc_df = pd.DataFrame(capm_wacc)

# Placeholder for Free Cash Flow & Firm Value
free_cash_flow = {
    "Year": years,
    "Free Cash Flow ($M)": [500, 550, 600, 450, 500],  # Example data
    "Growth Rate (%)": [5, 4, 6, -3, 2]  # Example growth rates
}
free_cash_flow_df = pd.DataFrame(free_cash_flow)

# Placeholder for Stock Valuation (using EPS and P/E to estimate stock price)
valuation = {
    "EPS ($)": [3.5, 3.7, 4.0, 3.6, 3.9],  # Example EPS data
    "Industry P/E": [15, 16, 15.5, 14, 15.2],  # Example P/E ratios
    "Stock Price Estimate ($)": [52.5, 59.2, 62.0, 54.4, 59.0]  # Calculated stock price estimate
}
valuation_df = pd.DataFrame(valuation)

# Define the file path
file_path = r"Comprehensive_stocks_2019_2024.xlsx"

# Save to Excel
try:
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        # Add company overview sheet
        company_info_df.to_excel(writer, sheet_name="Company Overview", index=False)

        # Financial data sheets
        income_statement_filtered.to_excel(writer, sheet_name="Income Statement", index=True)
        balance_sheet_filtered.to_excel(writer, sheet_name="Balance Sheet", index=True)
        cash_flow_filtered.to_excel(writer, sheet_name="Cash Flow", index=True)

        # Add financial ratios sheet
        ratios_df.to_excel(writer, sheet_name="Financial Ratios", index=False)

        # Add CAPM & WACC calculations
        capm_wacc_df.to_excel(writer, sheet_name="CAPM & WACC", index=False)

        # Add Free Cash Flow & Firm Value sheet
        free_cash_flow_df.to_excel(writer, sheet_name="Free Cash Flow & Firm Value", index=False)

        # Add Valuation & Pricing sheet
        valuation_df.to_excel(writer, sheet_name="Valuation & Pricing", index=False)

        # Add Recommendation sheet (leave empty for manual input later)
        recommendation = pd.DataFrame({"Recommendation": ["[Your Recommendation Here]"]})
        recommendation.to_excel(writer, sheet_name="Recommendation", index=False)

    print(f"Workbook created: {file_path}")

    # Open the workbook automatically
    try:
        os.startfile(file_path)  # This opens the file with the default application (Excel)
    except AttributeError:
        subprocess.run(["open", file_path], check=True)  # For MacOS
        # Add Linux handling if necessary

except Exception as e:
    print(f"Error saving Excel file: {e}")
