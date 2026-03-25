# Automated-Financial-Data-Pipeline-FMP-API-Power-Query-VBA-
This project demonstrates automated financial data collection from FinancialModelingPrep API, transformation using Power Query (M), and refresh via VBA. The solution simulates a lightweight data pipeline for financial reporting automation.

## Tech Stack
- Excel
- Power Query (M language)
- VBA
- REST API (FinancialModelingPrep)
- JSON parsing
  
## Architecture
1. VBA macro triggers data refresh
2. Power Query calls FMP REST API
3. JSON response is parsed and transformed
4. Cleaned dataset is loaded into structured Excel tables
5. Output ready for analysis / reporting

## Business Value
- Eliminates manual financial data collection
- Reduces risk of human error
- Speeds up reporting process
- Easily scalable to multiple companies

## How to Run
1. Insert your API key and ticker symbol(AAPL, TSLA atd.)
2. Click Enter and you'll receive the company's three key financial metrics, broken down into three tables (balance sheet, cash flow, income statement)

