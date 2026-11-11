# Budget Tracker

A Python script that builds a personal budgeting spreadsheet in Excel. 
- Creates an organizes system for financial budgeting
- Operates on a monthly basis(tracks expenses for one month)
- Requires manual input of transactions and expenses
- Useful for tracking purchases/transactions, keeping track of expenses, and estimating how much money you will have leftover at the end of the month. 

## What it creates
- **Dashboard:** budgeted vs. actual spending for any month, with a chart and red highlights when a category goes over budget
- **Budget:** income, tax, and a monthly plan for each category, plus a 50/30/20 (needs/wants/savings) check
- **Transactions:** a log of real purchases with a category dropdown

All totals are live Excel formulas (`SUMIFS`, `EDATE`, `IFERROR`), so the sheet updates as you add data.

## How to run
```
pip install openpyxl
python budget_tracker.py
```
This creates `Budget_Tracker.xlsx` in the same folder. It opens in Excel or Google Sheets (upload to Drive).

## Customize
Edit the income values and the `cats` list in `budget_tracker.py`, then run it again.

## Sample data
The default numbers model a starting computer programmer in San Diego ($125,000/year). They're example figures, not real personal data.
