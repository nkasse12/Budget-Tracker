# About This Project

## Inspiration

This project started in my **Introduction to Finance 1** class. One of our assignments was to build a personal budget in Google Sheets for a realistic future job. I chose a starting computer programmer in San Diego earning $125,000 a year and researched what life would actually cost: rent for a studio apartment, car insurance, gas, groceries, phone and internet plans, and savings for a 401(k).

The assignment made me realize how quickly a big salary shrinks once you account for taxes and everyday expenses. It also made me want to turn the budget into something I could actually use in real life, not just turn in once.

## From Class Assignment to Real Tool

When I looked back at my original spreadsheet, I noticed a few things I wanted to improve:

- **It was a plan, not a tracker.** It showed what I *expected* to spend, but there was no way to record what I *actually* spent. Comparing the two is where budgeting really helps.
- **Numbers were typed in by hand.** One annual total didn't match its monthly amount (gifts: $623 × 12 should be $7,476, not $7,070), which threw off the yearly totals.
- **Tax was listed as an expense.** It makes more sense to subtract tax from gross income first, so you budget from your *net* income: the money you actually take home.

So I rebuilt it using Python, with every total calculated by formulas so mistakes like that can't happen.

## How It Works

The script (`budget_tracker.py`) generates an Excel workbook with three tabs:

1. **Budget:** Enter your salary and tax rate, and it calculates your net monthly income. Each spending category gets a monthly budget, and the annual amount is calculated automatically. A "Left to Assign" line checks that every dollar of income has a purpose.
2. **Transactions:** A log of real purchases. Categories come from a dropdown so they always match the budget.
3. **Dashboard:** Pick any month to see budgeted vs. actual spending for each category, with a chart. Categories turn red when you go over budget.

## Finance Concepts Used

- **Gross vs. net income:** budgeting from take-home pay, not salary
- **Fixed vs. variable expenses:** separating costs that stay the same (rent, insurance) from ones you can control (food, entertainment)
- **The 50/30/20 rule:** a guideline of 50% needs, 30% wants, and 20% savings. My sample budget came out to about 53% / 15% / 32%, which is savings-heavy.
- **Short-term vs. long-term savings:** an emergency fund vs. retirement (401k)
- **Pay yourself first:** treating savings as a budget category, not whatever is left over

## What I Learned

- How to use Excel formulas like `SUMIFS`, `EDATE`, and `IFERROR` to make a spreadsheet update itself
- How to generate a formatted spreadsheet with Python and the `openpyxl` library
- How to use Git and GitHub to share and track a project
- [Add anything else you learned here]

## Future Plans

- Import transactions directly from a bank CSV file
- Automatically suggest categories for transactions
- Add a savings goal tracker
- Create a student version with categories like allowance, part-time job income, and school expenses

## Author

Created by [Your Name], inspired by my Introduction to Finance 1 class.
