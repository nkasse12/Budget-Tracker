# About This Project

## Inspiration

This project started in my **Introduction to Finance 1** class. One of our assignments was to build a personal budget in Google Sheets for a realistic future job. I chose a starting computer programmer in San Diego earning $125,000 a year and researched what life would actually cost: rent for a studio apartment, car insurance, gas, groceries, phone and internet plans, and savings for a 401(k).

For help with this assignment, I went straight to my father. I couldn't understand most of the terms in the template spreadsheet my teacher gave me and I wanted to have the most accurate information for my assignment. After all, I took this class to understand how personal financing works in the real world. I was shocked by the numbers my father told me. 
Most of his salary was cut off by tax, and he was spending an extreme amount of money on groceries, insurance, bills(power, electricity, water, etc), and so many other things. He earns much more money than the average software engineer in California, yet our family lived an extremely average middle-class lifestyle. Then I figured out the problem. He was spending much more money on these extra expenses than he was making. 

The assignment made me realize how quickly a big salary shrinks once you account for taxes and everyday expenses. I wanted to build a device that my father could use to keep track of his own money. I watch every day as he struggles to manage our expenses and I want to do everything possible to help him. 

## From Class Assignment to Real Tool

When I looked back at my original spreadsheet, I noticed a few things I wanted to improve:

- It was a plan, not a tracker. It showed what I expected to spend, but there was no way to record what I actually spent. 
- Numbers were typed in by hand. One annual total didn't match its monthly amount (gifts: $623 × 12 should be $7,476, not $7,070), which threw off the yearly totals.
- Tax was listed as an expense. It makes more sense to subtract tax from gross income first, so you budget from your net income. Net income is the money that is actually taken home, the salary after tax. 

So I rebuilt it using Python, with every total calculated by formulas so mistakes like that can't happen.

## How It Works

The script (`budget_tracker.py`) generates an Excel workbook with three tabs:

1. **Budget:** Enter your salary and tax rate, and it calculates your net monthly income. Each spending category gets a monthly budget, and the annual amount is calculated automatically. A "Left to Assign" line checks that every dollar of income has a purpose.
2. **Transactions:** A log of real purchases. Categories come from a dropdown so they always match the budget.
3. **Dashboard:** Pick any month to see budgeted vs. actual spending for each category, with a chart. Categories turn red when you go over budget.

## Finance Concepts Used

- **Gross vs. net income:** - Gross Income is the salary listed on paper wheras Net Income is the salary you actually take home after tax deductions.
- **Fixed vs. variable expenses:** Fixed Expenses are expenses that cannot be avoided(such as bills) whereas Variable Expenses are expenses that can be adjusted(personal expenses like shopping, going out, and so on)
- **The 50/30/20 rule:** The 50/30/20 rule talks about the best budgeting method, which is spending 50% of your Net Income on **Needs**, 30% on **Wants**, and 20% on **Savings**. My sample budget came out to about 53% / 15% / 32%, which is savings-heavy.
- **Short-term vs. long-term savings:** Short-term savings used for emergencies or extra money when needed, they're usually things like Emergency Funds. Long-term savings are money going towards retirement, like a 401(k). 
- **Pay yourself first:** Treating savings as a budget category, not whatever is left over.

## What I Learned

- How to use Excel formulas like `SUMIFS`, `EDATE`, and `IFERROR` to make a spreadsheet update itself
- How to generate a formatted spreadsheet with Python and the `openpyxl` library
- How to use Git and GitHub to share and track a project

## Future Plans

- Import transactions directly from a bank CSV file
- Automatically suggest categories for transactions
- Add a savings goal tracker
- Create a student version with categories like allowance, part-time job income, and school expenses(could be useful for high school and college students)

## Author

Created by Nitya Hamsini Kasse, inspired by my Introduction to Finance 1 class.
