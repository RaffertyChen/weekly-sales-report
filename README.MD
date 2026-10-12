# Weekly Sales Report

Turns a messy order spreadsheet into a clean sales report and emails it to the shop owner every Monday, so nobody has to clean or count orders by hand.

## What it does

- Cleans the messy order export: fixes prices, dates and names, and removes cancelled and duplicate orders
- Skips any row it can't read and names that order, instead of guessing
- Works out sales per customer and per month
- Asks AI (Claude) for a short summary of the results
- Builds a formatted Excel report and emails it with the summary
- Runs by itself every Monday on GitHub, and sends an alert email if the order file is missing

## The result

The practice file has 72 order rows. The script turns them into 63 clean orders worth $30,250.50, split by customer and by month, in an Excel report that lands in the owner's inbox.

## Check the AI's work

Python does all the maths. The AI only writes the summary, using the numbers it's given. If the summary and the report ever disagree, trust the report.

## How to run it

1. Install Python from python.org.
2. In a terminal in this folder, run: `python -m pip install -r requirements.txt`
3. Copy `.env.example` to a new file called `.env`, and add your Anthropic key(https://platform.claude.com/dashboard) and your Gmail details(https://myaccount.google.com/apppasswords).
4. Run: `python weekly_sales_report.py`

## Files

- `weekly_sales_report.py`: the script
- `orders_export.xlsx`: a practice order export full of messes (made up)
- `sales_report.xlsx`: the report the script made from it
- `.github/workflows/weekly_sales_report.yml`: tells GitHub to run it every Monday