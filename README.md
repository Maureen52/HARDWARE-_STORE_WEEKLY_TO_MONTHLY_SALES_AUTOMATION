# Hardware Store: Weekly-to-Monthly Sales Automation (Excel VBA)

An Excel VBA project that turns four weeks of messy hardware-store sales sheets into a clean monthly record, a set of data-quality checks, pivot tables and an interactive dashboard, all with one click.



## Why I built this

A busy hardware store sells cement, roofing, steel, plumbing, paint and more, and records sales every week. Turning those weekly sheets into a monthly report by hand means repeating the same cleaning, checking and pivoting every month, and that is where mistakes creep in. I wanted a process that runs the same way every time, shows what it changed, and is honest about what it could not trust.

## What it does

Running `RunMonthlyReport` does five steps:

| Step | What happens | Output sheet(s) |
|---|---|---|
| 1. Consolidate | Combines `Week1`...`WeekN` into one monthly sheet and checks the column headers | `Monthly_Raw` |
| 2. Clean | Fixes dates, numbers, spellings, categories and payment methods; removes blank and duplicate rows; fills missing prices and units from a product master; corrects x10 price typos | `Monthly_Clean` (Excel table), `Rejected_Rows`, `Cleaning_Log` |
| 3. Check consistency | Flags dates outside the month, weeks containing other weeks' dates, receipts with conflicting details, odd prices and unknown products | `Data_Checks` |
| 4. Pivot | Builds 8 pivot tables from one shared pivot cache | `Pivots` |
| 5. Dashboard | KPI cards, 6 pivot charts and 4 slicers (Week, Payment method, Salesperson, Category) | `Dashboard` |

Nothing is silently deleted: every rejected row is listed with the reason, and `Cleaning_Log` reconciles the row counts (rows read = blank + rejected + kept).

## Results on the sample data

The sample data is synthetic (see below) and contains deliberate errors so the cleaning has something to fix.

- 464 rows read, 8 blank rows removed, 28 rejected (11 duplicates, 17 invalid), 428 kept
- 123 text dates and 95 numbers stored as text converted
- 88 categories and 85 payment methods standardised
- 3 decimal-shift price typos corrected (for example 7,960 instead of 796)
- October 2026 revenue: KES 11,373,502 across 428 line items; the pivot total matches the clean-data total

## How to run it

Requires **Windows desktop Excel** (it uses `Scripting.Dictionary`; slicers need Excel 2013 or later).

1. Open `Hardware_Store_Workbook.xlsm` and enable macros.
2. Press `Alt+F8`, choose `RunMonthlyReport` and click Run (or run `SetupControlPanel` once to create buttons).
3. Read `Data_Checks` and `Cleaning_Log` before using the dashboard.

To use your own data: paste each week's sales into `Week1`...`Week4` using the same 10 columns (`Date, Receipt_No, Customer, Category, Product, Unit, Qty, Unit_Price, Payment_Method, Salesperson`), and update the `Lookups` sheet with your products and prices.

Other macros: `SaveMonthlyArchive` saves a dated copy, `StartNewMonth` clears the week sheets, `ResetSlicers` clears every slicer filter.

## Business rules and assumptions

These are written at the top of the VBA module and can be changed:

- Text dates are read as day/month/year.
- Rows with the same date, receipt, product, quantity and price are treated as duplicates and the second is removed.
- Rows with a missing or zero quantity, missing receipt, missing product or unreadable date are rejected.
- Negative quantities are treated as possible returns: not netted off revenue, but their value is reported in `Cleaning_Log`.
- A price about 10x or 1/10 of the standard price is treated as a decimal-shift typo and corrected; each correction is listed.
- Other unusual prices and dates outside the month are kept but flagged for review.

## Limitations

- Tested only on synthetic data, not on a real store's records.
- Windows Excel only. Week sheets must be named `Week1`, `Week2`... with a fixed 10-column layout.
- Receipts are not counted distinctly (pivot tables behind slicers can't do this), so the dashboard reports line items and average line value instead of receipts and average basket.
- The standard prices in `Lookups` are illustrative, not a real price list.

## About the data

All data is **synthetic**. I generated the four weeks of October 2026 sales, the customers and the prices for demonstration; none of it comes from a real business.

## How I built it

I designed the workflow and the business rules, tested the output, and wrote this documentation. I used an AI assistant (Claude) to help write and debug the VBA code. 

## Repository contents

- `Hardware_Store_Workbook.xlsm`: the workbook with sample data
- `modHardwareReport.bas`: the VBA module (readable without opening Excel)
- `Hardware_Store_Working`: the workbook where the Macro is run
- `Hardware_Report_2026-10`: Macro generated report
