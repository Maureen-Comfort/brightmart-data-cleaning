# BrightMart Orders: Data Cleaning Project (Excel)

Cleaning a messy online-store orders file so the sales numbers can be trusted.

**Author:** Maureen Giramiya | **Tools:** Excel (Power Query, formulas, pivot tables)

## Overview
BrightMart is a fictional online store in Uganda. Its staff typed 524 orders by hand, and management wanted accurate sales figures. The raw file had duplicate orders, text dates in several formats, inconsistent spellings, missing values and impossible numbers. My job was to profile the problems, fix what could be fixed, flag what could not, and prove the result.

## Dataset
- 524 rows, 14 columns (orders, customers, products, prices, dates, payment and status etc).
- The dataset was created for practice, so no real customers are involved.
- A data dictionary in the workbook describes what "clean" looks like for each column.

## Problems found (30 problem types across 13 columns)
| Area | Examples |
|---|---|
| Duplicates | 24 repeated orders |
| Dates | 415 text dates in 4 formats, 2 blanks, 7 impossible years (2026, 2035, 2052) |
| Names and emails | Extra spaces, wrong capitalization, missing @ |
| Phone numbers | 5 different formats, 6 too short, 10 blank |
| Categories, cities, payment, status | Misspellings, abbreviations (Kla), trailing spaces, mixed case |
| Quantities and prices | Text ("two", "UGX 25,000"), blanks, zero, negative, 10 times too high |
| Totals | Not equal to quantity x price |
| Delivery dates | Stored as text, some earlier than the order date |

## Approach
1. **Profile:** checked every column for blanks, unique values, data type, odd values and hidden spaces. Logged each problem in a cleaning log.
2. **Clean:** Power Query removed duplicates and converted dates. Formulas (TRIM, PROPER, LOWER, VLOOKUP, SUBSTITUTE) fixed text, phones and categories. Lookup tables mapped spelling variants to standard values.
3. **Recover or flag:** missing quantities and prices were recovered from related columns where possible. Anything that could not be recovered was flagged, never invented.
4. **Validate:** 15 tests check the final data against the data dictionary.
5. **Prove impact:** before and after pivot tables and charts.

## Results
| Measure | Raw | Clean |
|---|---|---|
| Rows | 524 | 500 |
| Total sales (UGX) | 77,080,000 | 73,550,000 |
| Order dates stored as text | 415 | 0 |
| Category labels in a pivot table | 11 | 4 |

The raw Total_Amount column overstated sales by **UGX 3,530,000 (about 4.8%)**, mostly from duplicate orders and incorrect totals. Messy category labels also split revenue across 11 lines instead of 4.

All **15 validation tests pass** (unique IDs, valid dates, standard spellings, totals equal quantity x price, and more).

## Decisions and limits
- 6 impossible order dates and 8 delivery dates earlier than the order date were set to blank and flagged, because the true dates are unknown.
- 2 orders have no order date. They were kept and flagged so their sales still count.
- 18 emails (missing or invalid) and 16 phone numbers (blank or too short) were left blank or marked invalid rather than guessed.
- 16 quantities were estimated as total / unit price and marked "Recovered". A few may be wrong, so they are flagged in the Qty_Source column.
- 62 blank delivery dates are valid: they are all cancelled orders.

## Files
- `BrightMart_Orders_CLEAN.xlsx`: Raw_Data, Clean_Data, Lookups, Validation, Impact, Cleaning_Log and pivot charts.
- `images/`: before(Messy Chart) and after(Clean Chart) screenshots.

## What I learned
- Steps to take while cleaning data.
- What mistakes in every column to look out for.
- I learnt advanced formulas for example SUBSTITUTE
  
