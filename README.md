# LagosScent Perfumes Ltd — Sales Data Cleaning & Analysis

A data cleaning and analysis project using Microsoft Excel, built on a simulated 100-row sales
dataset for a fictional Nigerian fragrance company.

## Problem

The raw export mirrored common real-world data issues:
- Inconsistent capitalization (e.g. `IBRAHIM YUSUF`, `ibrahim eze`, `Ibrahim Eze`)
- Leading/trailing/double spaces in text fields
- A combined `Location` field (`"City, State"`) that needed splitting
- Order dates stored in three different formats
- Missing values in `Quantity` and `Unit Price`
- 10 exact duplicate order rows

## What I did

1. **Standardized text fields** using `TRIM`, `PROPER`, and `UPPER` on Customer Name, Payment
   Method, and Sales Rep.
2. **Split the combined Location column** into separate City and State columns using
   Text to Columns (comma delimiter).
3. **Normalized Order Date** into one consistent date format, re-parsing text-stored dates
   with Text to Columns' date-conversion option.
4. **Identified missing values** using Filters, and flagged affected rows in a Notes column
   rather than guessing values.
5. **Removed 10 duplicate rows** using Excel's Remove Duplicates tool, verifying the count
   matched expectations (90 unique rows remained).
6. **Calculated Total Sale** per order (`Quantity × Unit Price`), using an `IF(OR(...))`
   formula to leave the result blank instead of `0` for rows with missing data.
7. **Built two PivotTables**:
   - Total Sales by Product Category
   - Total Sales by Sales Rep

## Key insight

Men's Fragrance was the top-performing category by revenue (₦1,002,000), outselling
Women's Fragrance (₦706,500) and Unisex Fragrance (₦217,000) combined. Body Spray was the
weakest performer (₦124,000), consistent with it being the lowest-priced product line.

| Category | Total Sales |
|---|---|
| Men's Fragrance | ₦1,002,000 |
| Women's Fragrance | ₦706,500 |
| Diffuser | ₦432,000 |
| Unisex Fragrance | ₦217,000 |
| Body Spray | ₦124,000 |
| **Grand Total** | **₦2,481,500** |

## Tools used

Microsoft Excel — TRIM, PROPER, UPPER, Text to Columns, Filters, Remove Duplicates,
IF/OR formulas, PivotTables.

## Files

- `LagosScent_Mini_Practice.xlsx` — full workbook (Instructions, Raw Data, cleaned data,
  and PivotTable sheets)
- `screenshots/` — before/after and PivotTable screenshots

---
*This project uses a simulated dataset built for practice purposes.*

