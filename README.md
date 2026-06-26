# Business Analytics Assignment – Part 1: Data Cleaning

## Project Overview

This project demonstrates a complete data cleaning workflow using Microsoft Excel. The objective was to identify and correct data quality issues in a retail orders dataset while maintaining complete documentation of every transformation performed. The cleaned dataset was then validated and summarized using quality reports and pivot table analysis.

## Repository Structure

```text
priyasingh_2511034_part1_data_cleaning/
├── README.md
├── data/
│   ├── raw_orders.xlsx
│   └── cleaned_orders.xlsx
├── outputs/
│   ├── cleaning_log.md
│   ├── data_quality_report.xlsx
│   └── pivot_summary.xlsx
└── screenshots/
    ├── raw_data_preview.png
    ├── cleaned_data_preview.png
    ├── pivot_summary_1.png
    └── pivot_summary_2.png
```

## Dataset Description

The dataset contains retail order transactions including customer information, product details, shipping information, sales, cost, profit, discounts, payment status, and order status.

### Major Fields
- Order ID
- Order Date
- Ship Date
- Customer Information
- Region, State, City
- Product Category
- Quantity
- Unit Price
- Discount
- Sales
- Cost
- Profit
- Payment Status
- Order Status

## Tools Used

- Microsoft Excel
- Excel Tables
- Pivot Tables
- Conditional Formatting
- GitHub

## Data Cleaning Performed

- Removed duplicate records
- Trimmed extra spaces
- Standardized text formatting
- Standardized date formats
- Corrected inconsistent values
- Handled missing values
- Validated discount values
- Verified sales and profit calculations
- Checked shipping date sequence
- Created data quality flags
- Applied business validation rules

## Business Rules Applied

- Ship Date must not be earlier than Order Date.
- Discount values must remain within the valid range.
- Sales and Profit calculations must be consistent.
- Payment Status and Order Status must follow valid business combinations.
- Duplicate Order IDs were removed.
- Text fields were standardized.

## Output Files

- data/raw_orders.xlsx
- data/cleaned_orders.xlsx
- outputs/cleaning_log.md
- outputs/data_quality_report.xlsx
- outputs/pivot_summary.xlsx

## Screenshots

- screenshots/raw_data_preview.png
- screenshots/cleaned_data_preview.png
- screenshots/pivot_summary_1.png
- screenshots/pivot_summary_2.png

## Project Summary

The raw dataset was successfully cleaned, validated, and transformed into a high-quality analytical dataset. All cleaning decisions were documented, business rules were enforced, and analytical summaries were created using Excel Pivot Tables.
