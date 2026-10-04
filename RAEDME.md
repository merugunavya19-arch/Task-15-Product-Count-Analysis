## Product Count Analysis – Excel
## Overview

This project analyzes the **Sample Superstore dataset** using Microsoft Excel to count products by category and identify the category with the highest product count.

## Objectives
Count products for each category.
Compare product counts across categories.
Identify the category with the highest product count.
Check the number of unique products.
## Dataset

## Sample Superstore Dataset

The analysis uses:

**Category** – Column O
**Product Name** – Column Q
## Excel Formulas Used
## Count Products by Category
=COUNTIF(RawData!O2:O9995,A2)

This formula counts records for each category.

## Top Category
=INDEX(A3:A5,MATCH(MAX(B3:B5),B3:B5,0))

This identifies the category with the highest product count.

## Highest Product Count
=MAX(B3:B5)
## Unique Products
=COUNTA(UNIQUE(RawData!Q2:Q9995))

This counts the unique product names in the dataset.

## Categories Analyzed
Furniture
Office Supplies
Technology
## Tools Used
Microsoft Excel
Sample Superstore Dataset
COUNTIF
INDEX
MATCH
MAX
COUNTA
UNIQUE
## Key Learning

This task helped strengthen practical skills in **Excel functions, categorical data analysis, product counting, and identifying the highest-performing category**.

## Conclusion

The Product Count Analysis provides a category-wise view of the Superstore product data and helps identify the category containing the highest number of unique products.

## Internship:
Veda Technology – Data Analytics
