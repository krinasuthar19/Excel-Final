Project Overview

This project focuses on building a complete data analysis and reporting solution using Microsoft Excel. The analysis is performed on a dataset containing over 200 customer transaction records. The objective is to extract meaningful business insights, perform advanced analysis, and present results through interactive dashboards and visual storytelling.

Dataset Description

Source Sheet: Final Project Dataset
Records: 200+ customer transactions

Key Columns Include:

Customer ID

Customer Name

Product / Product Category

Transaction Date

Sales / Revenue

Quantity

Discount

Profit

Region

Project Objective

To analyze customer transaction data using advanced Excel functions, analytical tools, and dashboards in order to:

Identify sales trends and high-value customers

Measure revenue and profitability

Understand purchasing behavior

Support data-driven decision making

Workbook Structure

Raw Data

Contains the original dataset without modifications.

Used as the base for all calculations and analysis.

Analysis

Intermediate calculations and formulas.

Use of Excel functions, scenarios, and regression outputs.

Visualizations

Pivot tables and pivot charts.

Supporting charts for trends and comparisons.

Dashboard

Interactive dashboard with KPIs, slicers, and timelines.

Executive-level summary view.

Key Topics Implemented
Date & Time Functions

TODAY() and NOW() for dynamic date tracking.

DATEDIF() to calculate durations such as customer age or order age.

EOMONTH() for month-end analysis and grouping.

FILTER Function

Used to extract multiple records dynamically.

Applied to identify high-value customers, specific regions, and top-performing products.

Conditional Formatting

Applied icon sets and arrows to indicate:

Sales growth or decline

Profit performance

Highlighted top-performing customers and products.

Timestamp Creation

Created a live timestamp column using NOW().

Used iterative calculation to preserve update time when required.

WHAT-IF Analysis

Scenario Manager used to evaluate the impact of different discount levels on profit.

Goal Seek used to determine required sales or discount to achieve target profit.

Linear Regression

Used Data Analysis Toolpak to perform regression analysis.

Analyzed relationship between Sales and Profit.

Interpreted coefficients, R-square, and significance.

High-Value Customer Analysis

Used GroupBy logic with formulas such as:

SUMIFS

INDEX

MATCH

FILTER

Identified customers contributing maximum revenue.

Dashboard Creation

Built using:

Pivot Tables

Pivot Charts (Bar, Line, Pie)

Slicers for Region, Product, Date

Timelines for time-based filtering

Included KPI indicators for:

Total Revenue

Total Profit

Top Customer

Most Sold Product

Frequency & Matching Analysis

Identified most frequently purchased products using formulas.

Extracted repeating names and products.

Compared two lists and extracted matching values in a new column.

TEXT Functions

Created abbreviations using:

LEFT

RIGHT

MID

UPPER

LOWER

Standardized names and codes for reporting.

Insights & Storytelling

A small percentage of customers generate a large share of revenue.

Discounts significantly impact profitability beyond a certain threshold.

Certain products and regions consistently outperform others.

Sales and profit show a strong positive correlation.

These insights are presented through dashboards and supported by visual and numerical evidence.

Deliverables

Excel workbook with:

Raw Data

Analysis Sheets

Visualizations

Interactive Dashboard

Summary insights and KPIs

This README file documenting the project

Conclusion

This project demonstrates strong proficiency in Excel-based data analysis, advanced formulas, business analytics, and dashboard design. The final solution provides actionable insights and an effective reporting framework suitable for real-world business scenarios.
