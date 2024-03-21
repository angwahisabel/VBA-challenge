# VBA-challenge

Stoch Analysis Project


Overview

This project aims to analyze stock data across multiple worksheets within an Excel workbook. The provided script automates the process of generating a summary table containing summarized information such as yearly change, percent change, and total volume for each stock. It calculates various metrics for each stock and identifies stocks with the greatest percent increase, greatest percent decrease, and greatest total volume. By automating the analysis process, one can quickly gain insights into stock performance and identify trends without the need for manual calculations.



Usage

Open the Excel workbook containing the stock data. Ensure that your stock data is organized in an Excel workbook with each worksheet representing a different year (2018,2019,2020). The data should be structured with specific columns:
Column A: Ticker symbol
Column C: Opening price
Column F: Closing price
Column G: Total volume

Open VBA Editor
Insert a new module by right-clicking on the project in the Project Explorer pane and selecting Insert > Module.
Run the provided VBA script named "StockSummary". The script will automatically loop through each worksheet, calculate summary statistics, apply conditional formatting, and output the results.



Results

Once the script has finished running, it adds a summary table to each worksheet with the following columns:
Ticker
Yearly Change
Percent Change
Total Volume
Additionally, it identifies the stocks with the greatest percent increase, greatest percent decrease, and greatest total volume. These values are displayed in a separate table on each worksheet, with the following columns:
Ticker
Value (Greatest percent increase, percent decrease, and total volume)



Important Notes

This script assumes that the data for each stock is contiguous and does not contain any blank rows within each dataset.



References

Kelly, P. (2021, November 4). VBA for Loop - A complete Guide - Excel Macro Mastery. Excel Macro Mastery. https://excelmacromastery.com/vba-for-loop/
