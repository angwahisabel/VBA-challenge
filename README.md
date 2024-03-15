# VBA-challenge

Module 2 Challenge
Stock Summary VBA Script


Overview

This VBA script, is designed to analyze stock data in multiple worksheets within an Excel workbook. It calculates various metrics for each stock, including yearly change, percent change, and total volume. Additionally, it identifies the stocks with the greatest percent increase, greatest percent decrease, and greatest total volume.


Usage

Open Excel Workbook: Open the Excel workbook containing the stock data.
Open VBA Editor: Press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
Insert Module: Insert a new module by right-clicking on the project in the Project Explorer pane and selecting Insert > Module.
Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
The total stock volume of the stock

Run Script: Close the VBA editor and run the StockSummary subroutine either by pressing F5 or navigating to Run > Run Sub/UserForm.



Output

The script adds a summary table to each worksheet with the following columns:
Ticker
Yearly Change
Percent Change
Total Volume
Additionally, it identifies the stocks with the greatest percent increase, greatest percent decrease, and greatest total volume. These values are displayed in a separate table at the beginning of each worksheet, with the following columns:
Ticker
Value (either percent increase, percent decrease, or total volume)


Important Notes

Ensure that the stock data is organized in each worksheet with columns for ticker, opening price, closing price, and volume.
This script assumes that the data for each stock is contiguous and does not contain any blank rows within each dataset.
Before running the script, make sure to save your Excel workbook in a macro-enabled format (.xlsm).


References

Kelly, P. (2021, November 4). VBA for Loop - A complete Guide - Excel Macro Mastery. Excel Macro Mastery. https://excelmacromastery.com/vba-for-loop/
