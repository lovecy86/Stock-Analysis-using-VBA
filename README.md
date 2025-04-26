# **Quarterly Stock Analysis VBA Script**

## **Overview**
This VBA script analyzes stock market data across multiple worksheets, each representing a quarter. It processes stock data to calculate key metrics for each ticker symbol, applies conditional formatting, and identifies stocks with the greatest percentage increase, decrease, and total volume. The script is designed to automate repetitive tasks, making it efficient for large datasets like the alphabetical_testing.xlsx file.

## **Features**
**Ticker Symbol:** Extracts unique ticker symbols.
**Quarterly Change:** Calculates the difference between the opening price at the start of the quarter and the closing price at the end.
**Percentage Change:** Computes the percentage change based on the quarterly change.
**Total Stock Volume:** Sums the trading volume for each ticker.
**Conditional Formatting:** Highlights positive changes in green and negative changes in red.
**Greatest Metrics:** Identifies the stocks with the greatest percentage increase, decrease, and total volume.
**Multi-Sheet Processing:** Runs on all worksheets in the workbook with a single execution.

## **Prerequisites**
* Microsoft Excel with VBA enabled.
* The alphabetical_testing.xlsx dataset for testing.
* Basic knowledge of Excel and VBA to install and run the script.

## **Installation**
    * Open Excel and press Alt + F11 to access the VBA Editor.          
    * In the VBA Editor, go to Insert > Module to create a new module.
    * Copy and paste the contents of Stock_Analysis.vbs into the module.
    * Save the Excel file as a macro-enabled workbook (.xlsm).

## **VBA Scriptng**
* 1. Created a script that loops through all the stocks for each quarter and outputs the following information:
    * The ticker symbol.
    * Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
    * The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
    * The total stock volume of the stock. 
* 2. Added functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
* 3. Applied conditional formatting:
    * Green for positive quarterly changes
    * Red for negative quarterly changes

## **License**
This project is licensed under the MIT License.

