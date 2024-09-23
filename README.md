# VBA-challenge
VBA-Challenge to make a Stock Market Analysis.

# Background
This project is a VBA scripting to analyze real stock market data.

# Instruction
Create a script that loops through all the stocks for each quarter and outputs the following information:

+ The ticker symbol
+ Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
+ The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
+ The total stock volume of the stock
+ Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
+ Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.
+ Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

# Solution

+ The script loops through all the stock data once and the following information is displayed.
+ The script will sort the distinct ticker symbol in one column in column "I" with a column header "Ticker.
+ The script will execute a Quarterly change from the opening price at the beginning of a given year to the closing price at the end of that year, and put the value in the "J" column. For this task, the code added conditional formatting that highlighted positive changes in green and negative changes in red.
>[!Note] 
>Conditional formatting is applied correctly and appropriately to the percent change column (Mentioned in Requirements)
+ The script also performs a percent change from the opening price at the beginning of a given year to the closing price at the end of that year and puts the value in the "K" column.
+ The total stock volume is also generated in the "L" column.
+ Calculate and display the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
