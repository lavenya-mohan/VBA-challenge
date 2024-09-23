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

+ The script loop through all the stocks data once and the following information displayed.
+ The script will sort the distinct ticker symbol in one column in column "I" with a column header "Ticker.
+ The script will excute Quarterly change from opening price at the beginning of a given year to the closing price at the end of that year, and put the value on "J" column. For this task the code added a conditional formatting that highlighted positive change in green and negative change in red.
>[!Note] 
>Conditional formatting is applied correctly and appropriately to the percent change column (Mentioned in Requirements)
+ The script also percent perform a change from opening price at the beginning of a given year to the closing price at the end of that year, and put the value on "K" column.
+ The total stock volume also genereated on "L" column.
+ Calculate and display the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
