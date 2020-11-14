This script was created to analyze over two million lines of stock market data in an Excel workbook by interating through all three of its worksheets and creating a summary table of relevant information next to the data originally provided. The summary table on each page calculates for each ticker symbol:
  - The ticker symbol
  - The change from the first opening price of the year to the closing price of the year, highlighted in green if the change was an increase and red if the change was a decrease
  - The stock's total volume
  
  As a bonus, this code creates a second summary table which shows the stocks which experienced the greatest increase, greatest decrease, and greatest total volume and their respective values in those categories.
  
  While it was not part of the original assignment, I chose to include a line of code that autofits each column's width to its widest piece of content in order to improve readability of the final chart.
