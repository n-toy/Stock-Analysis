# Stock-Analysis
Data Bootcamp

# Challenge


In the Module 2 Excel Macro-enabled Workbook is a refactored VBA script that has been refactored to be more efficient. The script is more efficient because it has been written to:
  1. Use arrays for the 12 tickers specified
  2. Uses a 'tickerIndex' that increments  when a new ticker is found (utilizing the data being sorted by ticker)
  3. There is only one loop reading through the data versus the original nested loop
  
Total Volume, Start Price, and End Price are all tracked using tickerIndex and stored into respective arrays. Once the for loop finishes iterating through the data a separate sheet is formatted and the data in the arrays is written for analysis. The format of the sheet has been constructed to be scalable if the ticker data was to be increased. We utlized row counts, and column counts to make the end format dymanic in the scenario where more than 12 tickers would be shown. 
