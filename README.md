# VBA Homework - The VBA of Wall Street

## Given Background:

"You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks."

### Files used for Scripting

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Brief Overview:

* VBA scripts were used to analyze real, historical stock market data!

### Part 1:
* A for-loop was created to iterate through each row of the data and provide a summary table with the following calculations:
* 1.)	A string indicating each unique stock ticker value
* 2.)	The change in stock value from open on first day of year to close on last day of year.
* 3.)	Applied conditional formatting to highlight positive vs negative change.
* 4.)	The percentage change of stock value from open on first day of year to close on last day of year.
* 5.)	Total annual stock volume.
* 6.)	Integrated into a for-loop to calculate deliverables across all worksheets in the workbook.

### Part 2:
* Built for-loops to iterate through a newly created table to create an additional summary table that highlights the following:
* 1.)	Largest stock value increase and it's associated ticker.
* 2.)	Largest stock value decrease and it's associated ticker.
* 3.)	Largest stock volume and it's associated ticker.
* 4.)	Integrated into a for-loop to calculate deliverables across all worksheets in the workbook.

### Solution Example Screenshot:

![hard_solution](Images/hard_solution.png)

# Deliveribles:
* The script that covers the scope of the assignment and should be ran for the final analysis can be found in the Multiple_year_stock_data_COMPLETE.xlsm file of the Deliverables directory.
* Screen shots of each of the completed scripts for three seperate years can also be found in the Deliverables/Images directory of this repository 
