# VBA_Assignment_Stock_Data

Assignment for VBA Coding 

This macro produces a summary table for each ticker, recording the difference between the opening price at the beginning opening price and closing price over a year.
It also records the yearly percentage value change for each stock and the total stock volume traded.
The yearly value change per ticker is highlighted with Green indicating an increase in total value over the year and Red indicating a decrease.
the Macro is designed to loop through each tab, iwth the data on each tab representing a single year. The dataset has multiple years contained wihtin it.

A second table has been produced (the bonus table), which identifies the Tickers that have recorded the Greatest Percentage Increae, the Greatest Percentage Decrease and the Greatest Total Volume traded for each year. the bonus table is recorded to the right of the summary table.

Notes:
The variables types have been set to Double as Integer and Long proved too small to hold the requisite numbers for the entire spreadsheet. They caused overflow errors.

Loops are used for production of the summary table, and to go through all the worksheets in the workbook. The bonus table only loops thorugh each worksheet, sunbing functions such as Match, Max and Min to retrieve the correct data to populate the summary tabe.

The macro will sort through data sets with varying numbers of rows and should be able to deal with any empty rows.

Research has been done to find efficient ways of processing the data, and the code reflects the reading undertaken. Particular issues with debugging, syntax and error messages were researched on a number of sites, such as StackOverflow. The only code actually used without adaptation is the worksheet loop. The source is also cited in the comments within the code. This was provided in the Read_Me file provided with the US Census Pt1 classroom challenge: 

https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

The code is extensively commented. Addiitonal formatting has been provided ensureing that titales are in bold and that the columns are sized correctly to the information displayed. the ScreenUpdate function is also turned off while the code is running - to save time and annoying screen flicker.

The creation of the bonus table is contained within a seperate sub-routine to the summary table. this is to make the code more manageable to anyone reading it. The Sub-routine for the bonus table is pulled in to the overall creation by including it as a seperate procedure at the end of the main routine.

Option Explicit has been set to ensure all variables are defined and do not default to "Variant" type.
