# VBA_Assignment_Stock_Data

<h3>Assignment for VBA Coding</h3>

<p>The purpose of this assignment is to produce a VBA script to analyse three years of generated stock market data. It consists of a summary table outputting the following information:</p><ul>
  <li>The ticker symbol</li>
  <li>Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year for each ticker.</li>
  <li>The results are printed using conditional formatting highlighting positive change in green and negative change in red.</li>
  <li>The total stock volume of the stock traded over the course of that year.</li>
  <li>The same analysis is performed for 2018, 2019 and 2020 simultaneously, each contained in different worksheets.</li>
  <li>A "Bonus" table is included listing the tickers with the greatest percentage increase, the greatest percentage decrease and the greatest total stock volume over each year.</li></ul>

<h3>Method</h3><ul>
  <li>The script loops thorugh each year of stock data and providing summaries on each worksheet of that year's data by ticker</li>
  <li>The loop reads and stores the following values:</li><ol>
      <li>ticker symbol;</li>
      <li>opening price;</li>
      <li>closing price;</li>
      <li>volume of stock.</li></ol>
  <li>On the same worksheet as the raw data, or on a new worksheet all columns are created for:</li></ol>
      <li>ticker symbol;</li>
      <li>total stock volume;</li>
      <li>yearly change ($);</li>
      <li>percent change.</li></ol>
 <li>Conditional formatting is appliedto:</li><ol>
      <li>the yearly change column;</li>
      <li>the percent change column.</li></ol>
 <li>Witin the bonus requirements, the following values are provided on each sheet:</li><ol>
       <li>Greatest % Increase;</li>
       <li>Greatest % Decrease;</li>
       <li>Greatest Total Volume.</li></ol></ul>

  <h2>Code</h2><ul>
    <li>Each variable's data-type is defined at the top of each sub-routine for ease of reading.</li>
    <li>Option Explicit is set at the start of the script to ensure all variables are defined.</li>
    <li>The last row of each worksheet containing data is defined using <i>ws.Cells(Cells.Rows.Count, 1).End(xlUp).Row</i> and later as the finishing point for For loop</li>
    <li>Summary table headings are set up and formatted.</li>
    <li>For loop defined, testing whether the ticker value contained in the row currently looked at is the same value as the cell above. If it is the loop will jump to the Else statement and moving on to the next row, whilst  adding the current stock volume to the stock volume counter.</li>
    <li>If the value of the Cell is different, then the loop continues through the rest of the steps, storing ticker value, the opening value and closing value, working out the percentage annual change and adding the final stock volume value to the counter. This is output into the respective columns in the output table.</li>
    <li>To retrieve the opeining value and closing value the Find function is used. Both functions employ the SearchDirection (<i>xlNext</i> and <i>xlPrevious</i>) and LookAt parameters (<i>LookAt:=xlWhole</i>) ensuring the first and last ticker rows found match the exact value of the ticker (ensuring ticker names like FAAT are not included in the results for AAT). </li>
    <li>The conditional formatting is contained in an If / else statement contained wihtin the loop.</li>
    <li>The bonus section employs a seperate For loop, looping thorugh the summary table using the Match, Max and Min functions to retrieve the correct data to populate the secondary summary</li></ul>

<h2>Observation</h2>

<p>This code takes greater than twenty minutes to run because of the amount of data within each spreadsheet (approx 750,000 lines each) and because the code is not written optimally. Refactoring the code is not part of this exercise, but there are efficiencies that could be employed. The data can be stored in an array, which is exponentially faster than using ranges and treating each column seperately. Additionally, functions such as Find and Match possibly have more efficient alternatives.</p>
<p>The only step taken that speeds up the code is to disable ScreenUdating, which does improve the speed, but not significantly.</p>

<h2>Sources</h2>
https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

