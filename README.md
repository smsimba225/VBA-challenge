Module 2 Challenge

These were the instructions : 
Create a script that loops through all the stocks for one year and outputs the following information:
-------
The ticker symbol
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
--------
Starting this challenge was the most difficult part.
Created a loop that goes down the Ticker Row A and checks to see if the next row had the same Ticker Symbol.
If the next Ticker Symbol on the next row was the same, then it would do nothing except add the Stock Volume to a TotalStock variable.
Finally when the loop reaches a line where the following line is a different Ticker Symbol, then the code would calculate the Yearly Change, Percentage Change and the final Total Stock Volume
The code then reported the information to the report section in the middle of the sheet.  It would also check to see if it is the Greatest Percent Increase or Decrease, as well as the Greatest Total Volume so far in the loop, store them into new variables, then reset      the TotalStock variable to be used in the loop for the next Ticker Symbol.

While Debugging I realized I also needed to put code that would reset the OpenPrice, ClosingPrice, PercentChange as well, otherwise it would mess up the results while working in other worksheets. Some of the challenges included putting the resets  back to "0" in the correct places among the If, Elseif, and Endif statements.

---------

VBA was a lot of fun to learn, and it was very rewarding to have the ability to process a huge spreadsheet in just a matter of seconds.  Thank you for looking at my work!

~Owen
  
  
