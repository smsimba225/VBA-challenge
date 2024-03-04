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
I created a loop that went down the Ticker Row A and checked to see if the next row had the same Ticker Symbol.
If the next Ticker Symbol on the next row was the same, then it would do nothing except add the Stock Volume to a TotalStock variable.
Finally when the loop reaches a line where the following line is a different Ticker Symbol, then the code would calculate the Yearly Change, Percentage Change and the final Total Stock Volume.
The code then reported the information to the report section in the middle of the sheet.  It would also check to see if it is the Greatest Percent Increase or Decrease, as well as the Greatest Total Volume so far in the loop, store them into new variables, then reset      the TotalStock variable to be used in the loop for the next Ticker Symbol.

While Debugging I realized I also needed to insert code that would reset the OpenPrice, ClosingPrice, PercentChange as well, otherwise it would mess up the results while working in other worksheets. 

I did put messages along my code to remind myself what each line I did was for.

---------
References:  The Microsoft  Visual Basic for Applications article "Macro to Loop Through All Worksheets in a Workbook" was very helpful in figuring out how to loop the code across all of the worksheets in the workbook.  I used the same variable name 'Current' from one of the examples on the page and used the For Each loop to accomplish what I needed. https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0


---------
VBA was a lot of fun to learn, and it was very rewarding to have the ability to process a huge spreadsheet in just a matter of seconds.  Thank you for looking at my work!

~Owen
  
  
