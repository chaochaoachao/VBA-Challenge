# VBA-Challenge

Create a script that will loop through all the stocks for one year for each run and take the following information.

Raw data have <ticker> <date> <open price> <high> <low> <close> <volume> columns

##The ticker symbol.
##Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
##The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
##The total stock volume of the stock.


Including conditional formatting that will highlight positive change in green and negative change in red.
The result  look as follows.
![](Image/moderate_solution.png)


it will also be able to return the stock with the 
"Greatest % increase",
"Greatest % Decrease" 
and "Greatest total volume". T
The solution will look as follows:


The VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

Other Considerations

alphabetical_testing.xlsx was used while developing my code. This data set is smaller and will allow to test faster. 
This code should run on this file in less than 1 minutes in my laptop(looping through 28,0000 rows of raw data)
The script acts the same on each sheet. 
The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.
