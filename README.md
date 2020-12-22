# VBA-challenge
This repo has been created to serve as a home for a series of VBA Challenges.

There are two resources files.
* Test File - for testing the VBA script
* Stock Data - the final dataset.

# Basic Task 

Create a script that will loop through all the stocks for one year and output the following information.
* The ticker symbol.
* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

There should also be conditional formatting that will highlight positive change in green and negative change in red.

# End Result

Honestly I didn't read the instructions properly to start with so I ended up shoving all the results to a new sheet. Then I realised that if additional data was added to the workbook in new yearly sheets, it wouldn't work properly so I had to exclude the first sheet. Then I realised that if we were adding new data we probably didn't want to process over the old data, so I needed to work out a way to skip that. Then I did that. Then I realised I didn't need to do that. 

The script should do the following:
* Determine if a summary sheet ("Final_Data") esists.
* If NO: Create Summary sheet with headers and "Greatest of" table.
* If YES: Scrape summary sheet for relevant data so that it can ignore sheets that have already been reviewed by the summary. Note: if the sheet has had changes made, but the sheet name has not changed, then the script will assume that no changes have been made.
* The script then goes through the workbook, compiles the following data points about each Stock and places it on the summary page:
  * Year (taken from worksheet name) (or letter, if using alphabetical data)
  * Total yearly trading volume
  * The Year opening price
  * The Year closing price
  * The year percentage change
  * The year amount change
* As it compiles the above data it works out which stock meets the following criteria and places it on the side of the summary page:
  * Greatest % Increase
  * Greatest % Decrease
  * Greatest trading volume
* Once everything is done, the script determines which stock on which year meets the above criteria of greatest changes and places it at the top of the greatest table on the summary tab. If we wanted to be *really* smart we would get the script to calculate the greatest changes over all data present. But I've made too much work for myself by being smart already, so lets not do that.  

If you copy the 2014 sheet (or if, as per the grading rubic you're using the alphabetical testing, the P tab) and rename it 2013 (or Q tab), and then re-run the script, it should *not* run through the entire workbook, and only create a summary for the data on the new sheet

At least it was good experience. 
