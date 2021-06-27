# Stock Analysis With Excel VBA
## Overview of the Project
Steve wants to help his parents to buy some profitable shares by analysing and running some automated VBA macros.In this challenge,we are writing VBA code to automate some tasks to do DQ Analysis and all stocks analysis too.Later ,will edit and refactor the  solution code to loop through all the data one time in order to collect the same informations. Then, we will determine whether refactoring our code successfully made the VBA script run faster.

### Purpose
As part of an assignment for the UT Data Boot Camp, an initial stock analysis was conducted for Steve.  The original code was then refactored to loop only once.  The purpose is to determine if the refactored changes made an impact on the run time.  

Using run buttons, Steve will have the ability to put in the year into an input box, which removes any magic numbers.The VBA code uses a timer, arrays, if-then conditional statements, assigns different data types, and adds static/conditional formatting.


### Analysis

Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring by,
*Creating a tickerIndex variable and set it to zero*
*The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.*
*The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.*
*formatted the cells in the spreadsheet as required.*
*Checked if the outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.*

## Results

-DQ Analysis
While analysing the yearly return of DQ shares for the year 2018,we could see that DQ has dropped over 63%.

![DQ-Analysis.png](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/DQ-Analysis.PNG?raw=true)

-Total Stocks Analysis
By analysing the yearly returns of all the stocks for the years 2017 and 2018 ,we could see that 
*For the year 2017,Except TERP share ,all other shares have grown.
*For the year 2018,except ENPH and RUN shares have shrunk.
![Allstocks2017.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/Allstocks2017.PNG?raw=true)
![allstocks2018.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/allstocks2018.PNG?raw=true)

-While running the All Stocks Analysis file before refactoring,I have got the following results.
![Runtime-before-refactoring2017.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/Runtime-before-refactoring2017.PNG?raw=true)
![Runtime-before-refactoring2018.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/Runtime-before-refactoring2018.PNG?raw=true)

-The original code was now refactored to loop only once.I have defined a ticker variable and initialised to zero.Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerclosingPrices.I have got the following results.
![VB_Challenge_2017.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/VB_Challenge_2017.PNG?raw=true)
![VB_Challenge_2018.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/VB_Challenge_2018.PNG?raw=true)

-Comparing the Original Run Times to the Refactored Run Times

**Run time for the original code took around 0.51 seconds.
Run time for the refactored code took around 0.09 seconds,which means Refactoring the code did make the run times decrease by optimizing the code.** 

-The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.
-Formatting of data is done successfully.
![Formated-datasheet.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/Formated-datasheet.PNG?raw=true)



-Here is the change of code while refactoring is done.

![Tickerindex&arrays.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/Tickerindex&arrays.PNG?raw=true)


![ticker-index-used.PNG](https://github.com/Praveeja-Sasidharan-Suni/stock-analysis/blob/main/Resources/ticker-index-used.PNG?raw=true)

# Summary
## 1. What are the advantages or disadvantages of refactoring code?
### Advantages
Finding the root cause of a potential bug can be done with refactoring.  A programmer can catch duplicated subroutines, unnecessary loops, redundant statements, or code that was used to run down an error but was accidently left in the script. 

### Disadvantages

programmers may have alternative logic steps, which will require testing to see how those differences play out in the script.  Refactoring a stable code to apply a different set of logic could be costly or introduce new bugs into the system.Sometimes it may be complicated and time consuming too.  
	
## 2. How do these pros and cons apply to refactoring the original VBA script?
Reducing the number of loops decreases the memory needed for processing the data, which reduces the run time and optimizes the performance of the script. To refactor the code, testing has to be done with each new addition to check for the efficiency of the new code.  

   
