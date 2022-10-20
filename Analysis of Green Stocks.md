# An Analysis of Green Energy Stock Data

## Overview and Purpose of Project
The analyst was tasked with creating a VBA script that analyzed green energy stock data for Steve, a recent college graduate.  Steve has taken on his parents as clients and wants to get a better understanding of a list of stocks so that he can guide his parents towards smart allocations of their investments.  Using VBA, the analyst was able to create a script that gathered the total daily volume and rate of return for a list of 12 stocks over the period of a year, and then refactored that code to make the program run more efficiently. 

## Analysis and Results

### User Interface & Analysis Calculations
As noted previously, the VBA script that was created gathered the total daily volume and rate of return for a list of 12 stocks over the period of a year, specifically analyzing 2017 and 2018.  The code was designed in such a way that the user could simply click a button:

![Run_Analysis_Button.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/Run_Analysis_Button.png)

Input the desired year to be calculated:

![Input_box.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/Input_box.png)

And an analysis would be conducted.

To turn the raw data into valuable information, the analyst created a series of For loops and if statements to calculate the total daily volume, starting, and ending prices of the array of stocks.  Those prices were then used to calculate the rate of return for the array.  The code for those calculations are shown below:

    'For i = 2 To RowCount      

        If Cells(i, 1).Value = tickers(tickerIndex) Then
            
        '3a) Increase volume for current ticker:
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
            
           
        '3b) Check if the current row is the first row with the selected tickerIndex:
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker:
        
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        '3d Increase the tickerIndex:
        tickerIndex = tickerIndex + 1
        
        End If
        
        
    Next i'

Using those values, the analyst was able to output the desired information for their analysis.  The code for the outputs of the stock names, total daily volume, and the calculation of the rate of return are shown below:

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i'

### Stock Analysis Comparison
The analysis of the 2017 and 2018 stocks showed that the while 2017 was lucrative in the rate of returns for many of the stocks (many positive returns), 2018 saw the majority of those rates of return dive into the negatives.  The comparisons of the 2017 and 2018 results are shown below:

2017:

![2017_All_Stocks.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/2017_All_Stocks.png)

2018:

![2018_All_Stocks.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/2018_All_Stocks.png)

The comparison shows that, other than ENPH and RUN, the return on investment drastically decreased from 2017 to 2018.  It is also worth noting that while TERP had a negative return each year, the stock gained 2.2% in return over the course of the year (-7.2% in 2017, -5.0% in 2018).

### Refactoring the VBA Script
The analyst was also tasked with refactoring the VBA script and timing the difference between the original code vs. the refactored code.  Refactoring code is a process in which a programmer goes through the script and modifies/maintains the code so that it stays up to date and runs efficiently.  Some advantages of refactoring are that is improves the design of the code, makes it easier to understand, helps find bugs, and make it run faster.  Disadvantages to the refactoring process are that it is time consuming, which can potentially cost the client or programmer money.  

In regards to script that the analyst created, the refactoring process made the code run far faster than the original.  The run times of the original and refactored code are shown below:

Original 2017 Run Time:

![2017_AllStocksAnalysis_Runtime.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/2017_AllStocksAnalysis_Runtime.png)

Original 2018 Run Time:

![2018_AllStocksAnalysis_Runtime.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/2018_AllStocksAnalysis_Runtime.png)

Refactored 2017 Run Time:

![2017_AllStocksAnalysis_Refactored_Runtime.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/2017_AllStocksAnalysis_Refactored_Runtime.png)

Refactored 2018 Run Time:

![2018_AllStocksAnalysis_Refactored_Runtime.png](https://github.com/hillmanj1995/stock-analysis/blob/main/Resources/2018_AllStocksAnalysis_Refactored_Runtime.png)

The comparison of the original vs. refactored runtimes show that after refactoring, the code ran anywhere from 5x to 6x faster.  The disadvantage of the refactoring was that it took a considerable amount of time for the analyst to modify and debug the code. 
