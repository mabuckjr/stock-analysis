# 2017-2018 Stock Analysis Project Using VBA
## Overview of Project
### Purpose
The purpose of the project was to refactor some VBA code that I created. The original macro ran well, but wasn't running as efficiently or fast as it could. It was designed to show which companies would be the best to invest in given stock data from 2017 and 2018. In order to edit the macro for larger datasets in the future, I created an index with multiple arrays so that the macro only looped through the data one time. Both macros accomplish the same task and utilize many of the same tools (i.e., indexes, for loops, if then statements, etc.), but they end up looking very different by the end.
### The Given Data
The original data is distributed into 8 columns: Ticker, Date, Opening Price, High Price, Low Price, Closing Price, Adjusted Closing Price, and Volume of Trades. There are 2 different sheets representing the values for both 2017 and 2018. The twelve companies (represented by their tickers) are AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR. For every day with active trades, the appropriate data is displayed in the columns. The ultimate goal was to extract the tickers, the total daily volume, and return for each stock from either the 2017 or 2018 sheet using one subroutine.
## Results

```
 '1a) Create a ticker Index
    tickerindex = 0
        
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
            
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
                    
        
        '3b) Check if the current row is the first row with the selected tickerIndex.

        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
        End If
            

        '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerindex = tickerindex + 1
        End If
            
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    ```
## Summary
