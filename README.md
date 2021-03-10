# Stock-Analysis
Exploring Green Energy Stock Performance by analyzing Financial Data using VBA

## Overview of Project
We helped steve analyze a handful of green energy stocks in addition to DAQO stocks where his parents were planning to invest their money. Steve loved the workbook we prepared for him. At the click of a button, he can analyze an entire dataset. 

Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this project, we will edit, or refactor, the Module 2 solution code to loop through all the data one time to collect the same information that we did in this module. Then, we will determine whether refactoring code successfully made the VBA script run faster.

### Elapsed Time for 2017 and 2018 - Module 2 

![image](https://user-images.githubusercontent.com/78935551/110395578-61ac3180-803c-11eb-80a1-b554a74b477c.png)

![image](https://user-images.githubusercontent.com/78935551/110395592-67a21280-803c-11eb-80e0-f814f48cce99.png)


## Refactor VBA code and measure performance
Using our knowledge of VBA and the starter code provided for the Challenge to refactor the Module 2 script we looped through the data and collected all information.
Our refactored code below now runs faster than it did in this module. 

    '1a) Create a ticker Index
     tickerIndex = 0

    '1b) Create three output arrays
    Dim TickerVolumes(12) As Long
    Dim TickerStartingPrices(12) As Single, TickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    TickerVolumes(i) = 0
       
    Next i
            
    ''2b) Loop over all the rows in the spreadsheet.
      For i = 2 To RowCount
    
     '3a) Increase volume for current ticker
         TickerVolumes(tickerIndex) = TickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
     '3b) Check if the current row is the first row with the selected tickerIndex.
       
     'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        TickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           
      End If
        
    '3c) check if the current row is the last row with the selected ticker
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         TickerEndingPrices(tickerIndex) = Cells(i, 6).Value
       
       End If

    '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        
    End If
        
    Next i
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = TickerVolumes(i)
    Cells(4 + i, 3).Value = TickerEndingPrices(i) / TickerStartingPrices(i) - 1
            
    Next i
        
## Summary 
### Results
- With these Macros created we can clearly see the table for each year.
- Steve can read the table lot easier due to conditional formatting.
- Updated Macro's can be used to run analysis for any year.
- We refactored the codes to run faster in VBA so that if Steve has a larger dataset he can analyze it quickly.

### Advanatages of Refactoring codes.
- Refactoring is a key part of coding process. It just makes the code more efficient with fewer steps.
- Refactoring codes helps reduce the run time of the Macro's.
- It looks much cleaner and helps future user to undersatnd and read it better.

### Disadvanatages of Refactoring codes.
- Refactoring codes can be time consuming
- It might be difficult for larger and more complicated codes to refactor
- 






- 
- 


