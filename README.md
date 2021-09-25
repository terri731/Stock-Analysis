# Stock-Analysis
## Overview of Project
### Purpose
The purpose of this project was to refactor VBA Code in Microsoft Excel. 
The file used to refactor contains stock data from 2017 and 2018. It will allow someone to determine which stocks are worth investing in.
### Results
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'If the next row's ticker doesn't match, increas the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row;s ticker doesn;t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cellls(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                
                
                
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
### Summary
The advantages of refactoring code is that it gives you the ability to make it organized and easier to read. 
Theis also allows for better design and debugging. It also offers the benefit of allowing others to read the code easier. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/89110920/134783573-d44067a0-ba12-4a25-93b0-cddd91fac12f.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/89110920/134783577-474e91fa-7085-41e8-bf07-1fbf4449b157.png)
