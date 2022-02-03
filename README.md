# Stocks analysis

## Overview of Project
### Assessing stock performance over a two year period (2017 - 2018)
The goal of this project is to analyze stocks trade volume, as well as average rate of return for twelve popular ticker symbols. Additionally, using VBA script in Excel, we automated various aspects of our analysis, which can easily be re-used for data that may be obtained in for later years or for other ticker symbols.

### Refactoring VBA script for efficiency
Another goal in this project is to find a way to make our VBA script more efficient (take less time to run). We achieved this by changing the code to loop through the data once (as opposed to once per stock) and saving the output in arrays.

## Results
### 2017 Data vs 2018 Data
![VBA_Challenge_2017](https://user-images.githubusercontent.com/97985062/152259632-ec78140e-8997-4246-a7e2-1ea1b6d9b88f.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/97985062/152259645-15094b9e-3a62-4ae9-a4e9-c32d8838c802.png)
### Stock Performance Comparison
Based on the analysis, 11 out of 12 stocks had positive returns in 2017 while only 2 out of 12 had positive returns in 2018. These two stocks (ENPH and RUN) are likely the stronger investment choices as we see that for both stocks, trade volume increased year over year, and 2018 data (most recent data available) showed high rates of return.

### Impact of Refactoring
After performing the analysis above, we refactored the VBA script such that the code looped through the 3,000+ lines of data only once (as opposed to once per stock ticker), and saved each stock's trade volume and start/end prices into output arrays, which could be easily referenced. 

```
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
    Next j
                                
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            'capture ticker Starting Price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            'capture ticker ending price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
        
    Next i

```
This resulted in drastic time savings with the old code running in ~0.37 seconds on average, and the refactored code running in ~0.05 seconds (over 7x faster!)

## Summary
### Refactoring Code in General
The main advantage of refactoring code is to improve the efficiency of the code. This is generally achieved by simplifying the code so that it takes fewer steps or utilizes less memory (all without sacrificing functionality). However, some disadvantages of refactoring code include cost (time & money) as well as risk of breaking the code by oversimplifying it such that it no longer considers all potential use cases (i.e. the code may no longer work in fringe/edge cases).
### Impact of Refactoring our VBA script
As noted above, the primary benefit of refactoring is that our VBA script now runs over 7x faster than the original script. Because our VBA script is rather simple to begin with, the cost of refactoring is low, and the risk of oversimplification is remote.