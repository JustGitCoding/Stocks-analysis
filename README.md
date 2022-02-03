# Stocks analysis

## Overview of Project
### Assessing stock performance over a two year period (2017 - 2018)
The goal of this project is to analyze stocks trade volume, as well as average rate of return for twelve popular ticker symbols. Additionally, using VBA script in Excel, we automated various aspects of our analysis, which can easily be re-used for data that may be obtained in for later years or for other ticker symbols.

## Results
### 2017 Data vs 2018 Data
![VBA_Challenge_2017](https://user-images.githubusercontent.com/97985062/152259632-ec78140e-8997-4246-a7e2-1ea1b6d9b88f.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/97985062/152259645-15094b9e-3a62-4ae9-a4e9-c32d8838c802.png)
### Comparisons
Based on the analysis, 11 out of 12 stocks had positive returns in 2017 while only 2 out of 12 had positive returns in 2018. These two stocks (ENPH and RUN) are likely the stronger investment choices, as we see that for both stocks, trade volume increased year over year, and 2018 data (most recent data available) had high return percentages (80%+).

```
For i = 2 To RowCount
	If Cells(i, 1).Value = tickers(tickerIndex) Then
    		tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            tickerIndex = tickerIndex + 1
        End If
        
Next i

```




Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.



## Summary
In a summary statement, address the following questions.
1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?
