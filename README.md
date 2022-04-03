# Stock-Analysis

## Overview of Project:

The goal of this project was to refactor code I put together in VBA to analyze stock information from 2017 and 2018, faster. I did this with the idea in mind of re-using this code in the future on larger data sets. The analysis runs through the data 1 time instead of running through ticker by ticker and will provide data regarding the Total Daily Volume for each stock ticker and the Return, depending on which year you enter when initiating the analysis. 

## Results:
    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Initialize array of all tickers
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    '1a) Create a ticker Index
    tickerIndex = 0
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        '3c) check if the current row is the last row with the selected ticker
        'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        '3d Increase the tickerIndex.
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
           Cells(i, 3).Interior.Color = vbGreen
        Else
           Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    End Sub


### Stock Performance: 

In 2017 every stock ticker had a positive return outside of TERP, compared to 2018 where every stock outside of ENPH and RUN had a negative return. 

### Execution Times: 

## Summary: 

### What are the advantages or disadvantages of refactoring code?


### How do these pros and cons apply to refactoring the original VBA script?
