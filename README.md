# Stock Analysis

## Overview of Project
A detailed analysis of several Green Energy Stocks to aid Steve's parents with their investment decisions.

### Purpose
The purpose of this project was to create a quick, easy-to-use workbook using VBA to filter each green stock's data down to the Total Daily Volume and Yearly Return. This would aid Steve in forming a substantiated opinion and in presenting his findings to his parents. Hopefully, convincing them to look beyond the DQ symbol and it's sentimental values when choosing stocks to invest in.
(https://github.com/Jrheldmann/stock-analysis/blob/main/VBA_Challenge.xlsm)
## Challenges and Difficulties Encountered

### The original body of work, however, was not as efficient as it could've been. Thus, the code required refactoring.
---
    Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    '1) Format the output sheet on the "All Stocks Analysis" worksheet.
    
        startTime = Timer
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
   
    '2) Initialize an array of all tickers.
    Dim tickers(11) As String
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
    
    '3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    '3b) Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    '3c) Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through the tickers.
    For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0
        
        '5)Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            
            '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
        
        '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders.LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3).Value > 0 Then
            'Set cell to green if positive
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3).Value < 0 Then
            'Set cell to red if negative
            Cells(i, 3).Interior.Color = vbRed
        Else
            'Clear cell color
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    End Sub
---
## Results

- What are two conclusions you can draw about the Stocks?

- What are some limitations of this dataset?
