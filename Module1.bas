Attribute VB_Name = "Module1"
Sub MacroCheck()
    Dim testMessage As String
    testMessage = "Hello World!"
    MsgBox (testMessage)
End Sub

Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "Foo DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    Dim totalVolume As Double
    Dim startingPrice As Double
    Dim endingprice As Double
    rowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).row
    totalVolume = 0
    For row = rowStart To rowEnd
        If Cells(row, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(row, 8).Value
        End If
        If Cells(row, 1).Value = "DQ" And Cells(row - 1, 1).Value <> "DQ" Then
            startingPrice = Cells(row, 6).Value
        End If
        If Cells(row, 1).Value = "DQ" And Cells(row + 1, 1).Value <> "DQ" Then
            endingprice = Cells(row, 6).Value
        End If
    Next row
    
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingprice / startingPrice - 1
    
End Sub

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer

    '1) Format the output sheet on the "All Stocks Analysis" worksheet.
    'Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Worksheets("All Stocks Analysis").Cells.Clear
    Range("A1").Value = "All Stocks (2018)"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2) Initialize an array of all tickers.
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
    
    '3) Prepare for the analysis of tickers.
    '3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingprice As Single
    
    '3b) Activate the data worksheet.
    'Worksheets("2018").Activate
    Worksheets(yearValue).Activate
    
    '3c) Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    '4) Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5) Loop through rows in the data.
        Worksheets("2018").Activate
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
                endingprice = Cells(j, 6).Value
            End If
        Next j
        
        '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingprice / startingPrice - 1
    
    Next i
    
    endTime = Timer
    MsgBox "This code ran n " & (endTime - startTime) & " seconds for the year " & (yearValue)
    'nrows = 10
    'ncols = 10
    'For row = 0 To nrows - 1
    '    ticker = tickers(row)
    '    For col = 0 To ncols - 1
    '        'Cells(row+1, col+1).Value = ticker
    '        Cells(row + 1, col + 1).Value = row + col
    '    Next col
    'Next row
   
End Sub

Sub formatAllStocksAnalysisTable()
    Worksheets("All Stocks Analysis").Activate
    'Worksheets("All Stocks Analysis").Cells.Clear
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$ #,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
           'Color the cell green
           Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
           'color the cell red
           Cells(i, 3).Interior.Color = vbRed
        Else
           'clear the cell color
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
End Sub

Sub foo()
    Dim x As integ
    x = 42
    MsgBox ("hello" + x)
End Sub

Sub ClearWorksheet()
    Cells.Clear
End Sub

Sub yearValueAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Worksheets(yearValue).Activate
    
    Worksheets("All Stocks Analysis").Activate
    Worksheets("All Stocks Analysis").Cells.Clear
   
End Sub

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
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
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
    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next i
    
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
