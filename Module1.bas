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

    '1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
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
    Worksheets("2018").Activate
    
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

Sub ClearWorksheel()
    Cells.Clear
End Sub

Sub yearValueAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Worksheets(yearValue).Activate
    
    Worksheets("All Stocks Analysis").Activate
    Worksheets("All Stocks Analysis").Cells.Clear
   
End Sub
