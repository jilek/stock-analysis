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
    Worksheets("All Stocks Analysis").Activate
    Worksheets("All Stocks Analysis").Cells.Clear
    
    'Range("A1").Value = "All Stocks (2018)"
    'Cells(3, 1).Value = "Tcker"
    'Cells(3, 2).Value = "Total Daily Volume"
    'Cells(3, 3).Value = "Return"
    
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
    
    nrows = 10
    ncols = 10
    For row = 0 To nrows - 1
        ticker = tickers(row)
        For col = 0 To ncols - 1
            'Cells(row+1, col+1).Value = ticker
            Cells(row + 1, col + 1).Value = row + col
        Next col
    Next row
   
End Sub
