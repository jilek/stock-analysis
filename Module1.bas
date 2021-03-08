Attribute VB_Name = "Module1"
Sub MacroCheck()
    Dim testMessage As String
    testMessage = "Hello World!"
    MsgBox (testMessage)
End Sub

Sub DQAnalysis()
    Worksheets("2018").Activate
    
    'Range("A1").Value = "DAQO (Ticker: DQ)"
    'Cells(3, 1).Value = "Year"
    'Cells(3, 2).Value = "Total Daily Volume"
    'Cells(3, 3).Value = "Return"
    
    'For i = 1 To 8
    '    MsgBox (Cells(1, i))
    'Next i
    
    Dim rowStart As Long
    Dim fowEnd As Long
    Dim totalVolume As Double
    Dim row As Integer
    Dim col As Integer
    rowStart = 2
    rowEnd = 3013
    totalVolume = 0
    col = 8
    For row = rowStart To rowEnd
        If Cells(row, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(row, col).Value
        End If
    Next row
    MsgBox (totalVolume)
    
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    
'3306038200
    
End Sub
