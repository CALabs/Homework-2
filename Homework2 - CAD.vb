Sub StockAnalyzer()

'Definition of needed variables
Dim Ticker As String
Dim Dates As Date
Dim xRange As String
Dim yRange As String
'Dim NewVal As Double
'Dim OldVal As Double
'Dim YoYChange As Double
'Dim PercentChange As Double
Dim VolSum As Double: VolSum = 0
'I tried to do the intermediate but I ran out of time :(

'Row counter of the data set
    RowCounter = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'MsgBox (RowCounter)
'Column counter of the data set
    ColCounter = Range("A1", Range("A1").End(xlToRight)).Columns.Count
    'MsgBox (ColCounter)
'Extract the unique tickers
    'y (Rows)Range for the last row value to string (Alphanumeric coordinate for the range)
    yRange = "A" & RowCounter
    'x (Columns)Range for the last column value to string (letter of a range given the count of columns)
    xRange = Split(Cells(1, ColCounter + 2).Address, "$")(1)
    'Command to select all the range (can be variable) of rows with data, specifically for the ticker, and creating the unique values in the column of the last data + 2
    Range("A1:" & yRange).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range(xRange & "1"), Unique:=True
    'Format
    Range(xRange & "1") = "Ticker"
    Range(xRange & "1").Font.Bold = True
    Range(xRange & "1").Font.Color = RGB(255, 255, 255)
    Range(xRange & "1").Interior.Color = RGB(0, 51, 153)
    
    'YoY Change
'     RowCounterOutput = Range(xRange & "1", Range(xRange & "1").End(xlDown)).Rows.Count
'
'    For i = 2 To RowCounterOutput
'        For o = 2 To RowCounter
'            If Cells(i, ColCounter + 2).Value = Cells(o, 1).Value And Cells(o, 2) = "20160101" Then
'            OldVal = Cells(o, 3)
'            ElseIf Cells(i, ColCounter + 2).Value = Cells(o, 1).Value And Cells(o, 2) = "20161230" Then
'            NewVal = Cells(o, 6)
'            YoYChange = NewVal - OldVal
'            PercentChange = NewVal / OldVal - 1
'            Cells(i, ColCounter + 3) = YoYChange
'            Cells(i, ColCounter + 4) = PercentChange
'            End If
'        Next o
'    Next i
'    'Format
'    xRangeb = Split(Cells(1, ColCounter + 3).Address, "$")(1)
'    xRangec = Split(Cells(1, ColCounter + 4).Address, "$")(1)
'    Range(xRangeb & "1") = "Yearly Change"
'    Range(xRangeb & "1").Font.Bold = True
'    Range(xRangeb & "1").Font.Color = RGB(255, 255, 255)
'    Range(xRangeb & "1").Interior.Color = RGB(0, 51, 153)
'    Range(xRangec & xRangec).NumberFormat = "0.0%_);(0.0%)"
'    Range(xRangec & "1") = "Percent Change"
'    Range(xRangec & "1").Font.Bold = True
'    Range(xRangec & "1").Font.Color = RGB(255, 255, 255)
'    Range(xRangec & "1").Interior.Color = RGB(0, 51, 153)

    'Total Stock Volue
    RowCounterOutput = Range(xRange & "1", Range(xRange & "1").End(xlDown)).Rows.Count
    
    For i = 2 To RowCounterOutput
        For o = 2 To RowCounter
            If Cells(i, ColCounter + 2).Value = Cells(o, 1).Value Then
            VolSum = VolSum + Cells(o, ColCounter)
            Cells(i, ColCounter + 3) = VolSum
            End If
        Next o
        VolSum = 0
    Next i
    'Format
    xRangeb = Split(Cells(1, ColCounter + 3).Address, "$")(1)
    Range(xRangeb & "1") = "Total Stock Volume"
    Range(xRangeb & "1").Font.Bold = True
    Range(xRangeb & "1").Font.Color = RGB(255, 255, 255)
    Range(xRangeb & "1").Interior.Color = RGB(0, 51, 153)
 
End Sub

