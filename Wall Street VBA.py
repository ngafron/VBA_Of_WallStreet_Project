Dim ws As Worksheet
Dim ticker As String
Dim Yearopen As Double
Dim Yearclose As Double
Dim Yearchange As Double
Dim Volume As Double
Dim lastrow As Double
Dim firstrow As Double
Dim tickerrow As Integer
Dim Increase As Integer
Dim Decrease As Integer
Dim BigVolume As Double

On Error Resume Next

'loop through each worksheet
For Each ws In Worksheets

'set last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
firstrow = ws.Cells(Rows.Count, 1).Start(xlUp).Row

tickerrow = 2
Increase = 0
Decrease = 0

'loop through every row
    For i = 2 To lastrow
    
    If Yearopen = 0 Then
    Yearopen = ws.Cells(i, 3).Value
    
    End If
    
    'insert end values of tickers
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Volume = Volume + ws.Cells(i, 7).Value
    Yearclose = ws.Cells(i, 6).Value
    ws.Cells(tickerrow, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(tickerrow, 10).Value = Yearclose - Yearopen
    Percent = FormatPercent((Yearclose / Yearopen) - 1)
    ws.Cells(tickerrow, 11).Value = Percent
    ws.Cells(tickerrow, 12).Value = Volume

    'set/prepare new values
    tickerrow = tickerrow + 1
    Yearopen = 0
    Yearclose = 0
    Volume = 0
    
    'add values of tickers together
    Else
    Volume = Volume + ws.Cells(i, 7).Value
    
    End If
    
    If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4

    Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    
    
    Next i

Next ws

End Sub
}
