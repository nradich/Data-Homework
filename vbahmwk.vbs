Sub Button1_Click()


'select ticker and add volumes together


Dim I As Double
Dim ticker  As String
Dim next_ticker As String
Dim WS_Count As Integer
Dim ws As Worksheet



Dim ticker_vol As Double
Dim position As Integer
Dim LastRow As Double
Dim lastColumn As Double

For Each ws In Worksheets


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column


position = 2


'loop through all worksheets

'loop through stocks


For I = 2 To LastRow
ws.Cells(1, lastColumn + 1).Value = "Ticker ID"
ws.Cells(1, lastColumn + 2).Value = "Total Stock Volume"
ticker = ws.Cells(I, 1).Value
next_ticker = ws.Cells(I + 1, 1).Value
ticker_vol = ticker_vol + ws.Cells(I, 7).Value
prev_ticker = ws.Cells(I-1, 1).Value
close_price = ws.Cells( I, 6).Value
open_price = ws.Cells( I, 3).Value
yearly_change = close_price - open_price
    
    If ticker <> next_ticker Then
        ws.Cells(position, lastColumn + 1).Value = ticker
        ws.Cells(position, lastColumn + 2).Value = ticker_vol
        position = position + 1
        ticker_vol = ws.Cells(I + 1, 7).Value
    End If

Next I

Next ws

End Sub
