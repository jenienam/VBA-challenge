Sub stocks():
    Dim ticker As String
    Dim ticker_count As Integer
    Dim lastRow As Long
    Dim lastRowTicker As Long
    Dim tickerVolume As Double
    Dim tickerSummaryRow As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim year_change As Doublea
    Dim percent_change As Double

    For Each ws In Worksheets
    tickerSummaryRow = 2
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastRowTicker = 1
    tickerVolume = 0
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To lastRow
        ClosePrice = ws.Cells(i, 6).Value
        OpenPrice = ws.Cells(i - lastRowTicker + 1, 3).Value
        ticker_count = 0
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            lastRowTicker = lastRowTicker + 1
        End If
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            year_change = ClosePrice - OpenPrice
            Cells(tickerSummaryRow, 10).Value = yearly_change
            tickerVolume = tickerVolume + ws.Cells(i, 7).Value
            If year_change <> 0 Then
                percent_change = (year_change / OpenPrice)
            End If
            ws.Cells(tickerSummaryRow, 9).Value = ticker
            ws.Cells(tickerSummaryRow, 10).Value = year_change
            ws.Cells(tickerSummaryRow, 11).Value = percent_change
            ws.Cells(tickerSummaryRow, 11).NumberFormat = "0.00%"
            OpenPrice = ws.Cells(i + 1, 3).Value
            SummaryRow = SummaryRow + 1
            stockTotal = 0
            
        If year_change > 0 Then
                ws.Cells(ticker_count + 1, 10).Interior.ColorIndex = 4
        Else
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 6
       End If
    End If
    
Next i
Next ws
End Sub






