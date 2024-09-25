Sub StockMarketAnalysis_D()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim LastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Volume As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim TickerIncrease As String
    Dim TickerDecrease As String
    Dim TickerVolume As String
    Dim i As Long
    Dim StartRow As Long
    Dim OutputRow As Long

    ' Set the worksheet manually to the desired sheet
    Set ws = ThisWorkbook.Sheets("D")

    ' Initialize values
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    OutputRow = 2 ' Set output to start at row 2

    ' Add column headers on each sheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' Find the last row of the worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through each ticker in the current sheet
    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i, 3).Value
            Volume = 0
            StartRow = i
        End If

        ' Add volume for the current row
        Volume = Volume + ws.Cells(i, 7).Value

        ' Check if next row is a different ticker or the end of the data
        If ws.Cells(i + 1, 1).Value <> Ticker Or i = LastRow Then
            ClosePrice = ws.Cells(i, 6).Value
            QuarterlyChange = ClosePrice - OpenPrice

            ' Avoid division by zero
            If OpenPrice <> 0 Then
                PercentChange = (QuarterlyChange / OpenPrice) * 100
            Else
                PercentChange = 0
            End If

            ' Output data for the first row of the current ticker
            ws.Cells(OutputRow, 9).Value = Ticker
            ws.Cells(OutputRow, 10).Value = QuarterlyChange
            ws.Cells(OutputRow, 11).Value = PercentChange
            ws.Cells(OutputRow, 12).Value = Volume

            ' Apply conditional formatting ONLY to the "Quarterly Change" column (column 10)
            If QuarterlyChange > 0 Then
                ws.Cells(OutputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive
            ElseIf QuarterlyChange < 0 Then
                ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative
            Else
                ws.Cells(OutputRow, 10).Interior.ColorIndex = xlNone ' No color for zero
            End If
            
            ' Apply conditional formatting to "Percent Change" (column 11)
            If PercentChange > 0 Then
                ws.Cells(OutputRow, 11).Interior.Color = RGB(0, 255, 0) ' Green for positive
            ElseIf PercentChange < 0 Then
                ws.Cells(OutputRow, 11).Interior.Color = RGB(255, 0, 0) ' Red for negative
            Else
                ws.Cells(OutputRow, 11).Interior.ColorIndex = xlNone ' No color for zero
            End If

            ' Check for greatest increase, decrease, and volume
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                TickerIncrease = Ticker
            End If

            If PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                TickerDecrease = Ticker
            End If

            If Volume > GreatestVolume Then
                GreatestVolume = Volume
                TickerVolume = Ticker
            End If

            OutputRow = OutputRow + 1
        End If
    Next i

    ' Output greatest increase, decrease, and volume to the right of the data
    ws.Cells(2, 14).Value = "Greatest % Increase: "
    ws.Cells(2, 15).Value = TickerIncrease
    ws.Cells(2, 16).Value = GreatestIncrease & "%"

    ws.Cells(3, 14).Value = "Greatest % Decrease: "
    ws.Cells(3, 15).Value = TickerDecrease
    ws.Cells(3, 16).Value = GreatestDecrease & "%"

    ws.Cells(4, 14).Value = "Greatest Volume: "
    ws.Cells(4, 15).Value = TickerVolume
    ws.Cells(4, 16).Value = GreatestVolume

End Sub

