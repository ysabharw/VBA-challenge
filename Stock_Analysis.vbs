Sub StockAnalysisEnhanced()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Row As Long
    Dim LastRow As Long
    Dim SummaryTableRow As Integer

    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Initialize variables
        TotalVolume = 0
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
        SummaryTableRow = 2
        
        ' Add headers for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        
        ws.Cells(1, 14).Value = "Ticker"
        ws.Cells(1, 15).Value = "Value"

        ws.Cells(2, 13).Value = "Greatest % Increase"
        ws.Cells(3, 13).Value = "Greatest % Decrease"
        ws.Cells(4, 13).Value = "Greatest Total Volume"

        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through rows
        For Row = 2 To LastRow
            ' Check if we are still within the same ticker
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ' Set ticker symbol
                Ticker = ws.Cells(Row, 1).Value
                
                ' Add to total volume
                TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
                
                ' Get open and close prices
                OpenPrice = ws.Cells(Row - (Row - SummaryTableRow), 3).Value
                ClosePrice = ws.Cells(Row, 6).Value
                
                ' Calculate yearly change and percent change
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If

                ' Output to summary table
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                ws.Cells(SummaryTableRow, 10).Value = TotalVolume
                ws.Cells(SummaryTableRow, 11).Value = YearlyChange
                ws.Cells(SummaryTableRow, 12).Value = PercentChange

                ' Apply conditional formatting
                If YearlyChange >= 0 Then
                    ws.Cells(SummaryTableRow, 11).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(SummaryTableRow, 11).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Check for greatest % increase, decrease, and total volume
                If PercentChange > MaxIncrease Then
                    MaxIncrease = PercentChange
                    MaxIncreaseTicker = Ticker
                End If

                If PercentChange < MaxDecrease Then
                    MaxDecrease = PercentChange
                    MaxDecreaseTicker = Ticker
                End If

                If TotalVolume > MaxVolume Then
                    MaxVolume = TotalVolume
                    MaxVolumeTicker = Ticker
                End If

                SummaryTableRow = SummaryTableRow + 1
                TotalVolume = 0

            Else
                TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
            End If
        Next Row

        ' Output greatest values
        ws.Cells(2, 14).Value = MaxIncreaseTicker
        ws.Cells(2, 15).Value = MaxIncrease & "%"
        ws.Cells(3, 14).Value = MaxDecreaseTicker
        ws.Cells(3, 15).Value = MaxDecrease & "%"
        ws.Cells(4, 14).Value = MaxVolumeTicker
        ws.Cells(4, 15).Value = MaxVolume
    Next ws

    MsgBox "Enhanced Analysis complete!"
End Sub
