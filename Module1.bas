Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim SummaryTableRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    Dim CellValue As Variant

    ' Loop through all worksheets
    For Each ws In Worksheets
        ws.Activate
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        SummaryTableRow = 2
        
        ' Add headers to the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Initialize variables for the greatest values
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        ' Initialize the first ticker
        Ticker = ws.Cells(2, 1).Value
        OpenPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        
        ' Loop through all rows in the worksheet
        For i = 2 To LastRow
            ' Check if we are still within the same ticker symbol
            If ws.Cells(i + 1, 1).Value <> Ticker Then
                ' If not, it's the end of the quarter for this ticker
                ClosePrice = ws.Cells(i, 6).Value
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                QuarterlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Add the values to the summary table
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                ws.Cells(SummaryTableRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryTableRow, 11).Value = PercentChange
                ws.Cells(SummaryTableRow, 12).Value = TotalVolume
                
                ' Apply conditional formatting
                If QuarterlyChange > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(SummaryTableRow, 10).Interior.Color = vbRed
                End If
                
                ' Check for greatest values
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If
                
                ' Move to the next row in the summary table
                SummaryTableRow = SummaryTableRow + 1
                
                ' Reset the variables for the next ticker
                Ticker = ws.Cells(i + 1, 1).Value
                OpenPrice = ws.Cells(i + 1, 3).Value
                TotalVolume = 0
            Else
                ' If we are still within the same ticker, accumulate the volume
                CellValue = ws.Cells(i, 7).Value
                If IsNumeric(CellValue) Then
                    TotalVolume = TotalVolume + CellValue
                Else
                    Debug.Print "Non-numeric value encountered at row " & i & ": " & CellValue
                End If
            End If
        Next i
        
        ' Output the greatest values
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = GreatestIncreaseTicker
        ws.Cells(2, 16).Value = GreatestIncrease
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = GreatestDecreaseTicker
        ws.Cells(3, 16).Value = GreatestDecrease
        
        ws.Cells(4, 14).Value = "Greatest Volume"
        ws.Cells(4, 15).Value = GreatestVolumeTicker
        ws.Cells(4, 16).Value = GreatestVolume
        
    Next ws

End Sub


