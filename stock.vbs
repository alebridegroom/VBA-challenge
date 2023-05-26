Sub VbaStocks()
For Each ws In Worksheets
    Dim opening As Double
    'setting the opening to be the first value where open is
    opening = ws.Cells(2, 3).Value
    Dim closing As Double
    
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'range of where the for loop will be 
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    'header cells place value
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Value"
    
    
    For i = 2 To RowCount
    'if the ticker values don't equal to each other
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            'closing value will be the value at open where the tickers don't match
            closing = ws.Cells(i, 6).Value
            
            YearlyChange = closing - opening
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            'setting the colors
            If YearlyChange >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            PercentChange = (closing / opening) - 1
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            'formatting the percent change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            'setting the new open value
            opening = ws.Cells(i + 1, 3).Value
            'adding the volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
            TotalStockVolume = 0
            
            'adding to the row
            Summary_Table_Row = Summary_Table_Row + 1
        Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
        End If
    Next i
    
    ' take the max and min and place them in a separate part in the worksheet
    ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    Decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    Volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    ws.Range("O2") = Cells(increase_number + 1, 9)
    ws.Range("O3") = Cells(Decrease_number + 1, 9)
    ws.Range("O4") = Cells(Volume_number + 1, 9)
    
    Next ws

End Sub
