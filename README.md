# VBA-challenge
Had help with this part with bootcamp suppor, they gave me a basic format and i changed the range values. 
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
