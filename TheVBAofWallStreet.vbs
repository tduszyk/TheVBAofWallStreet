Sub VBAofWallStreet()
'Loop through all sheets
For Each ws In Worksheets
    Dim i, j, TotalStockCount, SumaryTableRow As Integer
    Dim Ticker As String
    'Create cells headers
    Range("I1") = "Ticker"
    Range("J1") = "Total Stock Count"
    SummaryTableRow = 2
    'Find last row and loop through it
    For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            TotalStockCount = TotalStockCount + Cells(i, 7).Value
            Range("I" & SummaryTableRow).Value = Ticker
            Range("J" & SummaryTableRow).Value = TotalStockCount
            SummaryTableRow = SummaryTableRow + 1
            TotalStockCount = 0
        Else
            TotalStockCount = TotalStockCount + Cells(i, 7).Value
        End If
    Next i
Next ws
End Sub