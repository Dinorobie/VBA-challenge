Sub VBA_HW():
For Each ws In Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Value"
 
Dim tickername As String
Dim totalvolume As Double
totalvolume = 0
Dim table_row As Integer
Dim year_open, year_close, yearly_change, percent_change As Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
year_open = ws.Cells(2, 3).Value

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      tickername = ws.Cells(i, 1).Value
      ws.Range("I" & Summary_Table_Row).Value = tickername
      totalvolume = totalvolume + ws.Cells(i, 7).Value
      ws.Range("L" & Summary_Table_Row).Value = totalvolume
      year_close = ws.Cells(i, 6).Value
      yearly_change = year_close - year_open
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      If yearly_change < 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
        
        If year_open = 0 Then
        percent_change = yearly_change
        Else
        percent_change = yearly_change / year_open
        End If
        ws.Range("K" & Summary_Table_Row).Value = percent_change

      
      Summary_Table_Row = Summary_Table_Row + 1
      year_open = ws.Cells(i + 1, 3).Value
      totalvolume = 0
    Else
        totalvolume = totalvolume + ws.Cells(i, 7).Value
      
    End If
      
Next i



Next ws
End Sub
