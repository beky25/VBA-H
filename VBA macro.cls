
 Sub stock_Market_analysis()
  Dim Ticker As String
  Dim Totalstockv As Double
  Totalstockv = 0
 
  Dim summaryRow As Integer
  summaryRow = 2
  
For i = 2 To 71226
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 Ticker = Cells(i, 1).Value
 Totalstockv = TotalstockVol + Cells(i, 7).Value
 
 Range("I" & summaryRow).Value = Ticker
 Range("J" & summaryRow).Value = Totalstockv
  summaryRow = summaryRow + 1
  Totalstock = 0
  Else
  Totalstockv = Totalstockv + Cells(i, 7).Value
  End If
  
Next

 End Sub
