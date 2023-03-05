Attribute VB_Name = "Module1"
Sub stocks():

For Each ws In Worksheets


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change" 'YC
ws.Cells(1, 11).Value = "Percentage Change" 'PC
ws.Cells(1, 12).Value = "Total Stocks volume" ' TCV
Dim ticker As String

Dim TCV As Double
TCV = 0

Dim SummaryRow As Integer
SummaryRow = 2
 For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ticker = ws.Cells(i, 1).Value
      TCV = TCV + ws.Cells(i, 7).Value
     
      
      ws.Range("I" & SummaryRow).Value = ticker
      ws.Range("L" & SummaryRow).Value = TCV
    
      SummaryRow = SummaryRow + 1
      TCV = 0
     
      
    Else
      TCV = TCV + ws.Cells(i, 7).Value
    
      
     
End If
Next i

 ws.Columns("I:Q").EntireColumn.AutoFit
 Next ws
 
End Sub
