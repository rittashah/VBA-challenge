Sub great():

For Each ws In Worksheets


LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Volume"
Volume_Greatest_Decrease = 100000
Ticker_Greatest_Decrease = 100000

    
    For x = 2 To LastRow
    
    
    If ws.Cells(x, 11).Value > Volume_Greatest_Increase Then
        
        Ticker_Greatest_Increase = ws.Cells(x, 9).Value
        Volume_Greatest_Increase = ws.Cells(x, 11).Value
    
    End If
    
    
    If ws.Cells(x, 11).Value < Volume_Greatest_Decrease Then
        
        Ticker_Greatest_Decrease = ws.Cells(x, 9).Value
        Volume_Greatest_Decrease = ws.Cells(x, 11).Value
    
    End If
    
    
    If ws.Cells(x, 12).Value > Volume_Greatest_Total_Volume Then
        
        Ticker_Greatest_Total_Volume = ws.Cells(x, 9).Value
        Volume_Greatest_Total_Volume = ws.Cells(x, 12).Value
    
    End If
    
    Next x

ws.Cells(2, 16).Value = Ticker_Greatest_Increase
ws.Cells(2, 17).Value = Volume_Greatest_Increase
ws.Cells(2, 17).Style = "Percent"
ws.Cells(3, 16).Value = Ticker_Greatest_Decrease
ws.Cells(3, 17).Value = Volume_Greatest_Decrease
ws.Cells(3, 17).Style = "Percent"
ws.Cells(4, 16).Value = Ticker_Greatest_Total_Volume
ws.Cells(4, 17).Value = Volume_Greatest_Total_Volume
ws.Columns("I:Q").EntireColumn.AutoFit


Next ws

End Sub
