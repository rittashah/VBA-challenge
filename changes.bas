Attribute VB_Name = "Module2"
Sub change():

For Each ws In Worksheets


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim minopen As Variant
Dim maxclose As Variant
Dim x As Double


 x = 2
 i = 2
 minopen = ws.Cells(i, 3).Value

ws.Cells(x, 9).Value = ws.Cells(x, 1).Value

minopen = ws.Cells(i, 3).Value

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow


If ws.Cells(i, 1).Value = ws.Cells(x, 9).Value Then

maxclose = ws.Cells(i, 6).Value

 Else

ws.Cells(x, 10).Value = maxclose - minopen

            If maxclose <= 0 Then
        
                ws.Cells(x, 11).Value = 0
                
                Else


ws.Cells(x, 11).Value = (maxclose / minopen) - 1

End If

ws.Cells(x, 11).Style = "Percent"
                    
       If ws.Cells(x, 10).Value >= 0 Then
                            
           ws.Cells(x, 10).Interior.ColorIndex = 4
                                
        Else
                            
        ws.Cells(x, 10).Interior.ColorIndex = 3
                
        End If


minopen = ws.Cells(i, 3).Value

x = x + 1
 ws.Cells(x, 9).Value = ws.Cells(i, 1).Value

End If

Next i

ws.Cells(x, 10).Value = maxclose - minopen

            If DateMaxClose <= 0 Then
        
                ws.Cells(x, 11).Value = 0
                
                Else


ws.Cells(x, 11).Value = (maxclose / minopen) - 1
 
End If
ws.Cells(x, 11).Style = "Percent"
                    
        If ws.Cells(x, 10).Value >= 0 Then
                            
            ws.Cells(x, 10).Interior.ColorIndex = 4
                                
        Else
                            
        ws.Cells(x, 10).Interior.ColorIndex = 3
                
        End If



Columns("I:Q").EntireColumn.AutoFit

Next ws
End Sub
