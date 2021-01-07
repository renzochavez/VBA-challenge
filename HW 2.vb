Sub stockmarket()
    Dim openvalue, closevalue, percent_change, volume As Double
    Dim lrow As Long
    Dim ticker As String
    Dim start, newrow As Integer
    
    For Each ws In Worksheets:
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
   
    
    volume = 0
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    start = 2
    newrow = 2
    
    For i = 2 To lrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            
            volume = volume + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            ws.Cells(newrow, 9).Value = ticker
            ws.Cells(newrow, 12).Value = volume
            
        
            openvalue = ws.Cells(start, 3).Value
            closevalue = ws.Cells(i, 6).Value
            
            ws.Range("J" & newrow).Value = closevalue - openvalue
            
            If openvalue <> 0 Then
            ws.Range("K" & newrow).Value = (closevalue - openvalue) / (openvalue)
            Else
            ws.Range("K" & newrow).Value = (closevalue - openvalue) / 1
            End If
            
            ws.Range("K" & newrow).NumberFormat = "0.00%"
            
            If closevalue - openvalue > 0 Then
                ws.Cells(newrow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(newrow, 10).Interior.ColorIndex = 3
            End If
            
            start = i + 1
            volume = 0
            newrow = newrow + 1
        Else
        
            volume = volume + ws.Cells(i, 7).Value
            
        End If
    Next i

    Next

        

End Sub
