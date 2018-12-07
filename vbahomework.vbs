Sub vbahomework()

    'Make each worksheet loop start
    
    For Each ws In Worksheets
    
    'assign position on worksheet for ticker and total stock volume
    
    ws.Cells(1, 9).Value = "ticker"
    
    ws.Cells(1, 10).Value = "Total Stock Volume"
    
    Dim totalvolume As Single
    
    Dim counter As Integer
    
    'add counter for next in line
    
    counter = 2
    
    'cycle through each
    
    For i = 2 To 43398
    
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        
        'if value isnt equal this means new stock
        
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ws.Cells(counter, 10).Value = totalvolume
        totalvolume = 0
        counter = counter + 1
        
    End If
    
    Next i
    
    'next worksheet
    
Next ws

End Sub
