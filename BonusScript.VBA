Attribute VB_Name = "Module1"
Sub stocks():

    For Each ws In Worksheets
    
        Dim counter As Integer
        counter = 2
        
       'row headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'ws.Cells
        
        'save the first opening price
        Dim statrPrice As Integer
        startPrice = ws.Cells(2, 3).Value 'ws.Cells
        
        'save the first ticker name
        Dim currentTicker As String
        currentTicker = ws.Cells(2, 1).Value 'ws.Cells
        
        Dim totalVolume As LongLong
        totalVolume = 0
        
        For i = 2 To lastRow
    
            If (Not (ws.Cells(i + 1, 1) = ws.Cells(i, 1))) Then 'ws.Cells
            
                'add ticker symbols
                ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value 'ws.Cells
                
                'add totalVolume to the table
                totalVolume = totalVolume + ws.Cells(i, 7).Value 'ws.Cells
                ws.Cells(counter, 12).Value = totalVolume
                totalVolume = 0
                
                'calculate the difference in price
                yearChange = ws.Cells(i, 6).Value - startPrice 'ws.Cells
                ws.Cells(counter, 10).Value = yearChange
                
                'add color to the yearly change cells
                If (yearChange < 0) Then
                    ws.Cells(counter, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(counter, 10).Interior.ColorIndex = 4
                End If
                
                If startPrice <> 0 Then
                    'calculate percent change
                    percentChange = yearChange / startPrice
                    ws.Cells(counter, 11).Value = percentChange
                    yearChange = 0
                    percentChange = 0
                End If
                
                'update the current start price
                startPrice = ws.Cells(i + 1, 3).Value 'ws.Cells
    
                'update the counter for next row
                counter = counter + 1
                
            Else
                'calculate total stock volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
             
        Next i
        
    Next ws
            

End Sub



