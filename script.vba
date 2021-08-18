Attribute VB_Name = "Module1"
Sub stocks():

Dim counter As Integer
counter = 2

    'row headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    For Each ws In Worksheets
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'save the first opening price
        Dim statrPrice As Integer
        startPrice = ws.Cells(2, 3).Value
        
        'save the first ticker name
        Dim currentTicker As String
        currentTicker = ws.Cells(2, 1).Value
        
        Dim totalVolume As LongLong
        totalVolume = 0
        
        For i = 2 To lastRow
    
            If (Not (ws.Cells(i + 1, 1) = ws.Cells(i, 1))) Then
            
                'add ticker symbols
                Cells(counter, 9).Value = ws.Cells(i, 1).Value
                
                'add totalVolume to the table
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                Cells(counter, 12).Value = totalVolume
                totalVolume = 0
                
                'calculate the difference in price
                yearChange = startPrice - ws.Cells(i, 3).Value
                Cells(counter, 10).Value = yearChange
                
                If startPrice <> 0 Then
                    'calculate percent change
                    percentChange = yearChange / startPrice
                    Cells(counter, 11).Value = percentChange
                    yearChange = 0
                    percentChange = 0
                End If
                
                'update the current start price
                startPrice = ws.Cells(i + 1, 3).Value
    
                'update the counter for next row
                counter = counter + 1
                
            Else
                'calculate total stock volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
             
        Next i

    Next ws
    

End Sub



