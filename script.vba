Attribute VB_Name = "Module1"
Sub stocks():

Dim counter As Integer
counter = 2

    'row headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'    For Each ws In Worksheets
    
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row 'ws.Cells
        
        'save the first opening price
        Dim statrPrice As Integer
        startPrice = Cells(2, 3).Value 'ws.Cells
        
        'save the first ticker name
        Dim currentTicker As String
        currentTicker = Cells(2, 1).Value 'ws.Cells
        
        Dim totalVolume As LongLong
        totalVolume = 0
        
        For i = 2 To lastRow
    
            If (Not (Cells(i + 1, 1) = Cells(i, 1))) Then 'ws.Cells
            
                'add ticker symbols
                Cells(counter, 9).Value = Cells(i, 1).Value 'ws.Cells
                
                'add totalVolume to the table
                totalVolume = totalVolume + Cells(i, 7).Value 'ws.Cells
                Cells(counter, 12).Value = totalVolume
                totalVolume = 0
                
                'calculate the difference in price
                yearChange = Cells(i, 6).Value - startPrice 'ws.Cells
                Cells(counter, 10).Value = yearChange
                
                'add color to the yearly change cells
                If (yearChange < 0) Then
                    Cells(counter, 10).Interior.ColorIndex = 3
                Else
                    Cells(counter, 10).Interior.ColorIndex = 4
                End If
                
                If startPrice <> 0 Then
                    'calculate percent change
                    percentChange = yearChange / startPrice
                    Cells(counter, 11).Value = percentChange
                    yearChange = 0
                    percentChange = 0
                End If
                
                'update the current start price
                startPrice = Cells(i + 1, 3).Value 'ws.Cells
    
                'update the counter for next row
                counter = counter + 1
                
            Else
                'calculate total stock volume
                totalVolume = totalVolume + Cells(i, 7).Value
            End If
             
        Next i
    

End Sub




