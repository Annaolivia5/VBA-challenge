Attribute VB_Name = "Module2"
Sub stocks():

    For Each ws In Worksheets
    
        Dim greatestIncrease As Double
        greatestIncrease = 0
        Dim greatestIncreaseTicker As String
    
        Dim greatestDecrease As Double
        greatestDecrease = 0
        Dim greatestDecreaseTicker As String
        
        Dim greatestTotalVolume As LongLong
        greatestTotalVolume = 0
        Dim greatestTotalVolumTicker As String
        
        Dim counter As Integer
        counter = 2
        
       'row headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

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
                ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
                
                'add totalVolume to the table
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ws.Cells(counter, 12).Value = totalVolume
                
                'update greatest total volume if needed
                If (totalVolume > greatestTotalVolume) Then
                        greatestTotalVolumeTicker = ws.Cells(i, 1).Value
                        greatestTotalVolume = totalVolume
                End If
                    
                totalVolume = 0
                
                'calculate the year change for price
                yearChange = ws.Cells(i, 6).Value - startPrice
                ws.Cells(counter, 10).Value = yearChange
                
                'add color to the year change cells
                If (yearChange < 0) Then
                    ws.Cells(counter, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(counter, 10).Interior.ColorIndex = 4
                End If
                
                If startPrice <> 0 Then
                    'calculate percent change
                    percentChange = yearChange / startPrice
                    ws.Cells(counter, 11).Value = percentChange
                    ws.Cells(counter, 11).NumberFormat = " %0.00"
                    
                    
                    'update greatest increase if needed
                    If (percentChange > greatestIncrease) Then
                        greatestIncreaseTicker = ws.Cells(i, 1).Value
                        greatestIncrease = percentChange
                    End If
                    
                    'update greatest decrease if needed
                    If (percentChange < greatestDecrease) Then
                        greatestDecreaseTicker = ws.Cells(i, 1).Value
                        greatestDecrease = percentChange
                    End If
                    
                    yearChange = 0
                    percentChange = 0
                End If
                
                'update the start price for next stock
                startPrice = ws.Cells(i + 1, 3).Value 'ws.Cells
    
                'update the counter for next row
                counter = counter + 1
                
            Else
            
                'calculate total stock volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
             
        Next i
        
        
        'Bonus table
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(2, 16).NumberFormat = " %0.00"
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(3, 16).NumberFormat = " %0.00"
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatestTotalVolumeTicker
        ws.Cells(4, 16).Value = greatestTotalVolume


    Next ws
            

End Sub

