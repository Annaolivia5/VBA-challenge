Attribute VB_Name = "Module1"
Sub stocks():

Dim counter As Integer
counter = 2

'row headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

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
    
    If (Not (Cells(i + 1, 1) = Cells(i, 1))) Then
            
        'add ticker symbols
        Cells(counter, 9).Value = Cells(i, 1).Value
                
        'add totalVolume to the table
        totalVolume = totalVolume + Cells(i, 7).Value
        Cells(counter, 12).Value = totalVolume
        totalVolume = 0
                
        'calculate the year change for price
        yearChange = Cells(i, 6).Value - startPrice
        Cells(counter, 10).Value = yearChange
                
        'add color to the year change cells
        If (yearChange < 0) Then
            Cells(counter, 10).Interior.ColorIndex = 3
        Else
            Cells(counter, 10).Interior.ColorIndex = 4
        End If
                
        If startPrice <> 0 Then
            'calculate percent change
            percentChange = yearChange / startPrice
            Cells(counter, 11).Value = percentChange
            Cells(counter, 11).NumberFormat = " %0.00"
            yearChange = 0
            percentChange = 0
        End If
                
        'update the start price for next stock
        startPrice = Cells(i + 1, 3).Value
    
        'update the counter for next row
        counter = counter + 1
                
    Else
    
        'calculate total stock volume
        totalVolume = totalVolume + Cells(i, 7).Value
    End If
             
Next i
    

End Sub



