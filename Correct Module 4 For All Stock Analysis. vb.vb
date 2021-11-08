Sub AllStockAnalysis()

    
    Dim startTime As Single
    Dim endTime As Single
    Dim tickerIndex As Integer
    Dim yearValue As String
    
    Dim totalVolume() As Long
    Dim ticker() As String
    Dim endingPrice() As Double
    Dim startingPrice() As Double
    
    Worksheets("All Stocks Analysis").Activate
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    Sheets(yearValue).Activate

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
   

    For i = 2 To RowCount
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            ReDim Preserve ticker(tickerIndex + 1)
            ReDim Preserve totalVolume(tickerIndex + 1)
            ReDim Preserve startingPrice(tickerIndex + 1)
            ReDim Preserve endingPrice(tickerIndex + 1)
            If tickerIndex > 0 Then
                endingPrice(tickerIndex) = Cells(i - 1, 6).Value
            End If
            
            tickerIndex = tickerIndex + 1
                
            
            ticker(tickerIndex) = Cells(i, 1).Value
            startingPrice(tickerIndex) = Cells(i, 3).Value
            
            
        End If
        
        totalVolume(tickerIndex) = totalVolume(tickerIndex) + Cells(i, 8).Value
         
    Next i
    endingPrice(tickerIndex) = Cells(RowCount, 6).Value
          
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
   
    
    
    
    For j = 1 To UBound(ticker)
    
    
    
    
        Cells(3 + j, 1).Value = ticker(j)
        Cells(3 + j, 2).Value = totalVolume(j)
        
       
        
        Cells(3 + j, 3).Value = endingPrice(j) / startingPrice(j) - 1
        
    Next j
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


