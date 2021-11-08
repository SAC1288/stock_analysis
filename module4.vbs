Sub AllStockAnalysis()

    'declaring key variables
    Dim startTime As Single
    Dim endTime As Single
    Dim tickerIndex As Integer
    Dim yearValue As String
    
    'declaring key arrays
    Dim totalVolume() As Long
    Dim ticker() As String
    Dim endingPrice() As Double
    Dim startingPrice() As Double
    
    Worksheets("All Stocks Analysis").Activate
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'starting the clock on the performance of the sub-routine.
    startTime = Timer
    Sheets(yearValue).Activate

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
   

    For i = 2 To RowCount
        
        'At the point when values in the cells of column A change, increase each array by 1
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            'making key arrays for all four stock arrays dynamic so that they can take on a variety of stocks instead of the 12 that we currently have for both years for future analyses.
            ReDim Preserve ticker(tickerIndex + 1)
            ReDim Preserve totalVolume(tickerIndex + 1)
            ReDim Preserve startingPrice(tickerIndex + 1)
            ReDim Preserve endingPrice(tickerIndex + 1)
            If tickerIndex > 0 Then
                'when tickerIndex is greater than zero, this means that the value of the ticker within Column A has changed so please include the closing price of the last row for the previous ticker.
                endingPrice(tickerIndex) = Cells(i - 1, 6).Value
            End If
            
            'After increasing the size of each array, then increase the value of tickerIndex.
            tickerIndex = tickerIndex + 1
                
            'Next, store the actual ticker value in ticker(tickerIndex). This means that unti the cells in Column A change to a different value, ticker(tickerIndex) will contain the same ticker string value.
            ticker(tickerIndex) = Cells(i, 1).Value
            'store the value of the first opening price in startingPrice(tickerIndex). This changes each time the ticker value changes for column A.
            startingPrice(tickerIndex) = Cells(i, 3).Value
                       
        End If
        
        'For every interation of i, please increase the totalVolume array by the total volume for each row for a specific ticker until the ticker value changes in column A which will be denoted by the change in tickerIndex's value.
        totalVolume(tickerIndex) = totalVolume(tickerIndex) + Cells(i, 8).Value
         
    Next i
    'For the last row of the data set, include the closing price value within the endingPrice(tickerIndex) variable.
    endingPrice(tickerIndex) = Cells(RowCount, 6).Value
          
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'run a for loop from 1 to the upperbound of the variable ticker
    For j = 1 To UBound(ticker)
    
        Cells(3 + j, 1).Value = ticker(j)
        Cells(3 + j, 2).Value = totalVolume(j)
        Cells(3 + j, 3).Value = endingPrice(j) / startingPrice(j) - 1
        
    Next j
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    'Stopped the clock on the performance of the sub-routine and measured the time.

End Sub


