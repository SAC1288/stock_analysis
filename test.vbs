Sub AllStockAnalysis()

    Worksheets("All Stocks Analysis").Activate

    tickerIndex = 0
    totalVolume = 0
    Dim startTime As Single
    Dim endTime As Single
    Dim yearValue As String
    Dim endingPrice As Double
    Dim startingPrice As Double
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    Sheets(yearValue).Activate

    For i = 2 To RowCount
    
        totalVolume = totalVolume + Cells(j, 8).Value
        
        If Cells(j - 1, 1).Value <> Cells(j, 1).Value And Cells(j, 1) = Cells(j + 1, 1) Then
            startingPrice = Cells(j, 3).Value
            End If
            
        If Cells(j + 1, 1).Value <> Cells(j, 1).Value And Cells(j, 1).Value = Cells(j - 1, 1).Value Then
            endingPrice = Cells(j, 6)

            sheets("All Stocks Analysis").Cells(tickerIndex, "A").Value = sheets(yearValue).cells(j,"A").Value 
            sheets("All Stocks Analysis").Cells(tickerIndex, "B").Value = totalVolume
            sheets("All Stocks Analysis").Cells(tickerIndex, "C").Value = endingPrice / startingPrice - 1 
            totalVolume = 0
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
        
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
  
End Sub

