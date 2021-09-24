Sub allStocksAnalysisYear()
    'setting two variables to have a timer check the code execution
    Dim startTime As Single
    Dim endTime As Single
       
    
    
    
    'get input from user using an input box
    Dim yearValue As String
    yearValue = 0
    
    yearValue = InputBox("what year would you like to run the analysis on?")
    
    startTime = Timer
    
    'Format the output sheet on the "All Stocks Analysis" worksheet.
   
    Worksheets("All Stocks Analysis").Activate
    Cells(1, 1).Value = "All Stocks Anaysis (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker Symbol"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    
    'Prepare the analysis of tickers this is an rray of ticker in 2018 data, Initialize an array of all tickers.
        Dim tickers(11) As String
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
    
    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
      
    'Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'Loop through the tickers.
    
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
                
           'Loop through rows in the data to find total volume
           Worksheets(yearValue).Activate
           For j = 2 To RowCount
              'finds total volume of ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
            'Find the start and end price of the current ticker.
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
           Next j
    
    'output data of current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
  
  Next i
    
    Worksheets("All Stocks Analysis").Activate
    Range("A3:B3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuius
    Range("A4:A15").Font.Italic = True
    Cells(3, 3).Font.Underline = True
    Cells(3, 3).Font.Italic = True
    Range("A4:A15").Font.Color = RGB(255, 0, 85)
    Range("C4:c15").NumberFormat = "0.00%"
    Range("B4:B15").NumberFormat = "#,##0"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
          'color the cell green if > 0
          Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
         'color the cell red of < 0
         Cells(i, 3).Interior.Color = vbRed
        Else
            'clear cthe cell color
            Cells(i, 3).Interior.Color = xlNone
        
        End If
    Next i
    
   endTime = Timer
   
   MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
   
End Sub


Sub formatAllStocksAnalysisTable()
    Worksheets("All Stocks Analysis").Activate
    Range("A3:B3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuius
    Range("A4:A15").Font.Italic = True
    Cells(3, 3).Font.Underline = True
    Cells(3, 3).Font.Italic = True
    Range("A4:A15").Font.Color = RGB(255, 0, 85)
    Range("C4:c15").NumberFormat = "0.00%"
    Range("B4:B15").NumberFormat = "#,##0"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
          'color the cell green if > 0
          Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
         'color the cell red of < 0
         Cells(i, 3).Interior.Color = vbRed
        Else
            'clear cthe cell color
            Cells(i, 3).Interior.Color = xlNone
        
        End If
    Next i
End Sub










