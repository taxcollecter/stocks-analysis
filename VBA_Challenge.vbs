Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
       For i = 0 To 11
            ticker = tickers(i)

    '1b) Create three output arrays
            Dim tickerVolumes(11) As Long
            Dim tickerStartingPrices(11) As Single
            Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
            totalVolume = 0
            
    ''2b) Loop over all the rows in the spreadsheet.
                 For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
                    If Cells(j, 1).Value = tickers(i) Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                    End If
                    
        '3b) Check if the current row is the first row with the selected tickerIndex.
                    If Cells(j, 1).Value <> Cells(j - 1, 1).Value And Cells(j, 1).Value = tickers(i) Then
                    tickerStartingPrices(i) = Cells(j, 3).Value
                    End If
                    
        '3c) check if the current row is the last row with the selected ticker
                    If Cells(j, 1).Value <> Cells(j + 1, 1).Value And Cells(j, 1).Value = tickers(i) Then
                        'MsgBox (totalVolume)
                        tickerVolumes(i) = totalVolume
                        totalVolume = 0
                        tickerEndingPrices(i) = Cells(j, 6).Value
                        
            '3d Increase the tickerIndex.
                        'ticker = ticker + 1
                        'Not needed?
                        End If
                  Next j
        Next i

  
'Format Worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A3:C3").Font.FontStyle = "Bold"
   Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
   Range("B4:B15").NumberFormat = "#,##0"
   Range("C4:C15").NumberFormat = "0.0%"
   Columns("B").AutoFit

   Range("A1").Value = "All Stocks (" & yearValue & ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"
   
   'Loop and Place values from Arrays
    For k = 0 To 11
        Cells(k + 4, 1).Value = tickers(k)
        Cells(k + 4, 2).Value = tickerVolumes(k)
        Cells(k + 4, 3).Value = (tickerEndingPrices(k) / tickerStartingPrices(k)) - 1
    Next k

    'Add Color Formatting via Loop
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i


endTime = Timer
MsgBox ("Elapsed Run time for " & yearValue & " Analysis (Milliseconds): " & (endTime - startTime))
End Sub
