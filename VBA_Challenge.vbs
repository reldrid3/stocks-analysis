Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    Dim numTickers As Integer

    YearValue = InputBox("What year would you like to run the analysis on?")
    
    'Error Checking to make sure Sheet 'yearValue' exists, and exit the subroutine if it does not.
    If Not Evaluate("ISREF('" & YearValue & "'!A1)") Then
    
        MsgBox ("ERROR - Sheet '" & YearValue & "' does not exist!")
        Exit Sub
        
    End If

    'Start the timer at the current time
    startTime = Timer
    
    'Set the number of tickers in the data
    numTickers = 12
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers - ReDim is required, as
     'VBA will not allow a variable as an array length in a Dim statement
    Dim tickers() As String
    ReDim tickers(numTickers)
    
    'tickers = WorksheetFunction.Unique(Range("A2:A" & lastRow))
    
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
    Worksheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create (and ReDim) three output arrays
    
    Dim tickerVolumes() As Long
    ReDim tickerVolumes(numTickers)
    Dim tickerStartingPrices() As Single
    ReDim tickerStartingPrices(numTickers)
    Dim tickerEndingPrices() As Single
    ReDim tickerEndingPrices(numTickers)
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To numTickers - 1
        tickerVolumes(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6)
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6)
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    Worksheets("All Stocks Analysis").Activate
    Cells(1, 1) = "All Stocks(" & YearValue & ")"
    Cells(3, 1) = "Ticker"
    Cells(3, 2) = "Total Daily Volume"
    Cells(3, 3) = "Return"
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To numTickers - 1
        
        Cells(i + 4, 1) = tickers(i)
        Cells(i + 4, 2) = tickerVolumes(i)
        Cells(i + 4, 3) = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)

End Sub
