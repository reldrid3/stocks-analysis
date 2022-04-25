# Analysis of Stocks

## Overview of Project
After analyzing a dataset, Steve Smith asked me to refactor the VBA code to work for larger amounts of data, comparing the time for the script to complete.

### Current Code

The current code keeps track of each ticker and outputs the data by cycling through the entire sheet each time.  This is inefficient - if there were many more than 12 tickers, this would be a significant slowdown - each ticker would need to cycle through the rows of the sheet each time, leading to an O(n<sup>2</sup>) code efficiency.

### Proposed Changes

- Add in an input box to allow different years or sheets to be run, rather than hardcoding the year
- Refactor the code to loop through the tickers only once, keeping track of all the ticker information concurrently.  This will result in an O(n) efficiency and significant time improvement.

## Results

### Input Box
Code to add in an input box was simple enough, even with added error checking for a valid sheet.
```
YearValue = InputBox("What year would you like to run the analysis on?")
    
    'Error Checking to make sure Sheet 'yearValue' exists, and exit the subroutine if it does not.
    If Not Evaluate("ISREF('" & YearValue & "'!A1)") Then
    
        MsgBox ("ERROR - Sheet '" & YearValue & "' does not exist!")
        Exit Sub
        
    End If
```

### Refactoring
Using a tickerIndex, we can keep track of the current ticker, allowing us to collect all the tickers in one go.
```
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
```

## Summary
