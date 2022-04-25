# Analysis of Stocks

## Overview of Project
After analyzing a dataset, Steve Smith asked me to refactor the VBA code to work for larger amounts of data, comparing the time for the script to complete.

### Current Code

The current code keeps track of each ticker and outputs the data by cycling through the entire sheet each time.  This is inefficient - if there were many more than 12 tickers, this would be a significant slowdown - each ticker would need to cycle through the rows of the sheet each time, leading to an O(mn) code efficiency.

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
Using a tickerIndex, I can keep track of the current ticker, allowing us to collect all the tickers in one go, incrementing the ticker only at the end of a particular stock.
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

#### Time Comparisons

The original runtimes of 0.50s and 0.48s from the Module can be seen screenshotted in [Figure 1](#figure-1).  Compared to the [Figure 2](#figure-2) runtimes of 0.047s, and there is a more than 10x difference in the overall runtimes.  This makes sense, as instead of looping through the rows 12 times, it only loops over 1 time.

#### Stock Comparisons

Clearly, 2017 and 2018 had often vastly different stock market results -- only 1 stock even fell in 2017, while only 2 stocks went up in 2018 -- but there are some conclusions to be drawn.  Refer to [Figure 3](#figure-3) for a net gain/loss calculation for 2017 and 2018 combined.
- The only two stocks which rose in both years were 'ENPH' and 'RUN,' implying that those are strong stocks even during an economic downturn.  'ENPH' also had the best performance of any of the other stocks, hands down.
- Another stock to watch is 'SEDG.'  Although it fell by 7.8% in 2018, it's 2017 gains of 184.5% far outweigh it, actually making it a better overall (162.4%) investment than 'RUN' (net gain of 94.2%).
- Stock 'DQ,' with great performance in 2017, lost almost all of it's gains in 2018, impyling that it did not do well with the market changes in 2018.

## Refactoring Summary

Refactoring code comes with numerous advantages.  The obvious advantage would be to decrease the runtime of the script, increasing its efficiency and ability to handle larger amounts of input.  Refactoring also comes with the advantage of being able to dissect the code's inner workings, not only working to increase efficiency but also potentially adding robustness, error-checking, or additional features (for a new version, hopefully remaining backwards compatible - this is beyond "refactoring" but allows insight into improving the code).  The only disadvantage is that it is possible to introduce additional errors into the code, even potentially silent errors, so making sure that you test and confirm your output and functionality is important.

In the case here, I increased the efficiency of the code by more than 10-fold given 12 stocks, implying that with any number of stocks, the runtime will still be based mostly on the total number of rows, rather than the number of different stocks.  As well, while refactoring, I also added error checking to the user input, allowing the script to exit safely should the user enter a year that doesn't match an existing worksheet.

# Figures

## Figure 1
![Runtimes for the original code.](/Resources/VBA_Challenge_ModuleTimes.png)

## Figure 2
![Results and runtime for refactored code, 2017 data.](/Resources/VBA_Challenge_2017.png)
![Results and runtime for refactored code, 2018 data.](/Resources/VBA_Challenge_2018.png)

## Figure 3
![Net stock change from 2017 to 2018.](/Resources/VBA_Challenge_2017-2018.png)
