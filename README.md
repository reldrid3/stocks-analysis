# Analysis of Stocks

## Overview of Project
After analyzing a dataset, Steve Smith asked me to refactor the VBA code to work for larger amounts of data, comparing the time for the script to complete.

### Current Code

The current code keeps track of each ticker and outputs the data by cycling through the entire sheet each time.  This is inefficient - if there were many more than 12 tickers, this would be a significant slowdown - each ticker would need to cycle through the rows of the sheet each time, leading to an O(n<sup>2</sup>)

## Results

## Summary
