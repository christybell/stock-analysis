# Analyzing Stock Data with VBA

## Overview of Project
My friend Steve loves the workbook I prepared for him to help his parents analyze the performance of a dozen stocks. Now, he wants me to expand this initial research to include data from the entire stock market over the last few years.

### Purpose of analysis
Although the original code I prepared works well for a dozen stocks, it might not work as well for thousands of stocks like Steve wants. And if it does, it may take a long time to execute the script. Therefore the purpose of this project is to refactor the existing code and see if I can make it run more efficiently by measuring its performance against the original script's performance.

## Results
In order to make my code run more efficiently, I needed to switch the nesting order of my `for` loops. To achieve this, I first created a variable called the `tickerIndex` that accessed the correct index across the `tickers` array as well as the 3 new output arrays I created: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. The `tickerIndex` variable meant my code was able to assign the `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` to each ticker symbol before looping through the entire dataset. Refactoring the code in this manner made it run much faster than using the original script.

### Run-Times for the Original vs. Refactored Scripts

The run-times for the original script for 2017 and 2018 were as follows:

<img src="resources/Original_Run_Time_2017">
<img src="resources/Original_Run_Time_2018.png">

## Summary
1. What are the advantages or disadvantages of refactoring code in general?
2. How do these pros and cons apply to refactoring the original VBA script?
