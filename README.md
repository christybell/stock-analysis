# Analyzing Stock Data with VBA

## Overview of Project
My friend Steve loves the workbook I prepared for him to help his parents analyze the performance of a dozen stocks in 2017 and 2018. Now, he wants me to expand this initial research to include data from the entire stock market over the last few years.

### Purpose of analysis
Although the original code I prepared works well for a dozen stocks, it might not work as well for thousands of stocks like Steve wants. And if it does, it may take a long time to execute the script. Therefore the purpose of this project is to refactor the existing code and see if I can make it run more efficiently by measuring its performance against the original script's performance.

## Results
In order to make my code run more efficiently, I needed to switch the nesting order of my `for` loops. To achieve this, I first created a variable called the `tickerIndex` that accessed the correct index across the `tickers` array as well as the 3 new output arrays I created: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. Setting up the `tickerIndex` variable meant my code was able to assign the `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` to each ticker symbol before looping through the entire dataset. Refactoring the code in this manner allowed it to run much faster than using the original script.

### Run-Times for the Original vs. Refactored Scripts

The run-times for the original script for 2017 and 2018 were as follows:

<img src="Resources/Original_Run_Time_2017.png">
<img src="Resources/Original_Run_Time_2018.png">


The run-times for the refactored script for 2017 and 2018 were as follows:

<img src="Resources/VBA_Challenge_2017.png">
<img src="Resources/VBA_Challenge_2018.png">

Comparing the sets of screenshots above, I determined the refactored code runs about 0.5 seconds faster than the original code. 

## Summary

### General Advantages and Disadvantages of refactoring code
In general terms, there are several advantages to refactoring code. It can help uncover programming bugs, makes programs run faster, and it's easier to understand and interpret the code. Two disadvantages of refactoring code is it can be time consuming and it also requires a lot of skill and discipline.

### Pros and Cons to refactoring the original VBA script
As for this challenge assignment specifically, refactoring the original VBA script increased its efficiency and run-time performance. However, refactoring the code was time consuming and since I lack programming knowledge and skill, when there were syntax errors, it was challenging to debug the code.
