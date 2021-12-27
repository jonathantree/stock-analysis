# VBA of Wall Street

## Overview of Project

This project developed a macro tool using VBA for Steve to analyze datasets of stock market data. The dataset for this project included stock market data from two years, 2017 and 2018. For each date of the 12 stocks, the closing prices and volumes were used in the analysis. The macro that was designed had a program flow that looped through all of the tickers for each stock, calculated the total daily volume and the rate of return from the starting and ending price over the course of the year for each individual ticker. To ease usability within the workbook, macro buttons to run the analysis given a user input of the year were created for Steve.  

### Purpose

The original script that was written used a nested for loop in order to extract the data for each individual stock ticker which worked well enough for the 12 stocks in the dataset. In the future, Steve would like to run this analysis on thousands of stocks and does not want to have to wait for the analysis to finish running. The purpose of this project was to refactor the script to only use one loop to collect the necesary data to to the analysis and test if the refactored code performed more efficiently than the original.   

## Results

### Stock Performance between 2017 and 2018
The analysis of stocks reveal that over all, most stocks had net gains during the year of 2017 and net losses during the year of 2018 Figures 1 & 2). During 2017, the two stocks with the highest return were DQ (199.4%) and SEDG (184.5%). The only stock with negative returns during 2017 was TERP, which remained negative in 2018. TERP was the only stock which showed negative returns for both years. The analysis of the returns from 2018 showed that DQ had the lowest return (-62.6%). The only stock which had positve year end returns for both 2017 and 2018 was ENPH. If Steve were to suggest an investment option for his parents based on these analyses, ENPH has the best investment potential. 

**Figure 1. Bar chart of 2017 year end return**
![2017 Year-end return chart](./Resources/2017_year_end_return_chart.png)

**Figure 2. Bar chart of 2018 year end return**
![2018 Year-end return chart](./Resources/2018_year_end_return_chart.png)

### Performance enhanced using refactored script

**Figure 3. Line plot showing the performance of the refactored script vs the original over 100 iterations** 
![Line plot performance](./Iteration_Time_Analysis/Refactored_vs_Original_iterationanalysisplot.png)

**Table 1. Statistics of the refactored script and original over 100 iterations**

| **Year**            | **Total Iterations** | **Total Time Elapsed** | **Average Run Time** | **s<sup>2</sup><sub>n-1</sub>**       |
|-----------------|------------------|--------------------|------------------|-------------|
| Refactored 2017 | 100              | 25.546875          | 0.25546875       | 0.019096325 |
| Refactored 2018 | 100              | 25.21484375        | 0.252148438      | 0.026383425 |
| Original 2017   | 100              | 143.140625         | 1.460618622      | 0.141026767 |
| Original 2018   | 100              | 134.703125         | 1.374521684      | 0.176248114 |

## Summary

1. What are the advantages or disadvantages of refactoring code?
   - bndjasbgjs
2. How do these pros and cons apply to refactoring the original VBA script?
   - nasdjkbfsdjk
