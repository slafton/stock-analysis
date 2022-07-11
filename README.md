# All Stocks Analysis Efficiently

## Overview

### Purpose

#### The purpose of this analysis is to be able to analyze an entire dataset efficiently. We refactored existing code to reduce the run time of the code. This will be important as the data set grows and the client wants to quickly make comparisons. To do this we refactored the code so that the code would loop through all the data one time to collect the same information.

### Background

#### This data set is about stocks from 2017 and 2018 and how they are performed. We are looking at 12 different tickers and have their opening price, closing price, adjusted close, volume  and their high and low prices to draw conclusions from. We use this data set to make a table showing the total daily volume and the return for each ticker over the year that we run our code for. By gaining this insight from the table the client will be able to make good decisions on which tickers are better to invest in.

## Results

#### Stock Performance

#### 
#### From these tables we can see that the year 2017 was a bull market and the year 2018 was a bear market. Tickers ENPH and RUN both had positive returns both years. However, buying low and selling high is the core of successful investing. This means that the tickers with prices that are currently low may be a better investment. It is also important to look at the daily volume that was traded. In 2017, SPWR had the most followed by FSLR. In 2018, SPWR had the most followed by ENPH. Generally you want to invest in stock that has a high daily trade volume because you don't want to get stuck with stock. 

### Execution Times

####
#### As shown above, the refactored code execution times are significantly faster. This will save the client time as the data set grows larger. We refactored, or edited, the code so that there was only one loop to make it faster. 

## Summary

### Advantages and Disadvantages of Refactoring Code

#### Refactoring code can allow for better comprehensibility and may facilitate maintenance and extendibility of the software; however, imprecise refactoring code could introduce new bugs and errors into the code. It may seem faster to reuse old code or copy code but you need to take the time to understand and make sure it is applicable to your needs. Also it is important to make small changes and test the code as you go to lessen the number of bugs you have as you go. 

### Advantages and Disadvantages of Original and Refactored VBA Script

#### Refactoring our original code reduced the amount of code that we needed to write; however, I had to work out several bugs in the refactored code that took a lot of time to fix. We set more arrays and reduced the number of loops. In our refactored code the tickerIndex is set equal to zero before looping over the rows. We created arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. Then the tickerIndex is used to access the data set and loops through the data to get values. More comments were also added to the refactored code for better understanding.

