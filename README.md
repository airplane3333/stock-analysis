# Using VBA for Stock Analysis

## Overview of Project

 ---
The purpose of the project is to complete an alalysis of a set of environmentally 
friendly 'green' stocks using using daily volumne and closing price for a give year. 
Using VBA to read through the stock data then output the data for total daily volume
and the annual return on the scock while creating a report that is formatted and easy to read.
Finally once the report is complete, the code was refactored to include a new index and then 
measure results. 
### Results of Report
---
The analysis showed that many of the 'green'stocks had positive returns in 2017 however, when 
compared to 2018, only 2 stocks had positive returns as compared to the start of the year.  I'd consider 
additional stock data for 2019 or 2020 before making a decision on investing. 

### Performance 2017
![2017 Green Stock Performance](/resources/Results_Stock_2017.PNG) 
### Performance 2018
![2018 Green Stock Performance](/resources/Results_Stock_2018.PNG) 
 
## Refactoring Results
---
Refactoring the original VBA code was a challenge for me.  A timer function was used to capture the start 
and end of of the code execution. 

![Use of Timmer](/resources/Timer_fun.PNG) 
To improve the codes performance an index was added so as the script ran through the data and stored the 
information.  the prvious VBA code would output the data after each stock was analilized. 

![Initializing tickerIndex](/resources/int_tickerIndex.PNG) 
 
After the refactoing of the code, the time to execute the code was improved significantly approximetaly 10 times faster.

![Results 2017 Performance Increase](/resources/Results_Stock_2017_ref.PNG)
![Results 2018 Performance Increase](/resources/Results_Stock_2018_ref.PNG)
  
##Summary
 ---
After refactoring the VBA code, the performance increase was significant.  Refactoring code is an industry best 
prcatice, often the initial codes meets the requirements, but not until the code is complete can it then be reviewed 
to optimize resources and effecenciey.

In this example, I was able to refactory the code and learn something new.  However, it can take significant time 
to complete and may not have any measureable improved results.  Also, steps need to be taken to protect the 
initial code, so using GitHub or other measuraed are critical before starting on any refactor project.

 
 
