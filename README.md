# VBAChallenge
Utilizing VBA to look at a variety of stocks in order to help Steve show his parents what would be the best stocks for them to invest in.
## Overview of Project
---
The purpose of this project was to use code written in VBA in order to conduct an analysis of various stocks. The analysis of these stocks is to be used by Steve, who intends to share the data found from this analysis with his parents. Steve hopes to assist his parents in determining which stocks they should and shouldn't include in their overall portfolio.
---
## Results
---
Based on the findings from this data, we can clearly see that 2017 was a much better year than 2018, financially, for the stocks included in this data set.
---
![All_Stocks_2017](BerkleyBootcampWork/Resources/All_Stocks_2017.png)
---
![All_Stocks_2018](BerkleyBootcampWork/Resources/All_Stocks_2018.png)
---
As we can see, all stocks listed in the data set, other than TERP, yielded a positive return in the 2017 fiscal year. The strongest performers in this group are DQ, SEDG, ENPH, and FSLR, with returns of 199.45%, 184.47%, 129.52%, and 101.31%, respectively. In 2018, every stock had a negative return aside from ENPH and RUN, which went up 81.92% and 83.95% respectively. The biggest losers of this year were DQ, which had a -62.60% return, JKS, which had a -60.53% return, and SPWR, which has a -44.59% return. What we may be able to gather from this analysis is that ENPH and RUN are both stocks to target as they have has a positive return for both 2017 and 2018.
---
![Ticker_Prices](BerkleyBootcampWork/Resources/Ticker_Prices.png)
---
The code used above was used to determine the starting prices and ending prices of each stock in the respective fiscal year. By determining the starting and ending prices, we were able to calculate the return for each stock in both 2017 and 2018.
---
![VBA_Original_Code](BerkleyBootcampWork/Resources/VBA_Original_Code.png)
---
![VBA_Challenge_2017](BerkleyBootcampWork/Resources/VBA_Challenge_2017.png)
---
The above images show the time taken to run the original code versus the time taken to run the refactored code. As you can see, the refactored code takes incrementally less time than the original code.
---
## Summary
---
After looking at the time taken to run the original code versus the refactored code, we can clearly see that the refactored code takes less time to run. Beyond that, we can point out advantages of the refactored code. Primarily, this refactored code makes it simpler for someone to reach the same output as the initial code. Aside from this, while our code may work fine for just 12 stocks, there may be problems if we are given data with hundreds of stocks. Refactoring our code consolidates this process. In terms of disadvantages, the only real disadvantage is that this refactored code can be more complex and difficult to actually write, meaning there are more potential errors that can come as a result of this refactored code.
