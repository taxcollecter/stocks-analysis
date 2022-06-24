#  VBA of Wall Street 

## Overview of Project
In an effort to aid Steve in analyzing the performance of several Green Energy stocks for his parents, we plan to leverage VBA against stock data to parse and expose  market results of the industry occurring during the years of 2017 and 2018.

## Results
In assisting Steve in analyzing the performance of his Green Energy stocks, we found that all but 1 stock failed to improve on their outstanding Return performance noted in 2017 (See 2017 Performance below). 
![2017 Performance](https://github.com/taxcollecter/stocks-analysis/blob/78565f5e4333904abdf9d26ff97c11aa3ba30c66/Resources/2017_Stock.png)

In fact, 2018 appears to have been a negative year for the entire sector (See 2018 performance below).
![2018 Performance](https://github.com/taxcollecter/stocks-analysis/blob/78565f5e4333904abdf9d26ff97c11aa3ba30c66/Resources/2018_Stock.png)

Without additional context, we'd advise Steve and his Parents to stay away from investing in the Green Energy sector until more information to explain the drop in returns is readily available. 

Disclaimer: Analysis was done via VBA, leveraging stock data provided by 3rd party (Code snippet below and execution proof below.)
![Code Snippet](https://github.com/taxcollecter/stocks-analysis/blob/010b23031e4867102fed8662d96d2407c40a8114/Resources/Code_snip.png)
![Execution Timing](https://github.com/taxcollecter/stocks-analysis/blob/010b23031e4867102fed8662d96d2407c40a8114/Resources/VBA_Challenge_2018.png)

## Summary
In review, I can see advantages AND disadvantages of refactoring code. The speed to turning around a requirement is directly aided by not having to re-write every line of code in a frequent request. However, refactoring can also present issues with development as well. For example, bloated code has the potential to persist, AND GROW, if developers do not refactor with specific request in mind. This requires the developer to remove sections of code that do not pertain to the current code's task. If done properly, refactoring code should still require a decent investment in time.  

As it relates to this exercise, I found it a bit easier to not refactor the code previously provided within the VBA modules. When trying to nail down the logic driving the loop and nested loop, writing and reviewing the output for myself cleared up a bit of confusion on my end. 