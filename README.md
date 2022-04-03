# stock-analysis VBA_Challenge
##  Overview
 Our Client, Steve,  is advicing his parents in stocks investsments. We built for him a some MACROs in VBA and excel which analize all stocks from a explicit year, and displays is relevant data. From this, our Client is worried about   the perfomance of that procedure if we analize thousand of stocks. We will refactor previous code to improve its performance.
  
## Results
  We begin from our previous procedure, *Sub AllStockAnalysis()* and we choose to improve the time through using arrays for store data instead only variables.  We could use this arrays to displat final informations more quickly. The original code has 2 buttons to call this procedure and another for clear the worsheet *Sub ClearWorksheets*. I inserted a third button, to facilitate runnig bothan alysis, the original and the recfatored.  
The el


## Summary 
### Refactoring code in general 
When we start from an existing code, we must to be sure to understand what is the problem that the code resolves anb from that what does the whole code must do and what does each line of code do, because is very easy to do something wrong and spoil the things!  
#### Advantages
  -We begin from an existing code, so we will have a solution in less time (assuming previous code does what it supose to do). 
  -We work only in some parts of the code, not in the whole code.   
#### Disadvantages 
   -If previous code is not properly commented, we could spend a lot of time guessing what the previuos coder did.
   -If we are not careful, we can make the code more complex and affect the performance of the execution.
   -If we are adding new functions, we must be careful not to affect behavior that was previously working well. 
   
### Refactoring this VBA scrip
   As advantage, the previous code was written for me, so I knew what I did, so It minimized the error posssibility. This code was runnig and working well. The new code introduced a new way to store the values in an array  that improve the performance of the code. I don´t see any disavantages, because the resukt was succesful. 
  
 

