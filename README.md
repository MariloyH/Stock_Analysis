# stock-analysis VBA_Challenge
##  Overview
 The Client, Steve,  is advicing his parents in stocks investsments. We built for him some MACROs in VBA and excel which analize all stocks from a explicit year, and displays its  relevant data. This is working well, but our Client is worried about   the perfomance of that procedure if we analize thousand of stocks. We will refactor previous code to improve its performance.
  
## Results
The original code had an elapsed time of **0.5273438** seconds for the 2017 data and of **0.5373438** seconds for 2018, both with the same number or registers. So, the difference is not relevant.  From this procedure, *Sub AllStockAnalysis()* and  we were asked improve the execuction time through using arrays for store data instead only variables.  We used this arrays to access and display final informations more quickly. The original code has 2 buttons to call this procedure and another for clear the worsheet *Sub ClearWorksheets*. I inserted a third button, to facilitate both analysis, the original and the refatored.
The resulting time dropped until **.328125**, 37.7% faster than original version. 



## Summary 
### Refactoring code in general 
When we start from an existing code, so  we must to be sure to understand what is the problem that the code resolves. From that, we should understand what does the whole code must do and what does each line of code do, because is very easy to do something wrong and spoil the things! We must keep in mind in what aspect we want to improve the code, for example if we are going to correct an error, or if we want to make it faster or to add new functionality, and in that case, be always sure that the past performance not was afected. 

#### Advantages
   -We begin from an existing code, so we will have a solution in less time (assuming previous code does what it supose to do). 
   -We work only in some parts of the code, not in the whole code.   
#### Disadvantages 
   -If previous code is not properly commented, we could spend a lot of time guessing what the previuos code do.
   -If we are not careful, we can make the code more complex and affect the performance of the execution.
   -If we are adding new functions, we must be careful not to affect behavior that was previously working well. 
   
### Refactoring this VBA scrip
   As advantage, the previous code was written for me, so I knew what I did! It minimized the error posssibility, but no disappeaar!  This code was runnig and working well. The new code introduced a new way to store the values in an array  that improve the performance of the code. I don´t see any disavantages, because the resukt was succesful. Maybe } There are 
  
 

