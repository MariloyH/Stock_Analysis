# stock-analysis VBA_Challenge
##  Overview
 The Client, Steve,  is advicing his parents in stocks investsments. We built for him some MACROs in VBA and excel which analize all stocks from a explicit year, and displays its  relevant data. This is working well, but our Client is worried about   the perfomance of that procedure if we analize thousand of stocks. We will refactor previous code to improve its performance.
  
## Results
I started from the procedure, *Sub AllStockAnalysis()*. We were asked improve the execuction time through using arrays for store data instead only variables. An array is data type which can storage multiple data and its accesed due an index. In this case, I developed a new procedure, *AllStockAnalysisRefactored()*, in which  I used this arrays to storage tickers names, total volume of transactions per year and the starting and ending price of ticker. At the end of the process, the resulsts are displayed. The original code has two buttons on the worksheet: one  to call the procedure and other for clear the worsheet *Sub ClearWorksheets()*. I inserted a third button, to facilitate both analysis, the original and the refactored.
Once we compared both routines, I can observe that refactored procedure is faster than previous. In the case of the 2017 data, the resulting time  dropped from **0.5273438 seg** to  **.328125 seg**, 37.7% faster than original version. In the case of the 2018, the resulting time dropped from **0.53125 seg**  to  **0.3320132 seg**, %37.5 faster
1. 
![](https://github.com/MariloyH/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png) 
2. 
![](https://github.com/MariloyH/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png)

 The use os arrays makes de diferencie because we are storing all the values  in a single data type  meanwhie we read them from the Year tab and  show the results on the
 "AlStockAnalyisis" tab. I think thats the signiuficant difference betwen the two codes.
 ![](https://github.com/MariloyH/Stock_Analysis/blob/main/Refactored_Initializing_Arrays.png)
 ![](https://github.com/MariloyH/Stock_Analysis/blob/main/Refactored_Storing_Values.png)
 [](https://github.com/MariloyH/Stock_Analysis/blob/main/Refactored_Display_Results.png) 
 
## Summary 
### Refactoring code in general 
When we start from an existing code, so  we must to be sure to understand what is the problem that the code resolves. From that, we should understand what does the whole code must do and what does each line of code do, because is very easy to do something wrong and spoil the things! We must keep in mind in what aspect we want to improve the code, for example if we are going to correct an error, or if we want to make it faster or to add new functionality, and in that case, be always sure that the past performance not was afected. 

#### Advantages
   1. We begin from an existing code, so we will have a solution in less time (assuming previous code does what it supose to do). 
   2. We could find hidden bugs, or find code does nothing. 
   3. We can make the thigs easier, solving the problem from an alternative point of view.  
   4. We improve ourselves if we look our own code to refactor. 
#### Disadvantages 
   1. If previous code is not properly commented, we could spend a lot of time guessing what the previuos code do.
   2. If we are not careful, we can make the code more complex and affect the performance of the execution.
   3. If we are adding new functions, we must be careful not to affect behavior that was previously working well. 
   
### Refactoring this VBA scrip
   As advantage, the previous code was written for me, so it minimized the error posssibility, but no disappeaar!  This code was runnig and working well. The new code introduced a new way to store the values in an array  that improve the performance of the code. At the same time I added some fuctionality (Addin an aditional button) to ease the execution of the refactored code. Maybe I don´t see any disavantages, because the result was succesful. 
  
 

