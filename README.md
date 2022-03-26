# Stock-analysis

##**Overview**
Stock data was pulled for 2017 and 2018 and a macro was created to pull information for each stocks total volume, start and end prices, and returns. The code ran in .87 seconds. For larger projects with mass stocks over many years, this would take a fair amount of time. Thus the code needed to be refactored to be more effecient. Microsoft VBA was used in order to accomplish this task.

-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

##**Results**
To achieve a quicker run time the existing code was editted to make the process of iterating through stocks more streamlined. variables were vreated for the ticker volume, start, and end prices. The tickers for stocks were iterated over, pulling values for each of the variables declared. The start and end prices were then used to calculate the return of each stock and output to the excel sheet. For easy and clear visualization of the returns, formatting was done to color code positive and negitive returns and format the chart headers. 
The original code ran in 0.847 seconds. ![Code_Run_Time_Original_Module_Code](https://user-images.githubusercontent.com/100040705/160252069-e0881806-7616-4cf5-a151-08ea3499686a.png)
Once the code was refactored it ran in 0.765
![Refactored_code_run_time](https://user-images.githubusercontent.com/100040705/160252232-d3bab437-300d-4f05-b77f-395c59756fa1.png)
From this we can see an improvement in time, saving 0.082 seconds per run. Improvements were made to iterating through the stock tickers, as this is where the majority of refactoring was done. The updated code is as follows:
![Refactored_Code](https://user-images.githubusercontent.com/100040705/160252314-0b0205c5-903a-4070-aa03-3a409919679d.png)
![Refactored_Code_2](https://user-images.githubusercontent.com/100040705/160252315-f723c9ac-fb40-452a-84e3-eaef4154db2a.png)
![Refactored_Code_3](https://user-images.githubusercontent.com/100040705/160252317-37d3add3-c5fd-496e-8d07-46e2810683ab.png)

-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

##**Summary**
When refactoring code, the main advantage is having quicker run times as code is made to be more efficient and streamlined. This allows code to be used with larger files as the code can quickly work through great amounts of data. A disadvantage is that it could lead to error or bugs while editing old code that may be avoided by starting fresh. This can be avoided with careful commenting and organization to assist in avoiding errors. 
The original VBA script was slower, and would not perform well with a larger data file of thousands of stocks, but handles well with the small file it was originally created for. The refactored code is faster, but it took a significant amount of time to create.

##**Challenges**
I encountered many challenge throughout the refactoring process. I tried to get help in office hours, tutoring, and through askBCS and continued to get compile errors with my variables, especially when using tickerIndex. So, I ended up commenting the tickerIndex out or removing it later in the code because after hours of working with it I still could not get it to successfully run. My code still ended up being faster than the orignal, so I got to the goal through a different path. 
