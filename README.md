# Stock-Analysis

##Overview Of Project

Steve wants to do more research for his parents and find the best stock option to invest their retirement funds. In the initial analysis for one stock 'DQ' we determined that the company had a -63% return in 2018. This does not appear to be a safe option to invest. As a result, I expanded the dataset to include the entire stock market over the last few years. The challenge in this project is to refactor the initial code to now run for the entire stock market. 

###Results

 From my analysis, I determined that in 2017 all stocks received a 5% or higher return, however, in 2018 all stocks fell into a negative return of -3.5% or more. There are only two stocks that remained in a positive return for both 2017 and 2018, ENPH and RUN. However, 'ENPH' saw a decrease in return by 48% from 2017 to 2018 while 'RUN' saw an increase in return by 79%. 

![image](https://user-images.githubusercontent.com/103547108/166394249-023fb030-37df-4c9f-8f61-6a31e34b00ce.png)![image](https://user-images.githubusercontent.com/103547108/166394294-1b73a08f-00a4-49ae-89f7-de3e6338a630.png)

To refactor the code, I was challenged with consolidating the loops into one so we only run against the dataset once. In theory, this should cut  down on the runtime of the VBA script. Currently, the script takes 0.3125 seconds to run for 2017 and 0.3125 seconds to run for 2018.

 Pre Code-Refactor:
 
 ![image](https://user-images.githubusercontent.com/103547108/166394661-be5f6149-21a3-4c1a-a940-f49215c76239.png)


 Post Code-Refactor:
 
 ![image](https://user-images.githubusercontent.com/103547108/166394889-246d7742-32b1-4973-b28e-2200d2fdab32.png)

After running the refactored code, I confirmed my theory that the execution time would decrease. 

Pre-Code Refactor:

![2017 run time pre refactor](https://user-images.githubusercontent.com/103547108/166395486-1b06836c-0597-4fc7-813c-606ef6d0fb22.png)![2018 run time pre-update](https://user-images.githubusercontent.com/103547108/166395496-e56c4f5c-c860-43c6-a2ab-6373140fa0b7.png)

Post-Code Refactor:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/103547108/166395432-ca705e6b-bda0-431e-b0d5-34beb738d903.png)![VBA_Challenge_2018](https://user-images.githubusercontent.com/103547108/166395452-4cdb2322-03fe-4c88-8103-b62a1e9ce95d.png)

###Summary

The advantage to refactoring this code was a faster run time and shorter, more efficient code. However, refactoring code could be very confusing with unnecessary nested loops. Consolidate loops where possible to avoid convoluted code and help with future debugging. 

According to the data provided in the analysis, I believe Steve should recommend his parents invest in RUN and ENPH stock for best potential of return on their investment.
