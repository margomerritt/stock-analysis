#Stock Analysis 


##Overview of Project

###Purpose
	The purpose of this project is to write code to analyze green energy stock projects that can then be reused to analyze any type of stock. The dataset used was an excel file containing the stock data. The data provided had the ticker, date, open, high, low, close, adj close and volume information. This dataset had two years worth of data for the years 2017 and 2018. The same 12 stocks were used in the 2017 and 2018 datasets. 

##Results
###Run Times
####2017 Run Time
	The execution time for the original script for the 2017 stock data ran in 0.4960938 seconds. The refactored script ran in 0.08203125 seconds. Thus the refactored script had a much shorter execution time. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/111299372/195908973-7e2141ca-0252-4b00-a40c-5a386c36f8cd.png)

####2018 Run Time
	The original script for the year 2018 ran in 0.484375 seconds. The refactored script ran at a much faster rate of 0.0859375 seconds. The refactored script shortened the run time by .3984375 seconds. 
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111299372/195908991-a65c6877-8bed-4d73-abb4-e8a7e00f80aa.png)

###2017 and 2018 Stock Performance Comparison
	There are 12 stocks that are analyzed in both the 2017 and 2018 data set. The 2017 stock dataset had 11 stocks with a positive return and one with a negative return. The 2018 stock dataset only had 2 positive returns. The other 10 stocks from the 2018 dataset had negative returns. The positive return percentages in 2017 were significantly higher than the positive returns in 2018. For example, the only two positive returns for the 2018 dataset were ENPH at 81.9% and RUN at 84.0%. The 2017 dataset had multiple returns above the 100% return mark like SEDG at a 184.5% return. 
![All Stocks![All Stocks 2018](https://user-images.githubusercontent.com/111299372/195909048-1c557f9f-c6f5-4d1f-bb42-4946855bac5a.png)
 2017](https://user-images.githubusercontent.com/111299372/195909008-6a005977-e1e6-4195-b6b2-642f3e0f1b8c.png)

###Stock Return Calculation
In column C on the All Stocks Analysis worksheet we calculated the stock return. The general equation to calculate the return is: endingPrice / startingPrice - 1. The two arrays that were declared and used in the code are tickerStartingPrices and tickerEndingPrices. In VBA the line of code looked like: 
![code to calculate return](https://user-images.githubusercontent.com/111299372/195909073-7e2d3b7e-51de-4996-b8d7-1b5debbd35c4.png)



##Summary 
