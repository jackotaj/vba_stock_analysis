# VBA Rectored Code and Stock Analysis

## Analysis of Steve's Stocks using VBA and Excel

### Overview of Project
We originally set out to help our client, Steve, analyze a list of stocks that he had provided us with. His parents had wanted to invest their savings in one particular stock but he felt that it would be best to review the performance of multiple stocks first. He wanted to provide them with a 2 year performance review of these stocks so that they would be armed with more information to make an educated decision. 

At first, we delivered a simple excel workbook with a "Run Analysis" button that could analyze the Total Daily Volume as well as the Year-Over-Year Return. The data that he was able to view with our workbook was so eye-opening, he decided he wanted to review an even larger list of stocks. We knew that the VBA code that was used wouldn't be very efficient and that he would have trouble running large datasets so we decided the refactor the existing code. 

After many hours spent and a lot of time invested in additional research, we were able to execute a very fast and efficient refactored version of the original VBA code. Below are the results of both our stock analysis as well as the improved performance of our refactored code.

### Results

#### Stock Analysis

Steve felt that the most effective way to analyze the performance of a stock would be to look at it's Total Daily Volume. A stocks Total Daily Volume (TDV) can be an indicator of liquidity and overal investors overall interest in it. High volume stocks generally points to growth so this was the metric we decided to use for our analysis. 

The stock that Steve's parents were most interested in was DQ. DQ had seen tremendous growth in 2017 with a 199% increase in TDV which led to their overall excitement about it. What had concerned Steve was it's decline in TDV in 2018 where it dropped 62%. He felt that what was once a hot stock a couple years ago may not be as hot and he's advising against investing in DQ. 

In fact, most of the stocks he's reviewed so far, have seen a steady decline in TDV Year-Over-Year with the exception of two, ENPH and RUN. Both have continued to increase in their Total Daily Volume and are showing signs of continuing that trend going into 2019. Steve plans on reviewing more stocks with the updated version of our workbook but does like what he's seeing sof far with ENPH and RUN. 

#### VBA Refactored Code 

##### Below are screenshots and examples of the refactored code:

![alt text](https://github.com/jackotaj/vba_stock_analysis/blob/main/vba_code_tickerindex.png)

![alt text](https://github.com/jackotaj/vba_stock_analysis/blob/main/vba_looparrays.png)

##### Below are screenshots of the execution times for both 2017 and 2018 when running the refactored code. 

##### 2017:
https://github.com/jackotaj/vba_stock_analysis/blob/main/VBA_Challenge_2017.png

##### 2018:
https://github.com/jackotaj/vba_stock_analysis/blob/main/VBA_Challenge_2018.png

The analysis is well described with screenshots and code (4 pt).
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
