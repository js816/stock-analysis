# VBA of Wall Street

## Overview of Project

### Purpose

Steve, a finance graduate, wanted to help his parents gain insight into historic performance of several green energy stocks.  His parents, who have invested in DAQO New Energy Corp (DQ), wanted to see how the stock has performed compared to several other green energy stocks.  Steve requested my help with automating this analysis.  Neither Steve nor his parents are experts in Excel and ensuring the file was easy to use was a high priority for this project. 

## Results

### Analysis of stock performance between 2017 and 2018

During 2017, overall, the stocks had a very positive year.  All but one (TERP, -7.2% return) had positive returns.  DQ, the company Steve parents decided to invest in, saw its stock price nearly triple in 2017 (199.4%).  Three other stocks saw their stocks more than double (SEDG 184.5%, ENPH 129.5%, and FSLR 101.3%).  Other very positive returns included JKS (53.9%) and VSLR (50.%).  Overall, the total daily volume was 264,000,000 trades.

![2017_Stock_Performance](https://user-images.githubusercontent.com/82730954/117589280-37372d00-b0ee-11eb-9f83-b26d332e912d.png)

The following year, most of the stocks had much worse performance.  Only two stocks showed positive returns, ENPH (81.9% and RUN 84.0%).  Many of the stocks showed double digit negative returns with DQ having the worst performance (-62.6%).  Overall, the total daily volume trended upward slightly to 276,000,000 (an increase of about 4%).

 ![2018_Stock_Performance](https://user-images.githubusercontent.com/82730954/117589290-41592b80-b0ee-11eb-99f6-11c71623de42.png)

Only two years of data was available for analysis.  Based upon the analysis of this data, it is suggested that Steve and his parents closely consider the various options before deciding where to invest.  Although the future cannot be predicted, it appears possible that ENPH, which had strong positive returns in both years, may have more stability to be able to weather storms that some of the other companies lack.

### Analysis of execution times of code

According to Wikipedia, refactoring code is the process of modifying existing code without changing the end product.  In our case, code was refactored to improve the efficiency of the script thus reducing run time.  Refactoring the code led to dramatic improvements in the processing time.

![Script_Performance_Table](https://user-images.githubusercontent.com/82730954/117589303-4ddd8400-b0ee-11eb-9586-98e3c6c9debc.png)

![Script_Performance_Chart](https://user-images.githubusercontent.com/82730954/117589310-5635bf00-b0ee-11eb-89e9-ff06a09ded63.png)


Both the 2017 and 2018 worksheets contain 3012 rows of data, 251 for each stock.  

The original scripting, while providing the correct summary data, was not as efficient as it could have been.  It looped through all 3012 rows twelve different times (processing a total of 36,144 rows of data).  To get the summary data for the first stock, the script looped through all rows and then wrote the values to the summary table.  It then did a second loop through all the rows before writing the values for the second stock to the summary table, and so on as it processed all twelve stocks.


```
   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
        'set initial volume to zero
        totalVolume = 0
       
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
             If Cells(j, 1).Value = ticker Then
        
                'increase totalVolume by value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
    
            End If
           
           '5b) get starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
    
                'Setting startingPrice to value in current row
                startingPrice = Cells(j, 6).Value
        
            End If
           
           '5c) get ending price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
    
                'Setting ending price to value in current row
                endingPrice = Cells(j, 6).Value

            End If

       Next j
       
       '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
```


The refactored code takes advantage of its ability to hold many things in memory while on its journey.  These improvements of the second script were realized by using arrays to capture and hold the values (total daily volume, stock starting price, and stock ending price) for each stock as it looped through the data only once.  As the script worked through the rows, it noted when the data changed from the first stock to the second stock and it then stored the value for the first stock in an array and started calculating the data for the second stock.  At the end of its only pass through the data, the values from each of the arrays were added to the summary table.

```

        '1a) Create a ticker Index to be used with arrays for holding summary data until added to summary worksheet
    Dim tickerIndex As Integer
    
    'Initialize tickerIndex to 0
    tickerIndex = 0
    

    '1b) Create three output arrays for holding summary data until added to summary worksheet
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'JS Challenge - this doesn't seem to make sense.  I commented it out and code ran correctly and gave accurate values.  Talked through with TA Sasha.
    'For i = 0 To 11
    
    'tickerIndex = i
        
        'tickerVolumes(tickerIndex) = 0
        
    'Next i
        
    ''2b) Loop over all the rows in the spreadsheet, just one time
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            
            'If so, set starting price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            
            'Capture ticker ending price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i

```


### Challenges and Difficulties Encountered

During this challenge, initially I had difficulty wrapping my mind around using the arrays to hold the values.  Talking with TA Sasha in office hours, she suggested that I try to see the arrays as Excel tables even though they’re only held in memory until the final step of the scripting.  Her suggestion helped me to understand the concept more clearly and I was able to work through the coding with much less effort than I’d been thinking.  

Upon debugging a couple of errors in the code (including typos), I did, however, find that the code for step 2a didn’t seem to make sense.  I commented out those lines to be able to run the remainder of the code.  I anticipated that the outcome would be correct for the first stock but might increment for all stocks beyond that rather than resetting at 0 as each stock changed.  However, my thinking about this didn’t take into account the arrays I’d created.  And the output gave accurate data thus it appears that step 2a is not needed.

As I ran through the refactored code the first time, I did not capture the screenshot of the 2017 runtime because I anticipated the data was incorrect.  When I saw the data was correct and reran the code to capture the runtime, it gave a result that was so small it had to be shown in scientific notation.  So I closed out the file and reopened and captured the best runtime I was able to capture.  I talked through both of the hiccups above with Sasha during office hours.

## Summary

Refactoring code is an opportunity to step back and look at the goal of the code and ask some questions.  _What are we trying do here?  What’s our goal?  What other ways could we make this work?_

Taking a step back gives you an opportunity to see things from a different perspective and perhaps improve upon what already exists. 

Potential disadvantages to refactoring code include introducing bugs, potentially impacting other parts of code, or investing time that may not result in any meaningful improvement in processing time.
 
While the original code achieved the goal and provided accurate information, an opportunity existed.  The original code was inefficient as it looped through all 3012 rows twelve times.  Since the data provided was sorted first by stock, then by date, the performance could be optimized by iterating through the index value as the stocks changed.  For this particular use case, a reduction of about 0.5 seconds doesn’t seem huge.  However, if this code were to be used for hundreds or thousands of stocks, the reduction would be significant.  Regardless of the improvement, the time invested in this refactoring proved to be a valuable learning exercise providing a deeper understanding of various concepts of coding and being able to look at more than one solution to determine which is best for the goal.  As the old saying goes, there’s more than one way to skin a cat (or bake a cake).  While one method works, finding a way that works better provides value.
