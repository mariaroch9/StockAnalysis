# Overview of Project: 
## Background: 
Steve wants us wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, he fears it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

## Purpose: 

Refactor the code so that it can accommodate thousands of stocks. The new code should not only be capable of analysing thousands of stocks but also give the best investment recommendations to his parents in a few seconds or less. 

# Results:
## Performance of the stocks
When we compare the stocks for the years 2017 and 2018, we can see a stark difference in returns. While 2017 is green all the way except for one red; 2018 is a total contrast with red all the way except for two tickers. 
### The Year 2017
All stocks except one gave positive returns indicated by green for the year 2017. In fact, four of the 12 stocks gave over two times the returns. DQ was the best performing ticker giving 199% returns, closely followed by SEDG giving a return of 184.5%

<img width="619" alt="image" src="https://user-images.githubusercontent.com/111670866/188327280-96f6b8b3-fdea-4e30-b16c-1adb0a1cb026.png">


### The Year 2018
This wasn’t a great year for stocks, as highlighted in red the chart above. All except RUN and ENPH had negative returns. DQ was the worst performing stock in 2018 with negative returns of 62.6%
### Both years
Looking at both the years we can conclude that TERP was the most stable in both years. ENPH performed well and gave positive returns of 129% and 82% in the years  2017 and 2018. Looking at the positive returns given by ENPH Steve’s parents may definitely want to place their bets on that one!


## Original Code: 

This code includes 2 arrays, Starting Price and Ending price. 

```


Sub AllStocksAnalysis()

Dim startTime As Single
Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (yearValue)"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(11) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   
   '3a) Initialize variables for starting price and ending price
   
   Dim startingPrice As Single
   Dim endingPrice As Single
   
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       TotalVolume = 0
       
   '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           
   '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               TotalVolume = TotalVolume + Cells(j, 8).Value

           End If
    '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

    '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
           
       Next j
       
    '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = TotalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
```

## Refactored code: 
I made use of 4 Arrays, ticker volume, ticker index, starting price and ending price. All the 4 arrays are looped through to output the Ticker, Total Daily Volume, and Return. Finally, I tried to include a description for each line of code. All these changes have made the code more logical and the execution time is reduced. 

```
 '3a) Increase volume for current ticker
        
      If Cells(i, 1).Value = tickers(tickerindex) Then
      
      TotalVolume(tickerindex) = TotalVolume(tickerindex) + Cells(i, 8).Value
      
     End If
      
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
     If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1) <> tickers(tickerindex) Then
    
    startingPrice(tickerindex) = Cells(i, 6).Value
    
    End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1) <> tickers(tickerindex) Then
    
       endingPrice(tickerindex) = Cells(i, 6).Value
    
       tickerindex = tickerindex + 1
       
       End If

       '3d Increase the tickerIndex.
  
   If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1) <> tickers(tickerindex) Then
   tickerindex = tickerindex + 1
   
   End If
     
        'End If
    
Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
Worksheets("All Stocks Analysis").Activate
If startingPrice(i) <> 0 Then
    
 Cells(4 + i, 1).Value = tickers(i)

Cells(4 + i, 2).Value = TotalVolume(i)

Cells(4 + i, 3).Value = (endingPrice(i) / startingPrice(i)) - 1

Else

Cells(4 + i, 3).Value = 0
 
End If
        
    Next i
    
 ```   

## Execution times

Comparing the execution times of the original script and the refactored script

### For the year 2017. The original code took 0.32 seconds to run while the. Refactored code ran in 0.08 seconds. 
   
<img width="523" alt="image" src="https://user-images.githubusercontent.com/111670866/188327308-94e5725e-ae35-4bdd-a40a-96aa1a1508ab.png">
<img width="473" alt="image" src="https://user-images.githubusercontent.com/111670866/188327317-26b2c172-3521-421b-b841-b40d45bf768a.png">

### For the year 2018. Similar to the previous year 2017, the original code took 0.32 seconds to run while the Refactored code ran in 0.08 seconds. 

<img width="477" alt="image" src="https://user-images.githubusercontent.com/111670866/188327323-7984c234-56a2-443a-8a2f-bc8b61845062.png">
<img width="477" alt="image" src="https://user-images.githubusercontent.com/111670866/188327327-126a16b7-9421-4e82-a3d4-8e79eb552682.png">

 
# Summary: 
## Some of the advantages of refactoring code
1.	##### Refactoring Improves the Design of Software by making it more organized and cleaner
2.	##### Refactoring Makes Software Easier to Understand since it becomes more logical
3.	##### Refactoring Helps Finding Bugs
4.	##### Refactoring Helps make the program run faster as can be seen in the above pop-up examples between the original code and the refactored code.  

## Some disadvantages of refactoring code
1.	##### It is more time-consuming and therefore costs more
2.	##### It could sometimes introduce new bugs to the code which might require more time for debugging

## How do these pros and cons apply to refactoring the original VBA script?
When we look at the final output of the refactored it is undoubtedly more efficient. What took 0.3 seconds can now be run in 0.08 seconds. However, one cannot overlook the high risk that refactoring bears. What if in the bargain some new bugs were introduced? What if some important code got deleted in the process of refactoring? It’s better to save a copy of the original code which can be used as a reference just in case of these uncertainties. 
Refactoring also needed additional time investment. If we have a time-bound commitment then the additional time to refactor a code definitely must be accounted for. When I was refactoring, I ran into a couple of errors. Debugging these errors also took me some additional time. 
In conclusion, the benefits of refactoring including making a more logical code, that is clear, more organized and runs efficiently should be weighed against the additional time needed to refactor the code. 



![image](https://user-images.githubusercontent.com/111670866/188327242-23d1cbb7-fc9d-46e2-91f3-96781fcd8277.png)
