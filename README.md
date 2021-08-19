# Module 2: Stock Analysis with VBA

Steve has just got a degree in fincance and wants to use his knowledge to help out her parents to diversify their investments in the green energy market. He had created an Excel file with information about some stocks that he wants to analyze. For his analysis, he is relaying in our expertise using Visal Basic Applications (VBA) so the analysis process can be authomated.

## Overview of Project

Steve has already a VBA macro that allows him to have at the distance of a click, the *Yearly Return* of a handful of stocks for years 2017 and 2018.  However, he wants now to expand this analysis to the entire stock market. Although this code works well for a docens of stocks, it might not work as well for thousands of stocks. Therefore, we will need to edit the code such it will be more efficient,i.e. run faster.

### Purpose

The purpose of this project is to refactor the solution we code in Module to loop through all the data one time, such that the code successfully made the VBA script run faster.

### Analysis and Challenges

In order to get our initial VBA script to run in an efficient way, we need to create three variables as arrays to store three calculations needed to compute the *Yearly Return* for each one of the stocks.  These are: `tickerVolumes`,`tickerEndingPrice` and `tickerStartingPrice`, whose sizes are given by the number of stock markets that are considered, which in this case is 12.  Besides, we need a variable, `tickerIndex`, to get through the stocks that are stored in the array `tickers`.

The use of arrays to keep track of different variables, calculated for each data point (i.e. each row in the dataset), allow us to go through the data just one time and avoid the use of *Nested loops* that usually increases the time of execution of any code.  
The main part of the implementation of this week's module involved a nested loop with two *for loops*, the first for loop goes through the array that contains the tickers and the other one goes through the data rows.  This implies that the data set has to be traversed 11 times.  In contrast, the refactored version, does not have nested for loops, which implies the data is traversed just one time.  The follow lines show the important part of the code, emphasizing only the lines that differ between the two approaches:

- First approach code with two nested *For* loops: 
```vb
Dim StartingPrice As Double  
Dim EndingPrice As Double  'Variables declared as double values
 
For i = 0 To 11    'Loop through all stocks
    ticker = tickers(i)
    totalVolume = 0
   
    For j = RowStart To RowCount   'Loop through rows in data
    '5a) Find total volume for the current ticker
        If Cells(j, 1) = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
    '5b) Find starting price for the current ticker
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            StartingPrice = Cells(j, 6).Value
        End If
    '5c) Find ending price for the current ticker
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            EndingPrice = Cells(j, 6).Value
        End If
            
    Next j
        
    '6) Output the data forthe current ticker
        
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (EndingPrice / StartingPrice) - 1
Next i
```
 - Refactored approach with only one *For* loop and using arrays:
```vb
Dim tickerIndex As Integer
tickerIndex = 0
Dim tickerVolumes(12) As Long  'Create three output arrays
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single  

For i = 0 To 11  'Create a for loop to initialize the tickerVolumes to zero.
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i    
    
For i = StartRow To RowCount  'Loop over all the rows.
    
'3a) Increase volume for current ticker=         
    If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    End If
             
'3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

'3c) check if the current row is the last row with the selected ticker
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    End If
            
'3d Increase the tickerIndex. If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
    End If
Next i
```
## Conclusions and Results

As has been established, this refactoring does not add new functionality to the initial code, but it improves its performance by reducing the running time. Therefore, the numerical results for both approaches are still the same, although, the running time differs.  Table 1. shows the time of computation for both years 2017 and 2018 and for both approaches.

| Original Version      | Refactored Version     |
| :------------- | :----------: |
| | ![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/VBA_Challenge_2017.png)   |
| You Can Also   | ![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/VBA_Challenge_2018.png)  |

Regarding the numerical results, as it can be seen in Figure 1 and 2, there is no difference between the results using the original code and the refactored version.

## Summary

* What are the advantages or disadvantages of refactoring code?
Advantages: -More efficient in the running time
             -Less lines of coding
Disadvantages: - The logic of coding implies the use of arrays, which can increase the complexity of coding process. Therefore is more dense and difficult to follow.

* How do these pros and cons apply to refactoring the original VBS script?

