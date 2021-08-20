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


|    Original Version   | Refactored Version     |
| :-------------:        | :----------: |
| ![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/Original_2017_Time.png)| ![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/VBA_Challenge_2017.png)   |
| ![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/Original_2018_Time.png)   | ![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/VBA_Challenge_2018.png)  |


Regarding the numerical results, as it can be seen in Figure 1 and 2, there is no meaningful difference between the results using the original code and the refactored version. They only differ in the precision fraction.


![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/2017_Original_Results.png)
**Figure 2.**  Yearly Return for 12 different stocks for the 2017 year


![](https://raw.githubusercontent.com/LeidyDoradoM/StockAnalysis_Challenge/main/Resources/2018_Original_Results.png)
**Figure 1.**  Yearly Return for 12 different stocks for the 2018 year

## Summary

* A code refactoring process in general has advantages and disadvantages that need to be evaluate so a decision to refactor or not a given code could be taken.  Following are some advantages and disadvantages of any refactoring process:
 
    **Advantages:**
    1. Refactoring makes the code more efficient in terms of running time which can contribute to save money in terms of physical resources.
    2. The refactoring process can made the code more efficient in terms of the length of lines or complexity; making it easier for follow and understanding.
    3. Another advantage is the saving of time regarding the debugging process for future functionalities that can be added to the original code. If the original code is refactored such that its design is efficient, futher use of this code is easier to understand. 

    **Disadvantages:**
    1. In some cases, the refactoring process could be expensive or requieres more time and effort than the computational time of the code.  
    2. The programer or person in charge of the refactoring has to understand very well the existing code.
    3. For big applications or more complex codes, the refactoring process could be extensive and has to be tested against different cases. So, the refactoring does not add any bugs to the code.

* In specific for this project, the refactoring process makes the VBA script faster than the original version, so this is one **advantage** of the refactored code. As well as, I think the use of arrays for storing the variables needed in the analysis, make the refactored code more understandable and it is easier to follow the script.  At the same time, the use of the arrays implies that the person doing the refactoring needs to understand very well the concept of arrays and how to use them. This can be a **disadvantage** if the person does not have some background in programming, because he or she will requiere more time in understanding what it needs to be done.

