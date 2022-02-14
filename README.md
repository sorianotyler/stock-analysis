# stock-analysis

## Overview of Project
 The objective of this project was to create a quick way for our client Steve to Analyze his dataset. We created a VBA program that would allow him to compare and analyze a dozen stocks at the click of a button. He was able to extract the information needed to advise his parents which stocks were performing the best.   
### Purpose
    The purpose of this challenge is to refactor our code to make it more efficient. Although our program works great for smaller datasets, if Steve would like to expand his analysis to a bigger one, the code might take a long time to execute. The goal is to make our code run faster and capable of handling more data.

 
## Results 
### Original Code 2017 & 2018 performance
    The orginal code had a 0.3 second average run time for the 2017 and 2018 datasets. Please refference image "Stock_2017" and "Stock_2018".
### Refactored Code 2017 & 2018 performance
    The orginal code had a 0.09 second average run time for the 2017 and 2018 datasets. Please refference image "VBA_Challenge_2017" and "VBA_Challenge_2018".
    
### Differences in Code
    The Code loops through the rows and collumns to count the total Volume traded, the starting price, and the ending price for each ticker for Steven to analyze. 
    
    #### The orginal code to extract all information is as follows:
        
       For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            '5b) get starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1) <> ticker Then
                
                startingPrice = Cells(j, 6).Value
                
            End If
            
           '5c) get ending price for current ticker
           If Cells(j, 1).Value = ticker And Cells(j + 1, 1) <> ticker Then
                
                endingPrice = Cells(j, 6).Value
            
            End If


       Next j
       '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
        If Cells(4 + i, 3).Value > 0 Then
        
            Cells(4 + i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(4 + i, 3).Value < 0 Then
        
            Cells(4 + i, 3).Interior.Color = vbRed
            
        Else
            
            Cells(4 + i, 3).Interior.Color = xlNone
            
        End If
        
   Next i
   #### End Code
   
   #### Original code findings: 
   As you can see, you have to loop through all of the rows of the data for each instance of i (i being each ticker). That would mean you would loop through over 1000 rows of data 12 times to extract the neccesary information to conduct an analysis. This proved to be inefficient which can be seen in a 0.21 second difference in performace as well as repetitive. Seeing how the data was already organized and grouped together, there was a great opportinity to make it better.
   
    #### The Reffactored code to extract all information is as follows:
    
    For i = 0 To 12
        
        tickerVolumes(i) = 0
    
    Next i
    
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
       If Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
       End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    #### End code
   
    #### Refactored code findings: 
    As you can see in the refactored code, we loop through everything once and created arrays to hold all of the neccesary data. Because the data was organized with every ticker grouped together, we did not need to loop through the entire sheet more than once. As you can see with the screenshots we achived the same exact results at faster speeds. We did not change the function of the code nor did we change the logic. 
    


## Summary 

### Pros and Cons of Refactoring in General 
    Refactoring can be a positive when coding to create more efficient programs. It can also be a pro to create "leaner" or more readable code for you to build upon. I believe the main advantages to refactoring is organization and function.
    Refactoring can also be a negative. Not all refactoring leads to organized code. If you try and make a subroutine too complicated can easily loose functionality and confuse yourself and/or teammates. Refactoring also takes time away from working on other projects. I believe the main disadvantages to refactoring is time and introducing new logic.


### Pros and Cons of Refactoring in this project 
    Refactorting again is a pro in this case because it made our program run faster. If we were to analyze an extremely large dataset on a time crunch, an effience program is crucial. In this specific instance, Steve now has a program that could handle more tickers effectively. 
    Refactoring has it's disadvanatges as well. In this instance, refactoring lengthened the code and made the if then logic more complicated. It was also easy to get confused when making changes and working with new/improved logic.
