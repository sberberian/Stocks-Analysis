# All Stocks Analysis

## Overview of Project

### Purpose
After realizing DAQO might not be the best option, this analysis was developed for a finance graduate, Steve, to research various stocks. Knowing the details of these stocks, will help him guide his parents in their investment decisions. 

## Results

### Analysis 
The code written captures the details about the ticker name, total daily volume, and the yearly return for the array of stocks in the dataset. It was written to loop through the rows and recognizes the starting price and ending price for each stock to ultimately be used in calculating the yearly return as seen in the code below respectively.

        If tickerIndex = ticker And tickerIndex - 1 <> ticker Then
    
                tickerStartingPrice = tickerIndex
                
            End If

        If tickerIndex + 1 = ticker And tickerIndex = ticker Then
    
                tickerEndingPrice = tickerIndex
           
            End If

The Tickerindex variable being used also contributes in identifying the stock position for calculating the total daily volume as well that is progressively summed for each stock. 

## Summary

### Advantages and Disadvantages of Refactoring Code
Refactoring code can be benficial in boosting efficiency and reducing a long piece of code to a more concise one. However, it could also make it more difficult to follow for someone who is viewing it for the first time. Sometimes longer code, though less efficient, does provide more context and background for how the results were obtained. This issue could be avoid by adding more commentary as the refactoring process progresses. 

### Advantages and Disadvantages of Refactoring VBA Scripts
Refactoring VBA scripts in particular can be especially helpful by reducing the time necessary to run the script. A possible disadvantage the amount of precision required to ensure that no pieces of essential code or details are lost in the middle of editing. 
