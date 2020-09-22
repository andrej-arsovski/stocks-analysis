# Refactoring VBA code 

## Overview of Project 
This project analyzes and refactors VBA code to optimize performance. The original code analyzed the performance of 12 "Green Tech" Industries in 2017 and 2018, with the goal of determining the best returns on investment at the end of each year. This project makes changes to that original code in order to increase its efficiency as measured by run time.  

## Results

### Stock Performance
2017 was a good year for Green Technology industries. Of the 12 stocks analyzed all but one had positive returns. "DQ" had the highest return rate (199.4%), while "RUN" only returned 5%. "TERP" was the one stack that had negative return at the end of 2017. "SPWR" and "FSLR" were the most highly traded stocks with total daily volumes of 782,187,000 and 684,181,400, respectively. 2018 on the other hand was not as good for the majority of stocks in our analysis. At the end of 2018 only 2 stocks had positive returns, "ENPH" and "RUN" with 81.9 and 84.0% returns, respectively. Taken together the data from 2017 and 2018 suggests that "ENPH" may be a safe investment. It was the 4th highest yielding stock in 2017 and one of two to have a positive return in 2018. Its total daily volumes increased almost 3 fold from 2017 to 2018 (see figures below).

### Original vs. Refactored code

Both versions of the code ask the user to choose which year they want the analysis to run on. The orginal code completes the analysis in 0.71875 for 2017 and 2018.

![AllYear_2017](https://github.com/andrej-arsovski/stocks-analysis/blob/master/Resources/AllStocks_2017.png)

![AllYear_2018](https://github.com/andrej-arsovski/stocks-analysis/blob/master/Resources/AllStocks_2018.png)

The Refactored code makes one major change to the original code. Instead of looping through the entire data set and storing all the values for the different variable, it makes use of a single varibale "tickerIndex" to access the four different arrays used in the analysis. This way the code is using a lot less memory and runs quicker as a result.

tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
   
   For tickerIndex = 0 To 11
    
        tickerVolumes(tickerIndex) = 0

In comparison to the orginal code, the refactored code finishes the analysis in 0.1640625 seconds for 2017, and 0.1328125 seconds for 2018. 

![VBA_Challenge_2017](https://github.com/andrej-arsovski/stocks-analysis/blob/master/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/andrej-arsovski/stocks-analysis/blob/master/Resources/VBA_Challenge_2018.png)

## Summary

Refactoring code can have significant advantages for the developer and the system. It can aid in finding bugs by requiring the developer to break apart the code. Similarly, it can help the developer program faster in the future. It makes the software easier to understand and improves its design. 

There are disadvantages however, it can be risky when the application is big and when the developer does not understand the big picture, or what the software is all about. It can also be costly and time consuming.
