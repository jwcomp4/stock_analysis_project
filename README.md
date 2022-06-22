# Module 2 Challenge: Refactoring Stocks Analysis Script

## Project Overview

- For this analysis, we are mining the 2017 and 2018 green stocks dataset to find the total daily volume for each stock and calulate the annual return for each corresponding year to help Steve advise his parents investment strategies.

- In this particular challenge, I am refactoring the code to run more efficiently so Steve can use this script for any size dataset to return an accurate and easily read analysis quickly. 

## Results

### Comparing 2017 and 2018 Green Stocks

- Generally, all of the stocks performed better in 2017 than in 2018 with the exception of RUN. 
    - In both 2017 and 2018, RUN returned a profit with 2018 seeing an 84% return and a 5.5% return in 2017. 
    - Additionally, ENPH was the only other green stock to see consistant growth in both 2017 and 2018, returning 129.5% and 81.9% respectively.
    
    PHOTOS OF RESULTS

- I recommend investing in ENPH and RUN as both demonstrate good growth even in years that other green stocks fell.

### Original vs. Refactored Script

- While producing accurate results, the original script included a nested for-loop to find the total volume and annual return for each indixe of the tickers array, resulting in the script iterating through the entire dataset 12 times, taking .31 seconds for each the 2017 and 2018 datasets. Below is the code highlighting the inefficient nested for-loop: 

    
        For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        '5 Loop through the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
        '5a Find total volume for current ticker
        
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
        '5b find startingPrice for current ticker
        
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
        
                startingPrice = Cells(j, 6).Value
            
            'set starting price using another conditional if statement
            
        
             End If
        '5c find the endingPrice for current ticker
        
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            
                endingPrice = Cells(j, 6).Value
            
            'set ending price using another conditional if statement
            
             End If
            
        Next j
    
        '6 Output the data for the current ticker doing the same analysis as DQ
        Worksheets("All Stocks Analysis").Activate

        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

        Next i

- The refactored script only loops through the data one time producing the same results in only .08 second for each dataset (74% faster).
    - Below is the code highlighting the the more efficient structure:

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
            
         '3d Increase the tickerIndex.
                 
            
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    

PHOTOS of 2017 and 2018 TIMES 

## Summary 

- Refactoring code largely presents benefits if it results in more efficient processes or more easily read code. The only word of caution is that one should not pursue efficiency at the expense of rigorous analysis.

- In this case, refactoring the code produced the same accurate results in less time making the script more usable for much larger datasets. 

