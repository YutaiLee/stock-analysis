# Stock-Analysis With Excel VBA
## Overview of the project
### Purpose
The purpose of this project is to build an Excel VBA code for judging stock information in 2017 and 2018, and to analyze whether these stocks are worth investing in. In addition, through VBA code, to speed up the analysis time and reduce the file size.
## Results
### Analysis
First, I first copy the code provided by Module 2 to the new Module. Through the instructions of refactoring the code, start to set the value of each ticker and activate the worksheet. Follow each introduction to write for loop to complete the operation of the entire code. Below is the instruction and code as written in the file.

    '1a) Create a ticker Index
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = o To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6)
        End If
        
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6)
        End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1) = tickers(i)
        Cells(4 + i, 2) = tickerVolumes(i)
        Cells(4 + i, 3) = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
## Summary 
### Pros and Cons of Refactoring code
Refactoring helps to make our code clearer and easier to understand. Clear code can make design and software easy to improve, and debug and speed up programming. In addition, refactoring can also help the code run faster and save the time required for work. However, we do not always have the opportunity to refactor our code. Although refactoring code can make it easier for us to read and add new code, refactoring code has certain risks. If we input wrong instructions when refactoring the code, or change the original instructions, it may cause more serious errors.
### Pros and Cons of original and refactored VBA script.
![image](https://github.com/YutaiLee/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)
