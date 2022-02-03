# stock-analysis

Project overview
---------------------
Purpose
---------------------
The purpose of this analysis was to refactor VBA existing code to receive information of stock performance of some companies in the year 2017 and 2018 and determine whether stocks are worth investing. The main goal was to make a new refactored code run faster.

Background data
-------------------------
The data used for this project included stock information of 12 companies. The data braked down into several columns, as ticker, date, open price, highest and lowest price, adjusted price and stock volume. To determine the performance of the stock the ticker was retrieved, total daily volume and return were calculated.

Results
----------------------------
Before start refactoring, the original code runtime was tested. Below you can see the time for both 2017 and 2018 years.



The part of original code was copied to make an input box, to create a header row, to initialize the array of all tickers, to activate the worksheet needed and to set up the timer to determine code runtime. Also, original data formatting was saved. The rest of code was refactored. You can find it below with comments for each step:


'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
     Next i
     
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
           
            
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
            
                tickerIndex = tickerIndex + 1
            
            End If
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
              
        Cells(4 + i, 1).Value = tickers(1)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        
        
    Next i
    
Here are the results of code runtime after refactoring :

2017 after refactoring 2018 after refactoring

As we can see, the runtime decreased drastically.
                       

Summary
--------------------
Advantages and disadvantages of refactoring code
Although refactoring does not add features or functionalities, it definitely improves the design of software, helps debugging, makes code easier to understand and helps programming faster. The main disadvantages are that it takes a lot of time and may increase the chance of making a mistake that will make you waste more time solving it.

The pros and cons of the original and refactored VBA script
The biggest benefit of refactored code is sizeable decrease in code runtime. The refactored code now runs 6 times quicker for both analysis of the year 2017 and 2018 (0,53sec with stock code, 0,08sec with refactored code). There were no drawbacks found during making the project.
