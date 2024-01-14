# VBA Challenge

Overview of Project

The purpose of this analysis is to refactor preexisting code in order to be able to apply it to a wider set of variables. In doing so, we can save time and effort by applying the code for one array of tickers to other sets in information to gain new insights on stock performance.

# Results

In 2017 we see that all but one of the tickers realized gains, including DQ, which was the base for our original analysis. We can see that DQ was the highest performing stock in the set, with a return of approximately 199%. Conversely, TERP lost value, posting a return of -7%. In 2018, however, we see that DQ posted a loss of around 63%. The average 2 year return, combining both values is around 69% or so. Thus we can infer that despite the loss realized in 2018, DQ has an excellent return on investment. However, there is one other stock that posted an even higher return on investment. ENPH recorded a gain of 130% in 2017 and a gain of 82% in 2018, for a total gain of 212% over the two year period. Thus, all factors remaining constant for future growth, we should recommend that Steve's parents focus their nvestments on ENPH.

In the original code we can see that the macro produced the results in the time below:

<img width="256" alt="2017 Original" src="https://user-images.githubusercontent.com/99847786/158745716-f4edfda2-f558-423f-ac74-4293cedad71e.png">

<img width="256" alt="2018 Original" src="https://user-images.githubusercontent.com/99847786/158745704-cce0085b-5818-42b5-a98f-b31d5d5e7e22.png">


In the refactored code we can see that the code was somewhat slower, likely due to the code being larger and looping additional tockers and variables.

<img width="260" alt="2017 Timer Refactored" src="https://user-images.githubusercontent.com/99847786/158745733-403bf2f7-dfc1-4957-a62d-d025d68534ae.png">

<img width="256" alt="2018 Timer Refactored" src="https://user-images.githubusercontent.com/99847786/158745744-704e5eb1-933b-4ed3-acfb-9ad57deee22f.png">


By including the code below, we added several steps where we ran the macro through an array of new tickers, cells, as well as returning to the initial Stocks Analysis worksheet which likely lengthened the time it took to run the code. 

'1a) Create a ticker Index
    
    TickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    'Column 8
    Dim TickerStartingPrices(12) As Single
    'Column 3
    Dim TickerEndingPrices(12) As Single
    'Column 6
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.

    For i = 0 To 11
    'Since the index begins at 0
    
        tickerVolumes(i) = 0
        TickerStartingPrices(i) = 0
        TickerEndingPrices(i) = 0
        
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
    
        tickerVolumes(TickerIndex) = tickerVolumes(TickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> tickers(TickerIndex) And Cells(i, 1).Value = tickers(TickerIndex) Then
            
             TickerStartingPrices(TickerIndex) = Cells(i, 6).Value
             
             End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(TickerIndex) And Cells(i, 1).Value = tickers(TickerIndex) Then
            
            TickerEndingPrices(TickerIndex) = Cells(i, 6).Value
            
            End If

            '3d Increase the tickerIndex.
            
            If Cells(i + 1, 1).Value <> tickers(TickerIndex) And Cells(i, 1).Value = tickers(TickerIndex) Then
            TickerIndex = TickerIndex + 1
            
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = TickerEndingPrices(i) / TickerStartingPrices(i) - 1
        

Summary

One of the biggest advantages of refactoring code is that we can reuse preexisting code to improve efficiency and save time. We can also ensure quality control to a certain degree by using code that we know works and delivers results. On the other hand, one of the disadvantages is that it can be easy to overlook issues in the source code so it's critical to ensure that we understand what's happening in the original code. 

In this challenge, having access to the original code helped immensely because it gave me the opportunity to build off of someone else's work and I could then focus all of my efforts on directing the code to produce the results I was looking for (ex. creating a chart that shows the values for all tickers and not just DQ). However, the disadvantage was that without the notes already available in the code it would have been difficult to know where to start.
