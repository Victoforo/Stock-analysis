# Stock-analysis

Overview of Project:
This analysis was carried out with the main objective of being able to measure the profitability of several companies where, for practical purposes, we base ourselves on the stocks, as will be explained throughout the analysis but as a clue the goal is to retrieve the ticker, the total daily volume, and the return on each stock.


Results: 
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
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
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
     End If
            
            
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
     
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     
     End If
            

            '3d Increase the tickerIndex.
     
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
     End If
            
        
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Summary: 

An advantage would be the automation of the process in this case in the analysis, however an important disadvantage would be that it can only be done with a single database since it would be necessary to change all the code completely if one wanted to analyze with other data or in different places the information.
And other advantage when I Refactor the code i can understand that the code is cleaner and more organized.
Some advantages of having a clean code is that is faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward.

Also and advantage of  refactoring was an decrease in macro run time. The original analysis took approximately for 2017 Before: 0.5 seconds afrter the refactoring it takes: 0.085 seconds adn for 2018: Before it takes 0.47 seconds and after: 0.07 seconds (evidence is attached).
![Before 2017](https://user-images.githubusercontent.com/101935525/162636611-6109f718-6427-429b-9b75-3e2885564601.png)
![Before 2018](https://user-images.githubusercontent.com/101935525/162636631-fa4d926e-11e1-4861-8a85-fa97bf1f9293.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/101935525/162636634-81a72ed0-293c-4ef8-86c1-c6d7cc9d6424.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/101935525/162636635-d5cbeee5-0bba-4b78-9d74-f11bb2b6c508.png)
