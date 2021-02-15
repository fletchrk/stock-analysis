# Stock Analysis
## Overview of Project
Steve, a friend, asked to have an analysis done on green stocks to provide to his parents whether or not green stocks are worth investing in. The analysis was done in Microsoft Excel using Visual Basic Application (VBA), which provided the stock’s annual return and total daily volume for the years 2017 and 2018. There were 12 green stocks analyzed and provided to Steve in a automated spreadsheet that will show his parents what the best options are when it comes to green stocks.
### Purpose
The purpose of the project is to find efficient ways to pull daily volume and annual return on 12 stocks so they can be analyzed using Excel and VBA. The original VBA code was run BUT was not efficient so the VBA code was refactored. By refactoring the code, the analysis can be performed by Steve by just a push of a button and pulls and formats the data instantly. 
## Results
### Refactored Code Summary
To make code more efficient, it was necessary to refactor the original code. What needed done is the nesting order of for loops had to be rearranged. In order to change around the for loops four arrays were created: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The “tickers” array was used to determine a ticker symbol, while the other three arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices was used to match the “tickers” array using the tickerIndex variable.
The tickerIndex variable assigned a ticker symbol to the tickerVolumes, tickerStartingPrices, and tickerEndingPrices before integrating across the data set. By refactoring the code helps make the analysis run a lot faster than the original code by at least 0.625 seconds faster. 
### Refactored Code
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim Tickers(12) As String
    
    Tickers(0) = "AY"
    Tickers(1) = "CSIQ"
    Tickers(2) = "DQ"
    Tickers(3) = "ENPH"
    Tickers(4) = "FSLR"
    Tickers(5) = "HASI"
    Tickers(6) = "JKS"
    Tickers(7) = "RUN"
    Tickers(8) = "SEDG"
    Tickers(9) = "SPWR"
    Tickers(10) = "TERP"
    Tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index. tickerIndex is used to access the stock ticker index for the arrays (tickers, tickervolumes, tickerstartingprices and tickerendingprices.
   tickerIndex = 0

    '1b) Create three output arrays. Arrays are created for tickerVolumes, tickerStartingPrice, and tickerendingprice. for Ttickervolumes the variable is long while for the startingprice and ending price the variables are single.
    Dim tickerVolumes(12) As Long
    Dim tickerstartingPrices(12) As Single
    Dim tickerendingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. The for loop was created so that it can initialize the tickervolumes to zero. Then an ifthe next row's ticker doesn't match the tickerIndex will increase.
       For i = 0 To 11
       
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.This loop will loop over all the rows of the data spreadsheet (inside the first loop.this nested loop will increase the current tickVolumes variable and adds the ticker volume for current stock ticker.
    For j = 2 To RowCount
     
        '3a) Increase volume for current ticker
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.A If Then statement is used to check that the current row is the first row of the tickerindex. If it is the correct first row it will assign current closing price to both the tickerstartingprices and tickerendingprices variables.
        'If  Then
        If Cells(j - 1, 1).Value <> Tickers(tickerIndex) Then
            
            tickerstartingPrices(tickerIndex) = Cells(j, 6).Value
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(j + 1, 1).Value <> Tickers(tickerIndex) Then
        
            tickerendingPrices(tickerIndex) = Cells(j, 6).Value


            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
            End If
        
        Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = Tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerendingPrices(i) / tickerstartingPrices(i) - 1
        
    Next i
    
    'Formatting code so that the headerrows are bolded, and the returns are green if there is a positive return and red if there is a negative return.
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub
### Original Code Summary
The original code took an average of 0.7578 seconds to run, which is about 0.625 seconds slower than the refactored code. One thing the original code was lacking is the formatting of the cells were not included in the original code. A separate macro was built along with an extra button had to be added so that after the results appear the button can be pushed to format the results. Another issue was that not having a tickerIndex variable in the original code cause several extra for loops to be added.
### Original Code
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

'1) Format the output sheet on All Stocks Analysis worksheet
Worksheets("All Stocks Analysis").Activate
Range("A1").Value = "All Stocks (" + yearValue + ")"
'Create a header row
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"
    
'2) Initialize array of all tickers
Dim Tickers(12) As String
Tickers(0) = "AY"
Tickers(1) = "CSIQ"
Tickers(2) = "DQ"
Tickers(3) = "ENPH"
Tickers(4) = "FSLR"
Tickers(5) = "HASI"
Tickers(6) = "JKS"
Tickers(7) = "RUN"
Tickers(8) = "SEDG"
Tickers(9) = "SPWR"
Tickers(10) = "TERP"
Tickers(11) = "VSLR"
'3a) Initialize variables for starting price and ending price
Dim startingPrice As Single
Dim endingPrice As Single
'3b) Activate data worksheet
Worksheets(yearValue).Activate
'3c) Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through tickers
For i = 0 To 11
    ticker = Tickers(i)
    totalVolume = 0
    '5) loop through rows in the data
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        '5a) Find total volume for current ticker
        If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
        '5b) Find starting price for current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            startingPrice = Cells(j, 6).Value
            
        End If
        
        '5c) Find ending price for current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            endingPrice = Cells(j, 6).Value
            
        End If
    Next j
    '6) Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
### Run Time for each method:
The difference between runtimes for the Refactored Code vs Original Code is about 0.625 seconds.

Refactored Run Times:
![VBA_Challenge_2017](https://github.com/fletchrk/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/fletchrk/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

Original Run Times:
![Original Code 2017](https://github.com/fletchrk/stock-analysis/blob/main/Resources/Original%20Code%202017.png)
![Original Code 2018](https://github.com/fletchrk/stock-analysis/blob/main/Resources/Original%20Code%202018.png)
## Summary
### Advantages/Disadvantages of Refactoring Code in General
There are several advantages and disadvantages of refactoring code. One big advantage is that refactoring code helps make the code more efficient. Another advantage of refactoring code is that logical errors are easier to notice in well-structured code that uses nested for loops. The main disadvantage of refactoring code is that since you are using code that already works you can cause the refactored code unusable if the coding is not correct.
### Advantages/Disadvantages of Refactoring Code in VBA Script
An advantage of refactoring code in VBA script is that usually when you refactor code you are cleaning up the code and making it more orderly and the script runs more efficient. One disadvantage of refactoring code in VBA script is that if you do not know syntax very well you may struggle on refactoring the code. Syntax is something that helps the code run quicker and cleaner.
