# Stock-Analysis

Overview of Project

The purpose of this project is to utilize stock ticker data to assess whether each stock is worth investing in. I used Microsoft VBA code to automate the analysis of the ticker data, prompt user input, add conditional fomatting, and provide performance metrics. I then refactored this code to increase performance and simplify my original solution. This allowed for analysis of 2017 and 2018 ticker data at the click of a button.

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

**Analysis**
My original "AllStockAnalysis" VBA code (see below) declared the target worksheet, initialized a series of variables (including performance trackers, user input values, a standard ticker array, and pricing metrics). A For loop was then used to iterate through the values (based on the user inputted year) and calculate the total volume for each ticker in the array. For loops were also used to obtain a starting and ending price for each ticker based upon row positioning. The starting and ending price was then used to calculate a total return using the following formula: endingPrice / startingPrice - 1.

_Original Code_

Sub AllStocksAnalysis()
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   
   'Initialize start and end time variables
    Dim startTime As Single
    Dim endTime  As Single

   'Obtain user input for year (2017 or 2018)
   yearValue = InputBox("What year would you like to run the analysis on?")
   
   startTime = Timer
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(11) As String
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
   
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Sheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Sheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
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

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If
        
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Performance of Original Code
2017
![alt text](https://github.com/GrahamBSereno/Stock-Analysis/blob/main/Resources/PreRefactoring2017.png)

2018
![alt text](https://github.com/GrahamBSereno/Stock-Analysis/blob/main/Resources/PreRefactoring2018.png)

**Code Refactoring to Boost Efficiency**
To refactor the above code, I began by initializing a tickerIndex variable to allow for simplified ticker calling. I then created a series of output arrays that simplified the volume, starting price, and ending price calculations. These output arrays reference the singular tickerIndex variable to access the tickers for looping purposes. A series of IF statements were used to assess the starting and ending point row indexes for price gathering. Finally, a signular For loop was used to iterate through the entire list of tickers and populate ticker names, ticker volumes, and ticker returns. 

_Refactored Code_

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
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

Performance of Refactored Code
2017
![alt text](https://github.com/GrahamBSereno/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

2018
![alt text](https://github.com/GrahamBSereno/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)



**Summary**
The main advantages of refactoring code is simplification for the reader and increased efficiency due to reduced redundancy. The main disadvantages of refactoring code include the time lost refactoring code that already works and possible alteration of output accuracy.
When I refactored the original stock analysis code, I spent additional time refactoring (disadvantage). I was able to significantly improve the efficiency of the code because it took nearly half of the amount of time to execute once refactored. 

