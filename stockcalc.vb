Sub stocktracker()

    Dim Ticker As String ' Stock Name
    Dim tickerCount as DOuble ' Counter tickers
    tickerCount = 0
    Dim stockTotal As Double ' Total value of stocks
    Dim stockOpen as Double ' Opening value of stocks
    Dim stockClose as Double ' Close value of stocks
    Dim stockPcChange as Double ' Percentage change in stock value 
    stockTotal = 0 ' Set initial value of stocks to 0
    Dim Summary_Table_Row As Integer ' Temp table for initial testing 
    Summary_Table_Row = 2 ' Table for initial testing 

    ' Calculate all rows
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1

    ' Loop through all rows
    For counter = 2 To lastRow

        ' Check to see if the ticker (at counter+1,1) is new
        If Cells(counter + 1, 1).Value <> Cells(counter, 1).Value Then

            ' Set the Ticker
            Ticker = Cells(counter, 1).Value

            ' Add the stock total
            stockTotal = stockTotal + Cells(counter, 7).Value

            stockOpen = Cells((counter - tickerCount), 3).Value
            stockClose = Cells(counter, 6).Value
            stockPcChange = stockClose / stockOpen

            ' Print the Ticker in the test table 
            Range("J" & Summary_Table_Row).Value = Ticker

            ' Print the stock total in the test table 
            Range("K" & Summary_Table_Row).Value = stockTotal

            Range("L" & Summary_Table_Row).Value = stockOpen
            Range("M" & Summary_Table_Row).Value = stockClose
            Range("N" & Summary_Table_Row).Value = stockPcChange            

            ' Add one to the summary table row 
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the values
            stockTotal = 0
            stockOpen = 0
            stockClose = 0
            stockPcChange = 0
            tickerCount = 0

            ' If the cell immediately following a row is the ticker
            Else

            ' Add to the stockTotal
            stockTotal = stockTotal + Cells(counter, 7).Value
            tickerCount = tickerCount + 1


        End If

    Next counter

End Sub