Sub stocktracker()
    ' Define all the vars we want to use
    Dim Ticker As String ' Stock Name
    Dim tickerCount as Double ' Counter for tickers
    Dim stockTotal As Double ' Total value of stocks
    Dim stockOpen as Double ' Opening value of stocks
    Dim stockClose as Double ' Close value of stocks
    Dim stockPcChange as Double ' Percentage change in stock value 
    Dim Summary_Table_Row As Integer ' End table that we're dumping our data into 

    
    Summary_Table_Row = 2 ' Row number for dumping final data into

    ' Create a page where we are dumping our final data into
    Sheets.Add.Name = "Final_Data" ' Lets name this sheet Final_Data
    set final_data = worksheets("Final_Data") ' Defining it to make it easier to dump data
    final_data.Move Before:=worksheets(1) ' Putting it at the start
    final_data.Columns("F").NumberFormat = "0.00%" ' Set Column N to be a percentage
    ' Create a header
    final_data.Range("A1").Value = "<ticker>"
    final_data.Range("B1").Value = "<year>"
    final_data.Range("C1").Value = "<year_stock_volume>"
    final_data.Range("D1").Value = "<year_open_price>"
    final_data.Range("E1").Value = "<year_close_price>"
    final_data.Range("F1").Value = "<year_price_change>"

    ' For this task we're rolling through all worksheets
    for each ws in worksheets

        ' Initial worksheet cleaning and calcs
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1 ' Calculate all rows
        tickerCount = 0 ' Set ticker count to 0
        stockTotal = 0 ' Set initial value of stocks to 0
        

        ' Loop through all rows
        For counter = 2 To lastRow

            ' Check to see if the ticker (at counter+1,1) is new
            If ws.Cells(counter + 1, 1).Value <> ws.Cells(counter, 1).Value Then

                Ticker = ws.Cells(counter, 1).Value ' Set the Ticker

                stockTotal = stockTotal + ws.Cells(counter, 7).Value ' Add the stock total. Also, why does vscode print this 7 as blue?
                stockOpen = ws.Cells((counter - tickerCount), 3).Value ' We go back to the first instance of the ticker to get the opening value
                stockClose = ws.Cells(counter, 6).Value ' Last value of the ticker close is the close value

                ' If stockOpen is 0 then it causes an overflow.
                ' "Oh what stock will ever be worth $0", yeah well apparently 2014 was rough to Planet Fitness
                if not stockOpen = 0 then
                    stockPcChange = (stockClose - stockOpen) / stockOpen
                else
                    stockPcChange = 0
                end if 

                ' Print the Ticker in the test table 
                final_data.Range("A" & Summary_Table_Row).Value = Ticker

                ' Print the stock total in the test table 
                final_data.Range("C" & Summary_Table_Row).Value = stockTotal
                final_data.Range("B" & Summary_Table_Row).Value = ws.Name ' Putting the worksheet name into the 'Year' Column
                final_data.Range("D" & Summary_Table_Row).Value = stockOpen
                final_data.Range("E" & Summary_Table_Row).Value = stockClose
                final_data.Range("F" & Summary_Table_Row).Value = stockPcChange

                if stockPcChange > 0 then
                    final_data.Range("F" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                ElseIf stockPcChange = 0 then
                    final_data.Range("F" & Summary_Table_Row).Interior.ColorIndex = 46 ' Orange
                Else
                    final_data.Range("F" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                end if             

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

                stockTotal = stockTotal + ws.Cells(counter, 7).Value ' Add to the stockTotal
                tickerCount = tickerCount + 1 ' Just keep increasing until we get to the end


            End If

        Next counter
    Next ws 

End Sub