Sub stocktracker()
    ' Define all the vars we want to use
    Dim sheetExists As Boolean
    Dim yearExists As Boolean
    Dim sheetToFind As String
    Dim Ticker As String ' Stock Name
    Dim tickerCount As Double ' Counter for tickers
    Dim stockTotal As Double ' Total value of stocks
    Dim stockOpen As Double ' Opening value of stocks
    Dim stockClose As Double ' Close value of stocks
    Dim stockPcChange As Double ' Percentage change in stock value
    Dim Summary_Table_Row As Integer ' End table that we're dumping our data into
    Dim increasePercent As Double ' Greatest % increase
    Dim decreasePercent As Double ' Greatest % decrease
    Dim mostStocks As Double ' Most stocks traded
    Dim Top_Stat_Row As Integer ' "Greatest" table
    Dim yearCheck As String
    Dim yearValue As String

    StockTotal = 0 ' You're going to see this being set to 0 a bit, thats because I kept getting Overflows and I'm just tired of troubleshooting at this point

    ' We're checking to see if the Final_Data sheet has been created previously. This allows us to keep using the same summary sheet and keep adding yearly data sheets
    sheetExists = False
    sheetToFind = "Final_Data"
    For Each ws In Worksheets
        If sheetToFind = ws.Name Then
            sheetExists = True ' If sheet exists, then we don't need to create the Final_Data page
        End If
    Next ws

    ' If the above checks failed, then we're creating the final data page and all associated fun stuff
    If sheetExists = False Then

        ' Create a page where we are dumping our final data into
        Sheets.Add.Name = "Final_Data" ' Lets name this sheet Final_Data
        Set final_data = Worksheets("Final_Data") ' Defining it to make it easier to dump data
        final_data.Move Before:=Worksheets(1) ' Putting it at the start
        final_data.Columns("F").NumberFormat = "0.00%" ' Set Column F to be a percentage
        final_data.Range("K2").NumberFormat = "0.00%" ' Set Column K to be a percentage
        final_data.Range("K3").NumberFormat = "0.00%" ' Set Column K to be a percentage
        ' Create a header
        final_data.Range("A1").Value = "<ticker>"
        final_data.Range("B1").Value = "<year>"
        final_data.Range("C1").Value = "<year_stock_volume>"
        final_data.Range("D1").Value = "<year_open_price>"
        final_data.Range("E1").Value = "<year_close_price>"
        final_data.Range("F1").Value = "<year_price_change_%>"
        final_data.Range("G1").Value = "<year_price_change>"
        final_data.Range("J1").Value = "<ticker>"
        final_data.Range("K1").Value = "<value>"
        final_data.Range("L1").Value = "<year>"
        final_data.Range("I2").Value = "Overall Greatest % Increase"
        final_data.Range("I3").Value = "Overall Greatest % Decrease"
        final_data.Range("I4").Value = "Overall Greatest Total Stocks Sold"
        final_data.Range("L2:L4").Value = "tba"
        Summary_Table_Row = 2 ' Row number for dumping final data into
        Top_Stat_Row = 5
    Else
        Set final_data = Worksheets("Final_Data") ' Defining it to make it easier to dump data
        Summary_Table_Row = final_data.Cells(Rows.Count, "A").End(xlUp).Row + 1 ' Calculate all rows
        Top_Stat_Row = final_data.Cells(Rows.Count, "L").End(xlUp).Row + 1 ' Calculate all rows
    End If


    ' For this task we're rolling through all worksheets
    For Each ws In Worksheets

        ' Ok look I know I'm nulling these vars multiple times, but if I dont zero the stockTotal I get an overflow error when I'm adding the Totals and I'm tired of troubleshooting
        ' I probably don't need to zero all of these (or any of these) but it's superstition at this point
        tickerCount = 0 ' Set ticker count to 0
        stockTotal = 0 ' Set initial value of stocks to 0
        increasePercent = 0
        decreasePercent = 0
        mostStocks = 0
        stockOpen = 0
        stockClose = 0
        stockPcChange = 0
        tickerCount = 0

        ' Now we're going to check to see if the year has already previously been calculated
        ' If the yearly sheets gets updated, we won't detect that
        ' It takes the Greatest Table and looks through the years present in those datasets and compares it to the sheet name
        yearExists = False
        yearValue = ws.Name ' ws.Name not a string? Or something? Need to force it into a string var, going str(ws.name) wasn't working??
        For ycCounter = 2 To (Top_Stat_Row - 1)
            yearCheck = final_data.Range("L" & ycCounter).Value
            If yearCheck = yearValue Then
                yearExists = True
                Exit For
            End If
        Next ycCounter

        If (ws.Name <> "Final_Data") And (yearExists = False) Then ' If the sheet isn't Final_Data or already present in the Greatest Table
            ' Initial worksheet cleaning and calcs
            lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1 ' Calculate all rows

            ' These things again :(
            tickerCount = 0 ' Set ticker count to 0
            stockTotal = 0 ' Set initial value of stocks to 0
            increasePercent = 0
            decreasePercent = 0
            mostStocks = 0
            final_data.Range("L" & Top_Stat_Row & ":L" & Top_Stat_Row + 2).Value = ws.Name
            final_data.Range("I" & Top_Stat_Row).Value = ws.Name & " Greatest % Increase"
            final_data.Range("I" & Top_Stat_Row + 1).Value = ws.Name & " Greatest % Decrease"
            final_data.Range("I" & Top_Stat_Row + 2).Value = ws.Name & " Greatest Total Stocks Sold"

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
                    If Not stockOpen = 0 Then
                        stockPcChange = (stockClose - stockOpen) / stockOpen
                    Else
                        stockPcChange = 0
                    End If

                    ' Print the Ticker in the test table
                    final_data.Range("A" & Summary_Table_Row).Value = Ticker

                    ' Print the stock total in the test table
                    final_data.Range("C" & Summary_Table_Row).Value = stockTotal
                    final_data.Range("B" & Summary_Table_Row).Value = ws.Name ' Putting the worksheet name into the 'Year' Column
                    final_data.Range("D" & Summary_Table_Row).Value = stockOpen
                    final_data.Range("E" & Summary_Table_Row).Value = stockClose
                    final_data.Range("F" & Summary_Table_Row).Value = stockPcChange
                    final_data.Range("G" & Summary_Table_Row).Value = stockClose - stockOpen

                    If stockPcChange > 0 Then
                        final_data.Range("F" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                        final_data.Range("G" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                    ElseIf stockPcChange = 0 Then
                        final_data.Range("F" & Summary_Table_Row).Interior.ColorIndex = 46 ' Orange
                        final_data.Range("G" & Summary_Table_Row).Interior.ColorIndex = 46 ' Orange
                    Else
                        final_data.Range("F" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                        final_data.Range("G" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                    End If

                    ' Populating this years copy of the greatest table
                    If stockPcChange > increasePercent Then
                        increasePercent = stockPcChange
                        final_data.Range("J" & Top_Stat_Row).Value = Ticker
                        final_data.Range("K" & Top_Stat_Row).Value = increasePercent
                    End If
                    If stockPcChange < decreasePercent Then
                        decreasePercent = stockPcChange
                        final_data.Range("J" & Top_Stat_Row + 1).Value = Ticker
                        final_data.Range("K" & Top_Stat_Row + 1).Value = decreasePercent
                    End If
                    If stockTotal > mostStocks Then
                        mostStocks = stockTotal
                        final_data.Range("J" & Top_Stat_Row + 2).Value = Ticker
                        final_data.Range("K" & Top_Stat_Row + 2).Value = mostStocks
                    End If

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

            final_data.Range("K" & Top_Stat_Row).NumberFormat = "0.00%" ' Set Column K to be a percentage
            final_data.Range("K" & Top_Stat_Row + 1).NumberFormat = "0.00%" ' Set Column K to be a percentage

            ' This year is done, prepare to add details for the next year
            
            ' Comparing this years greatest stats to the overall calculated greatest stats
            If final_data.Range("K" & Top_Stat_Row).Value > final_data.Range("K2").Value Then
                final_data.Range("J2").Value = final_data.Range("J" & Top_Stat_Row).Value
                final_data.Range("K2").Value = final_data.Range("K" & Top_Stat_Row).Value
                final_data.Range("L2").Value = ws.Name
            End If
            If final_data.Range("K" & Top_Stat_Row + 1).Value < final_data.Range("K3").Value Then
                final_data.Range("J3").Value = final_data.Range("J" & Top_Stat_Row + 1).Value
                final_data.Range("K3").Value = final_data.Range("K" & Top_Stat_Row + 1).Value
                final_data.Range("L3").Value = ws.Name
            End If
            If final_data.Range("K" & Top_Stat_Row + 2).Value > final_data.Range("K4").Value Then
                final_data.Range("J4").Value = final_data.Range("J" & Top_Stat_Row + 2).Value
                final_data.Range("K4").Value = final_data.Range("K" & Top_Stat_Row + 2).Value
                final_data.Range("L4").Value = ws.Name
            End If

            Top_Stat_Row = Top_Stat_Row + 3 ' To allow us to build the next years greatest table
        
        End If

    Next ws

End Sub

