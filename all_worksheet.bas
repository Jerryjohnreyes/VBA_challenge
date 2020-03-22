Sub results()
    Dim ticker As String
    Dim stock As Double
    Dim beginning_date As Long
    Dim end_date As Long
    Dim num_tickers As Integer
    
    
    'auxiliar variables to find the greatest chance within the
    'percent chance column and in the total_stock column
    Dim index_max_prct As Integer
    Dim index_min_prct As Integer
    Dim index_max_vol As Integer
    
    'for loop to repeat the process to all the worksheets
    For Each ws In Worksheets
        'counts the number of tickers the worksheets has at least one
        num_tickers = 1
        'stock initialized with the first value from the stock column
        stock = ws.Cells(2, 7).Value
        'beginning_date saves the index of the first date from each ticker
        beginning_date = 2
        ticker = ws.Cells(2, 1).Value
        'total of rows counting from the las value plus one to make the
        'following for loop count the last row of data
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        'changing the Number format to percent and a 10 multiple for easier opeartions
        Range(ws.Cells(2, 7), ws.Cells(last_row, 7)).NumberFormat = "0.00E+00"
        Range(ws.Cells(2, 13), ws.Cells(last_row, 13)).NumberFormat = "0.00%"
        Range(ws.Cells(2, 14), ws.Cells(last_row, 14)).NumberFormat = "0.0000E+00"
                
        'assing the header of the results for all sheets
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
    
        'values initialized at second row, for loop starts at 3 will stream to all tickers
        For i = 3 To last_row
            'comparing the i-th row with the last one until get a different ticker
            If ws.Cells(i, 1) = ws.Cells(i - 1, 1) Then
                'adds the i-th stock value to total stock
                stock = stock + ws.Cells(i, 7).Value
                'end_date keeps the row index for the present ticker
                end_date = i
            Else
                'when different it's time to assign results, first assing the ticker name
                ws.Cells(num_tickers + 1, 11) = ws.Cells(beginning_date, 1)
                'assign the difference between the last closing price with first opening price
                ws.Cells(num_tickers + 1, 12) = ws.Cells(end_date, 6) - ws.Cells(beginning_date, 3)
                'in column 14 assign the total stock for last ticker
                ws.Cells(num_tickers + 1, 14) = stock
                'Calculating the percent the difference is from the initial opening price
                'checking the initial value to be different from zero to assign the percent change
                'when value is 0 do not assing a relation
                If ws.Cells(beginning_date, 3) <> 0 Then
                    'for percent change we divide the difference by the initial value
                    ws.Cells(num_tickers + 1, 13) = ws.Cells(num_tickers + 1, 12) / ws.Cells(beginning_date, 3)
                End If
                'Assign a cell color format to the cells have an positive or negative value for difference column
                If ws.Cells(num_tickers + 1, 12) < 0 Then
                    ws.Cells(num_tickers + 1, 12).Interior.ColorIndex = 3
                Else
                    ws.Cells(num_tickers + 1, 12).Interior.ColorIndex = 4
                End If
                
                'finally, storing the new values for all the variables created with the current i-th ticker
                stock = ws.Cells(i, 7)
                beginning_date = i
                end_date = i
                num_tickers = num_tickers + 1
            End If
        Next i
        
        'searching for max increase, decrease, stock from the
        'results column 13 percent change and column 14 total stock
        max_prct = ws.Cells(2, 13).Value
        min_prct = ws.Cells(2, 13).Value
        max_vol = ws.Cells(2, 14).Value
        For i = 2 To (num_tickers - 1)
            'searching for the index of the max increase
            If (ws.Cells(i, 13).Value > max_prct) Then
                max_prct = ws.Cells(i, 13).Value
                index_max_prct = i
            End If
            'searching for the index maximun decrease
            If (ws.Cells(i, 13).Value < min_prct) Then
                min_prct = ws.Cells(i, 13).Value
                index_min_prct = i
            End If
            'searching for the maximun total stock
            If (ws.Cells(i, 14).Value > max_vol) Then
                max_vol = ws.Cells(i, 14).Value
                index_max_vol = i
            End If
        Next i
        
        'Printing the values for greatest decrease, increase and total stock
        Range(ws.Cells(2, 20), ws.Cells(3, 20)).NumberFormat = "0.00%"
        Range(ws.Cells(4, 20), ws.Cells(4, 20)).NumberFormat = "0.0000E+00"
        ws.Cells(2, 18).Value = "Greatest percent Increased"
        ws.Cells(2, 19).Value = ws.Cells(index_max_prct, 11)
        ws.Cells(2, 20).Value = ws.Cells(index_max_prct, 13)
        ws.Cells(3, 18).Value = "Greatest percent Decreased"
        ws.Cells(3, 19).Value = ws.Cells(index_min_prct, 11)
        ws.Cells(3, 20).Value = ws.Cells(index_min_prct, 13)
        ws.Cells(4, 18).Value = "Greatest Volume"
        ws.Cells(4, 19).Value = ws.Cells(index_max_vol, 11)
        ws.Cells(4, 20).Value = ws.Cells(index_max_vol, 14)
    Next ws
End Sub
