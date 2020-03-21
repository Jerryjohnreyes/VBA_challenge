Attribute VB_Name = "Module1"
Sub results()
    Dim ticker As String
    Dim stock As Double
    Dim beginning_date As Long
    Dim end_date As Long
    Dim num_tickers As Integer
    
    Dim index_max_prct As Integer
    Dim index_min_prct As Integer
    Dim index_max_vol As Integer
    
    
    Dim columnResl(0 To 3) As String
    columnResl(0) = "Ticker"
    columnResl(1) = "Yearly Change"
    columnResl(2) = "Percent Change"
    columnResl(3) = "Total Stock Volume (Millions)"
    
    For Each ws In Worksheets
        num_tickers = 1
        stock = Round((ws.Cells(2, 7).Value / 1000000), 4)
        beginning_date = 2
        ticker = ws.Cells(2, 1).Value
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        
        For j = 0 To 3
            ws.Cells(1, 11 + j).Value = columnResl(j)
        Next j
        
        For i = 3 To last_row
            If ws.Cells(i, 1) = ws.Cells(i - 1, 1) Then
                stock = stock + Round((ws.Cells(i, 7).Value / 1000000), 4)
                end_date = i
            Else
                ws.Cells(num_tickers + 1, 11) = ticker
                ws.Cells(num_tickers + 1, 12) = ws.Cells(end_date, 6) - ws.Cells(beginning_date, 3)
                If ws.Cells(num_tickers + 1, 12) < 0 Then
                    ws.Cells(num_tickers + 1, 12).Interior.ColorIndex = 3
                Else
                    ws.Cells(num_tickers + 1, 12).Interior.ColorIndex = 4
                End If

                If ws.Cells(beginning_date, 3) = 0 Then
                    ws.Cells(num_tickers + 1, 13) = "NA"
                Else
                    ws.Cells(num_tickers + 1, 13) = Round((ws.Cells(num_tickers + 1, 12) / ws.Cells(beginning_date, 3) * 100), 2)
                End If
                ws.Cells(num_tickers + 1, 14) = stock
                stock = ws.Cells(i, 7)
                beginning_date = i
                end_date = i
                ticker = ws.Cells(i, 1).Value
                num_tickers = num_tickers + 1
            End If
        Next i
        
        max_prct = ws.Cells(2, 13).Value
        min_prct = ws.Cells(2, 13).Value
        max_vol = ws.Cells(2, 14).Value
        For i = 2 To (num_tickers - 1)
            If (ws.Cells(i, 13).Value > max_prct) Then
                max_prct = ws.Cells(i, 13).Value
                index_max_prct = i
            End If
            If (ws.Cells(i, 13).Value < min_prct) Then
                min_prct = ws.Cells(i, 13).Value
                index_min_prct = i
            End If
            If (ws.Cells(i, 14).Value > max_vol) Then
                max_vol = ws.Cells(i, 14).Value
                index_max_vol = i
            End If
        Next i
        ws.Cells(2, 18).Value = "Greatest percent Increased"
        ws.Cells(2, 19).Value = ws.Cells(index_max_prct, 11)
        ws.Cells(2, 20).Value = ws.Cells(index_max_prct, 13)
        ws.Cells(3, 18).Value = "Greatest percent Decreased"
        ws.Cells(3, 19).Value = ws.Cells(index_min_prct, 11)
        ws.Cells(3, 20).Value = ws.Cells(index_min_prct, 13)
        ws.Cells(4, 18).Value = "Greatest Volume"
        ws.Cells(4, 19).Value = ws.Cells(index_max_vol, 11)
        ws.Cells(4, 20).Value = ws.Cells(index_max_vol, 13)
    Next ws
End Sub
