Attribute VB_Name = "Module2"
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
    
    num_tickers = 1
    stock = Round((Cells(2, 7).Value / 1000000), 4)
    beginning_date = 2
    ticker = Cells(2, 1).Value
    last_row = Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        
    For j = 0 To 3
        Cells(1, 11 + j).Value = columnResl(j)
    Next j
        
    For i = 3 To last_row
        If Cells(i, 1) = Cells(i - 1, 1) Then
            stock = stock + Round((Cells(i, 7).Value / 1000000), 4)
            end_date = i
        Else
            Cells(num_tickers + 1, 11) = ticker
            Cells(num_tickers + 1, 12) = Cells(end_date, 6) - Cells(beginning_date, 3)
            If Cells(num_tickers + 1, 12) < 0 Then
                Cells(num_tickers + 1, 12).Interior.ColorIndex = 3
            Else
                Cells(num_tickers + 1, 12).Interior.ColorIndex = 4
            End If

            If Cells(beginning_date, 3) = 0 Then
                Cells(num_tickers + 1, 13) = "NA"
            Else
                Cells(num_tickers + 1, 13) = Round((Cells(num_tickers + 1, 12) / Cells(beginning_date, 3) * 100), 2)
            End If
            Cells(num_tickers + 1, 14) = stock
            stock = Cells(i, 7)
            beginning_date = i
            end_date = i
            ticker = Cells(i, 1).Value
            num_tickers = num_tickers + 1
        End If
    Next i
        
    max_prct = Cells(2, 13).Value
    min_prct = Cells(2, 13).Value
    max_vol = Cells(2, 14).Value
    For i = 2 To (num_tickers - 1)
        If (Cells(i, 13).Value > max_prct) Then
            max_prct = Cells(i, 13).Value
            index_max_prct = i
        End If
        If (Cells(i, 13).Value < min_prct) Then
            min_prct = Cells(i, 13).Value
            index_min_prct = i
        End If
        If (Cells(i, 14).Value > max_vol) Then
            max_vol = Cells(i, 14).Value
            index_max_vol = i
        End If
    Next i
    Cells(2, 18).Value = "Greatest percent Increased"
    Cells(2, 19).Value = Cells(index_max_prct, 11)
    Cells(2, 20).Value = Cells(index_max_prct, 13)
    Cells(3, 18).Value = "Greatest percent Decreased"
    Cells(3, 19).Value = Cells(index_min_prct, 11)
    Cells(3, 20).Value = Cells(index_min_prct, 13)
    Cells(4, 18).Value = "Greatest Volume"
    Cells(4, 19).Value = Cells(index_max_vol, 11)
    Cells(4, 20).Value = Cells(index_max_vol, 13)
End Sub

