Sub testing():

Dim total_volumn, max_total_volumn, last_row As LongLong
Dim year_open, year_close, year_change, percentage_change, greatest_inc, greatest_dec As Double
Dim ticker, inc_ticker, dec_ticker, max_ticker As String
Dim write_row, bonus_row As Long
    
For Each ws In Worksheets:

    total_volumn = 0
    max_total_volumn = 0
    greatest_dec = 0
    greatest_inc = 0
    write_row = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Year Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volumn"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volumn"
    
    year_open = ws.Range("C2")
    
    For r = 2 To last_row:
        total_volumn = total_volumn + ws.Cells(r, 7).Value
        ticker = ws.Cells(r, 1).Value
        If ticker <> ws.Cells(r + 1, 1).Value Then
            year_close = ws.Cells(r, 6).Value
            year_change = year_close - year_open
            percentage_change = year_change / year_open
            ws.Cells(write_row, 9) = ticker
            ws.Cells(write_row, 10) = year_change
            ws.Cells(write_row, 11) = percentage_change
            ws.Cells(write_row, 11).NumberFormat = "0.00%"
            ws.Cells(write_row, 12) = total_volumn
            
            If year_change < 0 Then
                ws.Cells(write_row, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(write_row, 10).Interior.ColorIndex = 4
            End If
            
            If total_volumn > max_total_volumn Then
                max_total_volumn = total_volumn
                max_ticker = ticker
            End If
                  
            If percentage_change > greatest_inc Then
                greatest_inc = percentage_change
                inc_ticker = ticker
            ElseIf percentage_change < greatest_dec Then
                greatest_dec = percentage_change
                dec_ticker = ticker
            End If
            
            total_volumn = 0
            year_open = ws.Cells(r + 1, 3)
            write_row = write_row + 1
            
            

        End If
    
    Next r

    ws.Range("O2") = inc_ticker
    ws.Range("O3") = dec_ticker
    ws.Range("O4") = max_ticker
    ws.Range("P2") = greatest_inc
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3") = greatest_dec
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("P4") = max_total_volumn
    
    ws.Range("A1:P1").EntireColumn.AutoFit
Next ws

End Sub
