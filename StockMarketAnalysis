Sub stock_market_analysis()

'define everything
Dim ws As Worksheet
Dim ticker As String
Dim summary_row As Double
Dim i As Double
Dim vol_total As Double
Dim open_price As Double
Dim close_price As Double
Dim max_per As Double
Dim min_per As Double
Dim max_vol As Double
Dim max_per_ticker As String
Dim min_per_ticker As String
Dim max_vol_ticker As String

'loop for each workssheet
'set up headers
For Each ws In ThisWorkbook.Sheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    i = 2
    summary_row = 2
    vol_total = 0
    
    'loop while ticker value is not empty
    Do While ws.Cells(i, 1).Value <> ""
    
        'pick out tickers without same ticker prior and count vol
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            open_price = ws.Cells(i, 3).Value
            
            vol_total = vol_total + ws.Cells(i, 7).Value
            
        'pick out tickers without same ticker after
        'calculate and print totals
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
            ws.Range("I" & summary_row).Value = ticker
            
            vol_total = vol_total + ws.Cells(i, 7).Value
            ws.Range("L" & summary_row).Value = vol_total
            
            close_price = ws.Cells(i, 6).Value
            ws.Range("J" & summary_row).Value = close_price - open_price
            
            'format
            If ws.Range("J" & summary_row).Value > 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & summary_row).Value < 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
            End If
            
            If open_price = 0 Then
                ws.Range("K" & summary_row).Value = "N/A"
            Else
                ws.Range("K" & summary_row).Value = ws.Range("J" & summary_row).Value / open_price
                ws.Range("K" & summary_row).NumberFormat = "0.00%"
            End If
            
            'move to next summary row
            summary_row = summary_row + 1
            vol_total = 0
            
        'calculate vol of remaining cells with same tickers adjacent
        Else
            vol_total = vol_total + ws.Cells(i, 7).Value
            
        End If
        i = i + 1
    Loop
    
    'reset i and set max/min value counters
    i = 2
    max_per = 0
    min_per = 0
    max_vol = 0
    
    'loop while totals are not empty
    Do While ws.Cells(i, 9).Value <> ""
        
        'pick out the winners
        If ws.Cells(i, 11).Value = "N/A" Then
            max_per = max_per
        ElseIf ws.Cells(i, 11).Value > max_per Then
            max_per = ws.Cells(i, 11).Value
            max_per_ticker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 11).Value < min_per Then
            min_per = ws.Cells(i, 11).Value
            min_per_ticker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value > max_vol Then
            max_vol = ws.Cells(i, 12).Value
            max_vol_ticker = ws.Cells(i, 9).Value
        End If
        
        i = i + 1
    Loop
    
    'print results
    ws.Range("Q2").Value = max_per
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = max_per_ticker
    ws.Range("Q3").Value = min_per
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = min_per_ticker
    ws.Range("Q4").Value = max_vol
    ws.Range("P4").Value = max_vol_ticker

    'make columns look nice
    ws.Cells.EntireColumn.AutoFit
    
'yay next worksheet
Next ws

End Sub