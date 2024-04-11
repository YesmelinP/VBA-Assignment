Attribute VB_Name = "Module1"
Sub stockmarket()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim tickerVolume As Double
    Dim tickerOpen As Double
    Dim tickerClose As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim greatestDecrease As Double
    Dim greatesIncrease As Double
    Dim greatesTotalVolume As Double
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
   'ticker, yearly change, percent change, total stock volume
   
    For Each ws In ThisWorkbook.Worksheets
    summary_table_row = 2
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'assign the header values
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "yearly_change"
        ws.Cells(1, 11).Value = "percent_change"
        ws.Cells(1, 12).Value = "total_stock_volume"
        
        ticker = ws.Cells(2, 1).Value
        tickerVolume = 0
        tickerOpen = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
            tickerVolume = tickerVolume + ws.Cells(i, 7)
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'open & close calculations
                tickerClose = ws.Cells(i, 6)
                yearlyChange = tickerClose - tickerOpen
                percentChange = (tickerClose - tickerOpen) / tickerOpen
                
                'formatting cells
                    ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                    If yearlyChange >= 0 Then
                        ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                    End If
                    
                
                'assign values
              ws.Cells(summary_table_row, 9).Value = ticker
              ws.Cells(summary_table_row, 10).Value = yearlyChange
              ws.Cells(summary_table_row, 11).Value = percentChange
              ws.Cells(summary_table_row, 12).Value = tickerVolume
                'increment counters
              summary_table_row = summary_table_row + 1
                'replace data for ticker change
              ticker = ws.Cells(i + 1, 1).Value
              tickerVolume = 0
              tickerOpen = ws.Cells(i + 1, 3).Value
              tickerClose = 0
            End If
        Next i
        
        'summary table
        ws.Cells(2, 15).Value = "Greates % Increase"
        ws.Cells(3, 15).Value = "Greates % Decrease"
        ws.Cells(4, 15).Value = "Greates Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'initial assignation of summary
        greatestDecrease = ws.Cells(2, 11).Value
        greatesIncrease = ws.Cells(2, 11).Value
        greatesTotalVolume = ws.Cells(2, 12).Value
        
        'formatting summary table
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'loop of summary table
        For i = 2 To summary_table_row
            If ws.Cells(i, 11).Value > greatesIncrease Then
                greatesIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = greatesIncrease
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < greatesDecrease Then
                greatesDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 17).Value = greatesDecrease
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > greatesTotalVolume Then
                greatesTotalVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 17).Value = greatesTotalVolume
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
    Next ws
End Sub

