Sub ticker_report()
  For Each ws In Worksheets
    
        Dim last_row As Long
        Dim ticker_symbol As String
        Dim year_open As Double
        Dim year_close As Double
        Dim year_change As Double
        Dim percent_change As Double
        Dim next_ticker As Integer
        Dim ticker_summary As String
        Dim ticker_volume As Double
        Dim ticker_rc_qty As Double
        
        'setting_resetting the variables
            ticker_volume = 0
            ticker_summary = 2
            year_open = 0
            year_close = 0
            
        'Finding last row
            last_row = Cells(Rows.Count, 1).End(xlUp).Row
            
        'Loop for all rows
        For i = 1 To last_row
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Set the ticker-name
                    ticker_symbol = ws.Cells(i + 1, 1).Value
                    
                    'capturing year open and close
                    year_open = ws.Cells(i + 1, 3).Value
                    ticker_rc_qty = WorksheetFunction.CountIf(ws.Range("A:A"), ticker_symbol)
                    year_close = ws.Cells(i + ticker_rc_qty, 6).Value
                    
                    'calculating Yearly change
                    year_change = year_close - year_open
                    
                    'calcualting Percent change
                        If year_open = 0 Then
                            percent_change = 0
                        Else
                            percent_change = year_change / year_open
                        End If
                    
                    'Volume
                    ticker_volume = ticker_volume + Cells(i + 1, 7).Value
                    
                    'populating the summary table
                    'values
                    ws.Cells(ticker_summary, 9).Value = ticker_symbol
                    ws.Cells(ticker_summary, 10).Value = year_change
                    ws.Cells(ticker_summary, 11).Value = percent_change
                    ws.Cells(ticker_summary - 1, 12).Value = ticker_volume
                          
                    'format
                        If ws.Cells(ticker_summary, 10).Value >= 0 Then
                          ws.Cells(ticker_summary, 10).Interior.ColorIndex = 4
                        Else
                          ws.Cells(ticker_summary, 10).Interior.ColorIndex = 3
                        End If
                    ws.Cells(ticker_summary, 10).NumberFormat = "0.00"
                    ws.Cells(ticker_summary, 11).NumberFormat = "0.00%"
                    ws.Cells(ticker_summary - 1, 12).NumberFormat = "###,###,####,####,###"
                    
                    'Counter
                    ticker_summary = ticker_summary + 1
                    ticker_volume = 0
                Else
                    ticker_volume = ticker_volume + Cells(i + 1, 7).Value
                End If
        Next i
            'Setting the table's headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Charge"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            'clean up -- Removing last record
            last_entry = ws.Cells(Rows.Count, 10).End(xlUp).Row
            Debug.Print last_entry
            ws.Cells(last_entry, 10).Clear
            last_entry = ws.Cells(Rows.Count, 11).End(xlUp).Row
            ws.Cells(last_entry, 11).Clear
            ws.Columns("I:R").EntireColumn.AutoFit
        '----------------Bonus point--------------
        
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Total_Volume As Double
        
        'setting the table's headers and labels
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        
        'capturing maxs and mins
        Greatest_Increase = WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(last_entry - 1, 11)))
        Greatest_Decrease = WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(last_entry - 1, 11)))
        Greatest_Total_Volume = WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(last_entry - 1, 12)))
        
        'format
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 18).NumberFormat = "###,###,####,####,###"
        
        'populating the summary table
        ws.Cells(2, 18).Value = Greatest_Increase
        ws.Cells(3, 18).Value = Greatest_Decrease
        ws.Cells(4, 18).Value = Greatest_Total_Volume
                
        'capturing ticker
        report_total = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To report_total
            If Greatest_Increase = ws.Cells(j, 11).Value Then
                ws.Cells(2, 17).Value = ws.Cells(j, 9).Value
                   
            ElseIf Greatest_Decrease = ws.Cells(j, 11).Value Then
                ws.Cells(3, 17).Value = ws.Cells(j, 9).Value
                
            ElseIf Greatest_Total_Volume = ws.Cells(j, 12).Value Then
                ws.Cells(4, 17).Value = ws.Cells(j, 9).Value
            End If
        Next j
        
  Next ws
End Sub

