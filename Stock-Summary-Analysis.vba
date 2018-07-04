Sub GetStockSummary():
    
    ' Define variables to store ticker_name, open, close, yearly_change, percent_change, volume and related
    Dim ticker_name As String
    Dim open_value, close_value, yearly_change, percent_change, greatest_percent_increase_value, greatest_percent_decrease_value As Double
    Dim volume, total_volume, greatest_total_volume_value As LongLong
    
    'Variable to keep track of current row when writing unique ticker
    Dim current_row As Integer
    
    'Variable to keerp track of open value for new ticker
    Dim new_ticker_start_row As Long
        
    'Loop over all the worksheets
    For Each ws In Worksheets:
        
        'Initailize and Reset all the variables for each new worksheet
        ticker_name = ""
        open_value = 0
        close_value = 0
        yearly_change = 0
        percent_change = 0
        greatest_percent_increase_value = 0
        greatest_percent_decrease_value = 0
        volume = 0
        total_volume = 0
        greatest_total_volume_value = 0
    
        current_row = 2
        
        new_ticker_start_row = 2
        
        'Create new header in the current sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
         'Get the Last Row in each worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'For loop over rows
        For r = 2 To lastRow
            
            'Get the volume for each stock row
            volume = ws.Cells(r, 7).Value
            total_volume = total_volume + volume
        
            If (ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value) Then
                
                'New Ticker found
                ticker_name = ws.Cells(r, 1).Value
                ws.Cells(current_row, 9).Value = ticker_name
                
                'Get open value for the ticker from the 3rd column
                open_value = ws.Cells(new_ticker_start_row, 3).Value
            
                'Get close value for the ticker from the 6th column
                close_value = ws.Cells(r, 6).Value
                
                'Calculate yearly change
                yearly_change = close_value - open_value
                ws.Cells(current_row, 10).Value = yearly_change
                
                'Conditional formation based on yearly change
                If (yearly_change >= 0) Then
                    ws.Cells(current_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(current_row, 10).Interior.ColorIndex = 3
                End If
                
                'Calculate percent change
                If (open_value <> 0) Then
                    percent_change = ((close_value - open_value) / open_value)
                Else
                    percent_change = 1#
                End If
                ws.Cells(current_row, 11).Value = percent_change
                ws.Cells(current_row, 11).NumberFormat = "0.00%"
                
                'Comparing and finding Greatest percent change Increase value and ticker name
                If (percent_change > greatest_percent_increase_value) Then
                    greatest_percent_increase_value = percent_change
                    greatest_percent_increase_ticker = ticker_name
                End If
                
                'Comparing and finding Greatest percent change Decrease value and ticker name
                If (percent_change < greatest_percent_decrease_value) Then
                    greatest_percent_decrease_value = percent_change
                    greatest_percent_decrease_ticker = ticker_name
                End If
                
                'Set the total volume for this ticker
                ws.Cells(current_row, 12).Value = total_volume
                
                'Comparing and finding Greatest total volume value and ticker name
                If (total_volume > greatest_total_volume_value) Then
                    greatest_total_volume_value = total_volume
                    greatest_total_volume_ticker = ticker_name
                End If
                
                'Reset the total_volume for next stock
                total_volume = 0
                
                'Increment the current_row to write next stock/ticker
                current_row = current_row + 1
                
                'New ticker start row
                new_ticker_start_row = r + 1
                
            End If
        
        Next r
        
        'Autofit columns: I, J, K and L
        ws.Range("I:L").Columns.AutoFit
        
        'Summary
        ws.Range("P1").Value = "Ticker Name"
        ws.Range("Q1").Value = "Value"
        
        'Greatest Increase in percent ticker name and value
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = greatest_percent_increase_ticker
        ws.Range("Q2").Value = greatest_percent_increase_value
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'Greatest Decrease in percent ticker name and value
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = greatest_percent_decrease_ticker
        ws.Range("Q3").Value = greatest_percent_decrease_value
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Greatest Total Volume ticker name and value
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = greatest_total_volume_ticker
        ws.Range("Q4").Value = greatest_total_volume_value
        
        'Autofit columns: O, P and Q
        ws.Range("O:Q").Columns.AutoFit
        
    Next ws
    
End Sub
