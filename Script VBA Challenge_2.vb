Sub tarea()

    Dim year_change As Double
    Dim percentage_change As Double
    Dim total_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim ticker As String
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String
    
    Dim greatest_increase_value As Double
    Dim greatest_decrease_value As Double
    Dim greatest_volume_value As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Dim summary_row As Long
    
    
    'Set ws = ThisWorkbook.sheets("A")
    For Each ws In ThisWorkbook.Worksheets
    
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
     
        ws.Cells(1, 9).value = "Ticker"
        ws.Cells(1, 10).value = "Yearly Change"
        ws.Cells(1, 11).value = "Percent Change"
        ws.Cells(1, 12).value = "Total Stock Volume"
        ws.Cells(1, 15).value = "Ticker"
        ws.Cells(1, 16).value = "Value"
        
        
        summary_row = 2
        greatest_increase_value = 0
        greatest_decrease_value = 0
        greatest_volume_value = 0
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).value <> ws.Cells(i - 1, 1).value Then
                year_change = close_price - open_price
                If open_price <> 0 Then
                    percentage_change = year_change / open_price
                Else
                    percentage_change = 0
            
                End If
            
                ws.Cells(summary_row, 9).value = ticker
                ws.Cells(summary_row, 10).value = year_change
                ws.Cells(summary_row, 11).value = percentage_change
                ws.Cells(summary_row, 12).value = total_volume
            
                If percentage_change > greatest_increase_value Then
                    greatest_increase_value = percentage_change
                    greatest_increase_ticker = ticker
            
                End If
                
                
                If percentage_change < greatest_decrease_value Then
                    greatest_decrease_value = percentage_change
                    greatest_decrease_ticker = ticker
                
                End If
                
                If total_volume > greatest_volume_value Then
                    greatest_volume_value = total_volume
                    greatest_volume_ticker = ticker
                
                End If
                
                If year_change > 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                
                ElseIf year_change = 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 2
        
                Else
        
                ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            
                End If
                
                summary_row = summary_row + 1
                ticker = ws.Cells(i, 1).value
                open_price = ws.Cells(i, 3).value
                close_price = ws.Cells(i, 6).value
                total_volume = ws.Cells(i, 7).value
                
            Else
                close_price = ws.Cells(i, 6).value
                total_volume = total_volume + ws.Cells(i, 7).value
            
            End If
    
    
        Next i
    
        ws.Cells(2, 14).value = "Greatest % increase"
        ws.Cells(3, 14).value = "Greatest % decrease"
        ws.Cells(4, 14).value = "Greatest total volume"
        
        ws.Cells(2, 15).value = greatest_increase_ticker
        ws.Cells(3, 15).value = greatest_decrease_ticker
        ws.Cells(4, 15).value = greatest_volume_ticker
        
        ws.Cells(2, 16).value = greatest_increase_value
        ws.Cells(3, 16).value = greatest_decrease_value
        ws.Cells(4, 16).value = greatest_volume_value
    
        ws.Range("K2:K" & summary_row).NumberFormat = "0.00%"
        ws.Range("P2", "P3").NumberFormat = "0.00%"
        ws.Cells.EntireColumn.AutoFit
        
    Next ws

End Sub
