Sub Stock_Data()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("I1").Font.Bold = True
        ws.Range("I1").HorizontalAlignment = xlCenter
    
        ws.Range("J1").Value = "Yearly_Change"
        ws.Range("J1").Font.Bold = True
        ws.Range("J1").HorizontalAlignment = xlCenter
    
        ws.Range("K1").Value = "Percent_Change"
        ws.Range("K1").Font.Bold = True
        ws.Range("K1").HorizontalAlignment = xlCenter
    
        ws.Range("L1").Value = "Total_Stock_Volume"
        ws.Range("L1").Font.Bold = True
        ws.Range("L1").HorizontalAlignment = xlCenter

        Dim ticker As String
    
        Dim volumne_total As Long
        volume_total = 0
    
        Dim ticker_table_row As Integer
        ticker_table_row = 2
        
        Dim price_difference_open As Double
        
        Dim price_difference_open_percent As Double
        
        Dim price_difference_close As Double
        
        Dim price_difference_close_percent As Double
        
        price_difference_open = 0
        price_difference_open_percent = 0
        price_difference_close = 0
        price_difference_close_percent = 0
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastrow
            
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                price_difference_open = ws.Cells(i, 3)
                
                price_difference_open_percent = ws.Cells(i, 3)
                
                End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
            
                volume_total = volume_total + ws.Cells(i, 7).Value
            
                ws.Range("I" & ticker_table_row).Value = ticker
            
                ws.Range("L" & ticker_table_row).Value = volume_total
                
                price_difference_close = ws.Cells(i, 6)
                
                price_difference_close_percent = ws.Cells(i, 6)
            
                ws.Range("j" & ticker_table_row).Value = price_difference_close - price_difference_open
                
                ws.Range("k" & ticker_table_row).Value = ((price_difference_close_percent - price_difference_open_percent) / price_difference_open_percent) * 100
                
                    If ws.Range("K" & ticker_table_row).Value >= 0 Then
                    
                    ws.Range("K" & ticker_table_row).Interior.ColorIndex = 4
                    
                    Else: ws.Range("K" & ticker_table_row).Interior.ColorIndex = 3
                    
                    End If
                
                ticker_table_row = ticker_table_row + 1
            
                volume_total = 0
                
                price_difference_close = 0
                
                price_difference_close_percent = 0
            
            Else
            
                volume_total = volume_total + ws.Cells(i, 7).Value
            
            End If
    
        Next i
    
    Next ws
    
End Sub
