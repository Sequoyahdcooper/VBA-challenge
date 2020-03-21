VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Multiple_year_stock_data()

For Each WS In Worksheets
    
        MsgBox ("Looping for Year " & WS.Name)
    
        Dim lastrow As Double
        lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim ticker As String
        
        Dim total_volume As Double
        total_volume = 0
        
        Dim yearly_change As Double
        
        Dim percentage_change As Double
        
        Dim summary_row As Integer
        summary_row = 2
        
        Dim closing_price As Double
        
        Dim opening_price As Double
        
        WS.Cells(1, 9) = "Ticker"
        WS.Cells(1, 9).Font.Bold = True
        WS.Cells(1, 10) = "Yearly Change"
        WS.Cells(1, 10).Font.Bold = True
        WS.Cells(1, 11) = "Percentage Change"
        WS.Cells(1, 11).Font.Bold = True
        WS.Cells(1, 12) = "Total Volume"
        WS.Cells(1, 12).Font.Bold = True
         
        For I = 2 To lastrow
        
        '  Calculating the total volume
            
            total_volume = total_volume + WS.Cells(I, 7)
            
            If WS.Cells(I, 1) <> WS.Cells(I - 1, 1) Then
            opening_price = WS.Cells(I, 3)
            
            ElseIf WS.Cells(I, 1) <> WS.Cells(I + 1, 1) Then
            ticker = WS.Cells(I, 1)
            closing_price = WS.Cells(I, 6)
            yearly_change = (closing_price - opening_price)
                        
                If opening_price = 0 Then
                percentage_change = 0
                
                Else
                percentage_change = yearly_change / opening_price
                
                
                End If
              ' Returning tickers along with total_volume 'Calculating the yearly change and percentage change
              
            
            WS.Cells(summary_row, 9) = ticker
            WS.Cells(summary_row, 10) = yearly_change
            WS.Cells(summary_row, 11) = percentage_change
            WS.Cells(summary_row, 12) = total_volume
                    
            total_volume = 0
            summary_row = summary_row + 1
            opening_price = 0
            closing_price = 0
                        
            End If
        
        Next I
        
        WS.Range("K:K").NumberFormat = "0.00%"

    
        lastrow_new = WS.Cells(Rows.Count, 9).End(xlUp).Row
    
    
            For I = 2 To lastrow_new
    
    'if yearly change > 0 then green and if less than zero then red
                
                If WS.Cells(I, 11) > 0 Then
                WS.Cells(I, 11).Interior.ColorIndex = 4
                Else
                WS.Cells(I, 11).Interior.ColorIndex = 3
        
            End If
    
            Next I
'3 variables of ticker to return 3 different tickers
        
        Dim ticker1, ticker2, ticker3 As String
        
        Dim greatest_value As Double
        greatest_value = 0
        
        Dim lowest_value As Double
        lowest_value = 0
        
        Dim value As Double
        value = 0
        
        Dim summary_row2 As Integer
        summary_row2 = 2
        
        WS.Cells(1, 15) = "Ticker"
        WS.Cells(1, 15).Font.Bold = True
        WS.Cells(1, 16) = "Value"
        WS.Cells(1, 16).Font.Bold = True
        WS.Cells(2, 14) = "Greatest % Increase"
        WS.Cells(2, 14).Font.Bold = True
        WS.Cells(3, 14) = "Greatest % Decrease"
        WS.Cells(3, 14).Font.Bold = True
        WS.Cells(4, 14) = "Greatest Total Volume"
        WS.Cells(4, 14).Font.Bold = True
        
                For I = 2 To lastrow_new
    
    ' Calculating the greatest_value, lowest_value
                    
                    If WS.Cells(I, 11) > WS.Cells(I + 1, 11) And WS.Cells(I, 11) > greatest_value Then
                    greatest_value = WS.Cells(I, 11)
                    ticker1 = WS.Cells(I, 9)
                
                    ElseIf WS.Cells(I, 11) < WS.Cells(I + 1, 11) And WS.Cells(I, 11) < lowest_value Then
                    lowest_value = WS.Cells(I, 11)
                    ticker2 = WS.Cells(I, 9)
                
                    ElseIf WS.Cells(I, 12) > WS.Cells(I + 1, 12) And WS.Cells(I, 12) > value Then
                    value = WS.Cells(I, 12)
                    ticker3 = WS.Cells(I, 9)
                
                    End If
            
                    Next I
        
        WS.Cells(2, 15) = ticker1
        WS.Cells(3, 15) = ticker2
        WS.Cells(4, 15) = ticker3
        WS.Cells(2, 16) = greatest_value
        WS.Cells(2, 16).NumberFormat = "0.00%"
        WS.Cells(3, 16) = lowest_value
        WS.Cells(3, 16).NumberFormat = "0.00%"
        WS.Cells(4, 16) = value
            
    Next WS
        
End Sub
