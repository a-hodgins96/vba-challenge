# vba-challenge
Sub tickerloop()

    Dim tickername As String
    Dim tickervolume As Double
    tickervolume = 0
    
    Dim summary_ticker_row As Double
    summary_ticker_row = 2
    
    
    
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
 lastrow = 22771
    
    For i = 2 To lastrow
    Dim open_price As Double
    open_price = Cells(i, 3).Value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        tickername = Cells(i, 1).Value
        
        tickervolume = tickervolume + Cells(i, 7).Value
        
        Range("I" & summary_ticker_row).Value = tickername
        
        Range("L" & summary_ticker_row).Value = tickervolume
        
        close_price = Cells(i, 6).Value
        
        yearly_change = Cells(i, 3) - Cells(i, 6)

        
        If (open_price = 0) Then
            percent_change = 0
            
        Else
        
            percent_change = yearly_change / open_price
            
        End If
        
        
    Range("K" & summary_ticker_row).Value = percent_change
    Range("K" & summary_ticker_row).NumberFormat = "0.00%"
    
    Range("J" & summary_ticker_row).Value = percent_change
    
    summary_ticker_row = summary_ticker_row + 1
    
    tickervolume = 0
    
    open_price = Cells(i + 1, 3)
    
    Else
        
        tickervolume = tickervolume + Cells(i, 7).Value
        
        End If
        
    Next i
    
      lastrow_summary_table = 91

    For i = 2 To lastrow_summary_table
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 10
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    Next i
    
 
End Sub
