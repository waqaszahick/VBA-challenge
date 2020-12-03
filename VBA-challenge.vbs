Sub TickerSymbol()

   
    ActiveSheet.Select
    Application.ScreenUpdating = False
    
    Range("i:p").ClearContents
    Columns("j:j").Interior.Pattern = xlNone

    
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    Range("n3").Value = "Greatest % Increase"
    Range("n4").Value = "Greatest % Decrease"
    Range("n5").Value = "Greatest Total Volume"
    Range("o2").Value = "Ticker"
    Range("p2").Value = "Value"
    
    Range("i1:l1, n3:n5, o2:p2").Font.Bold = True
    Range("i1:l1").HorizontalAlignment = xlCenter
    Range("o2:p2").HorizontalAlignment = xlCenter

            
    ticker_last_row = Cells(Rows.Count, 1).End(xlUp).Row

        
    volume = 0
    j = 2
        
    For i = 2 To ticker_last_row
        Ticker = Cells(i, 1).Value
        volume = volume + Cells(i, 7).Value
        
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
           opening_price = Cells(i, 3).Value
        End If
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
           closing_price = Cells(i, 6).Value
           yearly_change = closing_price - opening_price
           
           If yearly_change = 0 Or opening_price = 0 Then
              percent_change = FormatPercent(0)
           Else
                percent_change = FormatPercent(yearly_change / opening_price)
           End If
           
           Cells(j, 9).Value = Ticker
           Cells(j, 10).Value = yearly_change
            
           If yearly_change < 0 Then
                Cells(j, 10).Interior.Color = vbRed
           Else
                Cells(j, 10).Interior.Color = vbGreen
           End If
                        
           Cells(j, 11).Value = percent_change
           Cells(j, 12).Value = volume
            
           j = j + 1
           volume = 0
        End If
    Next i


    greatest_percent_increase = FormatPercent(WorksheetFunction.Max(Range("k:k")))
    greatest_percent_decrease = FormatPercent(WorksheetFunction.Min(Range("k:k")))
    greatest_total_volume = WorksheetFunction.Max(Range("l:l"))
            
    Range("p3").Value = greatest_percent_increase
    Range("p4").Value = greatest_percent_decrease
    Range("p5").Value = greatest_total_volume
    
    greatest_percent_increase_ticker = Range("i" & WorksheetFunction.Match(Range("p3").Value, Range("k:k"), 0)).Value
    greatest_percent_decrease_ticker = Range("i" & WorksheetFunction.Match(Range("p4").Value, Range("k:k"), 0)).Value
    greatest_total_volume_ticker = Range("i" & WorksheetFunction.Match(Range("p5").Value, Range("l:l"), 0)).Value
   
    Range("o3").Value = greatest_percent_increase_ticker
    Range("o4").Value = greatest_percent_decrease_ticker
    Range("o5").Value = greatest_total_volume_ticker
    
    Columns("i:p").Columns.AutoFit

Application.CutCopyMode = False
    
Application.ScreenUpdating = True

End Sub