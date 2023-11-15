Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()
   'variables
    
    Dim ws As Worksheet
    
For Each ws In Worksheets
    Dim lastRow As Double
    Dim i As Double
    Dim Ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim summary_row As Double
summary_row = 2
total_volume = 0

'For Each ws In Worksheets

'Column Headers
    
    'itiatialize variables
    'total volumen equal to zero
    'open price to the first open price
    'display ro number equal to zero
     'used in 4 below
     
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
 
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    open_price = ws.Cells(2, 3).Value

' start
' Loop through each row from 2 to the last row
    For i = 2 To lastRow
    
    
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
      Ticker = ws.Cells(i, 1).Value
      close_price = ws.Cells(i, 6).Value
      yearly_change = close_price - open_price
      percent_change = yearly_change / open_price
      

      
      total_volume = total_volume + ws.Cells(i, 7).Value

      
      ws.Range("I" & summary_row).Value = Ticker

      
      ws.Range("J" & summary_row).Value = yearly_change
      
      ws.Range("K" & summary_row).Value = percent_change
      
      ws.Range("K" & summary_row).NumberFormat = "0.00%"
      
      
      ws.Range("L" & summary_row).Value = total_volume
      
      
      If ws.Range("J" & summary_row).Value > 0 Then
      ws.Range("J" & summary_row).Interior.ColorIndex = 4
      ElseIf ws.Range("J" & summary_row).Value < 0 Then
      ws.Range("J" & summary_row).Interior.ColorIndex = 3
      End If
      
      

      
      summary_row = summary_row + 1
      
      
      total_volume = 0
      
      open_price = ws.Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      total_volume = total_volume + ws.Cells(i, 7).Value

End If
    
        
        
    
  
    Next i
    
            ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
            ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
            ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
            
            Max_Index = WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K2:K" & lastRow), 0)
            Min_Index = WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("K2:K" & lastRow), 0)
            Volume_Index = WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L2:L" & lastRow), 0)
            
            ws.Cells(2, 16).Value = ws.Cells(Max_Index + 1, 9).Value
            ws.Cells(3, 16).Value = ws.Cells(Min_Index + 1, 9).Value
            ws.Cells(4, 16).Value = ws.Cells(Volume_Index + 1, 9).Value
            
            ws.Range("Q2").NumberFormat = "0.00%"
            
            ws.Range("Q3").NumberFormat = "0.00%"
    
    
    Next ws
End Sub


