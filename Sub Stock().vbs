Sub Stock()

For Each ws In Worksheets

Dim Worksheetname As String
Dim Column As Integer
Dim i As Long, j As Long, k As Long
Dim Volume As Double, Yearly_Change As Double
Dim Ticker As String
Dim Open_year As Double, Close_year As Double
Dim Percent_Change As Double, Previous_i As Double
Dim Previous_k As Integer


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Columns("I:K").ColumnWidth = 15
ws.Columns("L").ColumnWidth = 20

Yearly_Change = 0
Volume = 0
Ticker_v = 2
Previous_i = 2

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                  Ticker = ws.Cells(i, 1).Value
                  
                  Volumne = Volume + ws.Cells(i, 7).Value
                           
                  Close_year = ws.Cells(i, 6).Value
                  Open_year = ws.Cells(Previous_i, 3).Value
                  
                  If Open_year = 0 Then
                  Perecent_Change = Close_year
                  Else
                  Yearly_Change = Close_year - Open_year
                  Percent_Change = Yearly_Change / Open_year
                  End If
                  
                  ws.Range("I" & Ticker_v).Value = Ticker
                  ws.Range("L" & Ticker_v).Value = Volume
                  ws.Range("J" & Ticker_v).Value = Yearly_Change
                  ws.Range("K" & Ticker_v).Value = Percent_Change
                  ws.Range("K" & Ticker_v).NumberFormat = "0.00%"
                 
                  
                  
                  Ticker_v = Ticker_v + 1
                  Volume = 0
                  Yearly_Change = 0
                  Percent_Change = 0
                  Previous_i = i + 1
                  
              
Else
     
     Volume = Volume + ws.Cells(i, 7).Value

End If

Next i

lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To lastrow2

                If ws.Cells(j, 10).Value > 0 Then

               ws.Cells(j, 10).Interior.ColorIndex = 4
               
               Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
               
             
End If
Next j
                  
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Columns("O").ColumnWidth = 20

lastrow3 = ws.Cells(Rows.Count, "K").End(xlUp).Row

Increase = 0
Decrease = 0
Greatest = 0

For k = 3 To lastrow3

                Previous_k = k - 1
                
                Current_Change = ws.Cells(k, 11).Value
                Previous_Change = ws.Cells(Previous_k, 11).Value
                
                Volume = ws.Cells(k, 12).Value
                Previous_Vol = ws.Cells(Previous_k, 12).Value
                
If Increase > Current_Change And Increase > Previous_Change Then
        
                Increase = Increase
        
        ElseIf Current_Change > Increase And Current_Change > Previous_Change Then
                
                Increase = Current_Change
                Increase_name = ws.Cells(k, 9).Value
                
        ElseIf Previous_Change > Increase And Previous_Change > Current_Change Then
        
                Increase = Previous_Change
                Increase_name = ws.Cells(Previous_k, 9).Value
End If
               
If Decrease < Current_Change And Decrease < Previous_Change Then
        
                Decrease = Decrease
        
        ElseIf Current_Change < Increase And Current_Change < Previous_Change Then
                
                Decrease = Current_Change
                Decrease_name = ws.Cells(k, 9).Value
                
        ElseIf Previous_Change < Increase And Previous_Change < Current_Change Then
        
                Decrease = Previous_Change
                Decrease_name = ws.Cells(Previous_k, 9).Value
End If
                              
If Greatest > Volume And Greatest > Previous_Vol Then
        
                Greatest = Greatest
        
        ElseIf Volume > Increase And Volume > Previous_Vol Then
                
                Greatest = Volume
                Volume_name = ws.Cells(k, 9).Value
                
        ElseIf Previous_Vol > Increase And Previous_Vol > Current_Change Then
        
                Greatest = Previous_Vol
                Volume_name = ws.Cells(Previous_k, 9).Value
End If
Next k

ws.Cells(2, 16).Value = Increase_name
ws.Cells(2, 17).Value = Increase
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = Decrease_name
ws.Cells(3, 17).Value = Decrease
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = Volume_name
ws.Cells(4, 17).Value = Greatest

Next ws


End Sub