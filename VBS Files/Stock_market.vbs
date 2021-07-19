Attribute VB_Name = "Module2"
Sub Stock_market()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        Dim Stock_name As String
        Dim i As Double
        Dim Table_row As Integer
        Dim Col_heading As String
        
        Table_row = 2
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        
        Dim Open_price As Double
        Open_price = Cells(2, 3).Value
        
        Dim Total_volume As String
                
        Total_volume = 0
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
            Total_volume = Total_volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Stock_name = ws.Cells(i, 1).Value
                
                ws.Range("I" & Table_row).Value = Stock_name
                
                Dim Close_price As Double
                
                Close_price = ws.Cells(i, 6).Value
                
                Change_price = Close_price - Open_price
                
                ws.Range("J" & Table_row).Value = Change_price
                
                    If Change_price >= 0 Then
                    
                        ws.Range("J" & Table_row).Interior.ColorIndex = 4
                    
                    Else
                    
                        ws.Range("J" & Table_row).Interior.ColorIndex = 3
                        
                    End If
                    
                Dim Percent_change As Double
                
                If Open_price <> 0 & Close_price <> 0 Then
                
                    Percent_change = Change_price / Open_price
                    
                Else
                
                    Percent_change = 0
                    
                End If
                
                ws.Range("K" & Table_row).Value = FormatPercent(Percent_change, 2)
                    
                ws.Range("L" & Table_row).Value = Total_volume
                
                Open_price = ws.Cells(i + 1, 3).Value
               
                Table_row = Table_row + 1
                
                Total_volume = 0
                
                
                
            End If
            
            
        
        Next i
        
    Debug.Print (LastRow)
    
    Next ws
    
End Sub
