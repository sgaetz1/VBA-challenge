Attribute VB_Name = "Module1"
Sub Stock_market_bonus()
    
    'loop through all the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        
        'put headings on new table columns
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        'declare and initialize row number
        Dim Table_row As Integer
        Table_row = 2
        
        'declare and set the opening stock price to the first opening price
        Dim Open_price As Double
        Open_price = Cells(2, 3).Value
        
        'declare and initialize the total volume to zero
        Dim Total_volume As Double
        Total_volume = 0
        
        'declare and initialize the greatest total volume
        Dim Greatest_total_volume As Double
        Greatest_total_volume = 0
        
        'declare and initialize variables to hold the greatest increase and decrease
        Dim increase As Double
        Dim decrease As Double
        increase = 0
        decrease = 0
        
        'find the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'loop through all the rows
        Dim i As Double
        For i = 2 To LastRow
        
            'keep a running total of volume
            Total_volume = Total_volume + ws.Cells(i, 7).Value
            
            'find where the stock names don't match
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'set the stock name
                Dim Stock_name As String
                Stock_name = ws.Cells(i, 1).Value
                
                'put the stock name in the table
                ws.Range("I" & Table_row).Value = Stock_name
                
                'find the closing price
                Dim Close_price As Double
                Close_price = ws.Cells(i, 6).Value
                
                'calculate the change from opening to close
                Change_price = Close_price - Open_price
                
                'put the change in price in the table
                ws.Range("J" & Table_row).Value = Change_price
                
                    'add color green to positive change and red to negative change
                    If Change_price >= 0 Then
                    
                        ws.Range("J" & Table_row).Interior.ColorIndex = 4
                    
                    Else
                    
                        ws.Range("J" & Table_row).Interior.ColorIndex = 3
                        
                    End If
                    
                'calculat the percent change, set it to zero if the open and close are zeroes
                Dim Percent_change As Double
                
                If Open_price <> 0 Then
                
                    Percent_change = Change_price / Open_price
                    
                Else
                
                    Percent_change = 0
                    
                End If
                
                'put percent change in the table and format it to a percent
                ws.Range("K" & Table_row).Value = FormatPercent(Percent_change, 2)
                    
                'put total volume in the table
                ws.Range("L" & Table_row).Value = Total_volume
                
                'reset the opening price for the next stock
                Open_price = ws.Cells(i + 1, 3).Value
                
                'move down a row in the table
                Table_row = Table_row + 1
                
                'calculate greatest total volume
                If Total_volume > Greatest_total_volume Then
                    
                    Greatest_total_volume = Total_volume
                    Greatest_stock = Stock_name
                    
                End If
                
                'reset total volume to zero for the next stock
                Total_volume = 0
                
                
                'calculate greatest/least % change
                If Percent_change > increase Then
                
                    increase = Percent_change
                    Dim ticker As String
                    ticker = Stock_name
                    
                ElseIf Percent_change < decrease Then
                
                    decrease = Percent_change
                    Dim ticker2 As String
                    ticker2 = Stock_name
                
                End If
                
                
            End If
            
            
        
        Next i
        
        'put greatest increase/decrease in table, autofit all new columns
        ws.Range("P2").Value = ticker
        ws.Range("P3").Value = ticker2
        ws.Range("P4").Value = Greatest_stock
        ws.Range("Q2").Value = FormatPercent(increase, 2)
        ws.Range("Q3").Value = FormatPercent(decrease, 2)
        ws.Range("Q4").Value = Greatest_total_volume
        ws.Columns("I:Q").AutoFit

    'Debug.Print (LastRow)
    
    Next ws
    
End Sub

