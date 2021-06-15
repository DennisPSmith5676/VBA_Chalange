Attribute VB_Name = "Ticker"
Sub Ticker():
'Set Dimentions
    Dim Ticker As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Total_Volume As Double
    Dim i As Long
    Dim v As Long
    Dim Sum_Table As Long
    Dim Pct As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim PCTCHG As String
    
    
    For Each ws In Worksheets
            
            'Set Row Labels
            ws.Range("J1").Value = "Ticker"
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L1").Value = "Percent Change"
            ws.Range("M1").Value = "Total Stock Volume"
            
            'Set Summary table count to 2
            Sum_Table = 2
            
            'Print Opening Price
            Opening_Price = ws.Cells(2, 3).Value
                    
            'set the count
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
            'Loop thru
            For i = 2 To LastRow
                             
                'Print Volume rows in Summary
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                 
                'Start If then loop
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    Sum_Table = Sum_Table + 1
                    
                    'Define Ticker
                    Ticker = ws.Cells(i, 1).Value
            
                    'Print Ticker & Total_Volume rows in the summary table
                    ws.Range("J" & Sum_Table).Value = Ticker
                    ws.Range("M" & Sum_Table).Value = Total_Volume
                    
                    'Print closing price in summary table
                    Closing_Price = ws.Cells(i, 6).Value
                   
                   'Calc yearly Price Change
                    ws.Range("K" & Sum_Table).Value = Closing_Price - Opening_Price
              
                    'Format % Change
                    ws.Range("L" & Sum_Table).NumberFormat = "0.00%"
                    
                    'Skiping if divide by zero
                    If Opening_Price = 0 Then
                        ws.Range("L" & Sum_Table).Value = 0
                    Else
                        ws.Range("L" & Sum_Table).Value = ws.Range("K" & Sum_Table).Value / Opening_Price
                                                
                    End If
                     
                   
                    
                    'Reset Total_Volume
                    Total_Volume = 0
                    
                    'Set Opening Price
                    Opening_Price = ws.Cells(i + 1, 3).Value
                End If
            
            Next i
            
        'Bonus steps set rows
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    
    
        'Looking for the > Increase
        For i = 2 To LastRow
                If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Cells(i, 11).Value
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                End If
            
            'Looking for the > Decrease
            If ws.Cells(i, 12).Value > 0 Then
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                End If
            
            'Looking for the > Total Volume
            If ws.Cells(i, 11).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                End If
        
            'Format > , < and Volume
                ws.Range("Q4").NumberFormat = "0.00"
                ws.Range("Q2").NumberFormat = "0.00"
                ws.Range("Q3").NumberFormat = "0"
            
            'Conditional Formatting
            If ws.Cells(i, 12).Value >= 0 Then
                 ws.Cells(i, 12).Interior.ColorIndex = 4
            Else
                 ws.Cells(i, 12).Interior.ColorIndex = 3
            End If
        
        Next i
         
    Next ws
    
End Sub
      




