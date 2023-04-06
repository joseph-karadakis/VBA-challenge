Attribute VB_Name = "Module1"
Sub StockSummary()
    Dim ws As Worksheet
    Dim ticker As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    Dim max_Increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    
    'worksheet loop
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
            'Create column headers for new tables
            
            Range("I1").Value = " Ticker"
            Range("J1").Value = " Yearly Change"
            Range("K1").Value = " Percent Change"
            Range("L1").Value = " Total Volume"
            Range("O2").Value = " Greatest % Increase"
            Range("O3").Value = " Greatest % Decrease"
            Range("O4").Value = " Greatest Total Volume"
            Range("P1").Value = " Ticker"
            Range("Q1").Value = " Value"
            
             'Format columns and cells of new tables
            
            Columns("K:K").NumberFormat = " 0.00%"
            Range("Q2").NumberFormat = " 0.00%"
            Range("Q3").NumberFormat = " 0.00%"
            
                'initialize variables
                
                max_Increase = 0
                max_decrease = 0
                max_volume = 0
                
                Dim summary_table As Range
                Dim row_index As Long
                
                Set summary_table = Range("I1:L1")
                row_index = 2
                
                'Loop through all rows of data
                For i = 2 To Range("A1").End(xlDown).Row
                
                    'Check if we're still on the same ticker, if not, record the new ticker
                    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                        ticker = Cells(i, 1).Value
                        opening_price = Cells(i, 3).Value
                    End If
                    
                    'Add to the total volume for the current ticker
                    total_volume = total_volume + Cells(i, 7).Value
                    
                    'Check if we've reached the last row of the current ticker, if yes, record the closing price and write to the summary table
                    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                        closing_price = Cells(i, 6).Value
                        yearly_change = closing_price - opening_price
                        
                        If opening_price <> 0 Then
                            percent_change = yearly_change / opening_price
                        Else
                            percent_change = 0
                        End If
                    
                    'find max % and volume
                        If percent_change > max_Increase Then
                            max_Increase = percent_change
                            max_increase_ticker = ticker
                        ElseIf percent_change < max_decrease Then
                            max_decrease = percent_change
                            max_decrease_ticker = ticker
                        End If
                        
                        If total_volume > max_volume Then
                            max_volume = total_volume
                            max_volume_ticker = ticker
                        End If
                        
                        'Write the summary data to the summary table
                        summary_table.Cells(row_index, 1).Value = ticker
                        summary_table.Cells(row_index, 2).Value = yearly_change
                        'conditional formating for Yearly Change
                            If yearly_change >= 0 Then
                            summary_table.Cells(row_index, 2).Interior.Color = RGB(0, 255, 0)
                            Else
                            summary_table.Cells(row_index, 2).Interior.Color = RGB(255, 0, 0)
                            End If
                        summary_table.Cells(row_index, 3).Value = percent_change
                        summary_table.Cells(row_index, 4).Value = total_volume
                        
                        'write the % data to the % table
                        Range("P2").Value = max_increase_ticker
                        Range("Q2").Value = max_Increase
                        Range("P3").Value = max_decrease_ticker
                        Range("Q3").Value = max_decrease
                        Range("P4").Value = max_volume_ticker
                        Range("Q4").Value = max_volume
                        
                        'Reset the variables for the next ticker
                        total_volume = 0
                        row_index = row_index + 1
                    End If
                Next i
        Next ws
End Sub


