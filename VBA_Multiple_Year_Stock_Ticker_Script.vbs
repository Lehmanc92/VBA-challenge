Sub ticker():

Dim lastRow As Long
Dim tableRow As Integer
Dim openPrice As Double
Dim volume As LongLong
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total_volume As LongLong
Dim winner As String
Dim loser As String
Dim volume_winner As String
Dim sheet As Worksheet


For Each ws In Worksheets

tableRow = 2
volume = 0
openPrice = ws.Cells(2, 3).Value
greatest_increase = 0
greatest_decrease = 0
greatest_total_volume = 0



lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Iterate through all rows


    For i = 2 To lastRow
    
    volume = volume + ws.Cells(i, 7).Value
       
        'Set conditional when ticker changes
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
             
            'Copy individual ticker symbols from column A to column I
                
            ws.Cells(tableRow, 9).Value = ws.Cells(i, 1).Value
            
            ws.Cells(tableRow, 10).Value = ws.Cells(i, 6).Value - openPrice
            
    
            
            'In case Open Price was equal to zero (avoids division by zero)
            
            If openPrice <> 0 Then
            
                ws.Cells(tableRow, 11).Value = ws.Cells(tableRow, 10).Value / openPrice
            
            End If
                        
            ws.Cells(tableRow, 11).NumberFormat = "%0.00"
            
            ws.Cells(tableRow, 12).Value = volume
            
            'Capture the greatest numbers and the companies
            
                    If volume > greatest_total_volume Then
                        greatest_total_volume = volume
                        volume_winner = ws.Cells(tableRow, 9).Value
                    End If
                    
                    If ws.Cells(tableRow, 11).Value > greatest_increase Then
                        greatest_increase = ws.Cells(tableRow, 11).Value
                        winner = ws.Cells(tableRow, 9).Value
                        
                    ElseIf ws.Cells(tableRow, 11).Value < greatest_decrease Then
                        greatest_decrease = ws.Cells(tableRow, 11).Value
                        loser = ws.Cells(tableRow, 9).Value
                    End If
            
            
            'Conditional Formatting for Yearly Change
            
            If ws.Cells(tableRow, 10).Value > 0 Then
            ws.Cells(tableRow, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(tableRow, 10).Interior.ColorIndex = 3
            End If
            
            'Setting values for the next i to proceed
            
            tableRow = tableRow + 1
            openPrice = ws.Cells(i + 1, 3)
            volume = 0
                  
            

          
        
        
        End If
    

    
    
    
    Next i
    
'Print summaries of biggest movers

    
ws.Cells(2, 16).Value = winner
ws.Cells(2, 17).Value = greatest_increase

ws.Cells(3, 16).Value = loser
ws.Cells(3, 17).Value = greatest_decrease

ws.Cells(4, 16).Value = volume_winner
ws.Cells(4, 17).Value = greatest_total_volume

ws.Range("Q2:Q3").NumberFormat = "%0.00"




Next ws


'widen columns
For Each sheet In ThisWorkbook.Worksheets
    sheet.Range("J:Q").EntireColumn.AutoFit
Next sheet

End Sub

