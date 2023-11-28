Sub SortMultipleColumns()

    For Each ws In Worksheets

        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        Dim tickerSymbol As String
        Dim openPrice As Currency
        Dim closePrice As Currency
        
        Dim volTotal As Double
        Dim summaryTableRow As Integer
        summaryTableRow = 2
        
        Dim gratstIncreaseSymbol As String
        Dim gratstDecreaseSymbol As String
        Dim gratstVolumeSymbol As String
        
        Dim gratstIncrease As Double
        Dim gratstDecrease As Double
        Dim gratstVolume As Double
       
        '  Add Column titles format columns
            ws.Range("I1").Value = "Ticker"
            ws.Columns("I").ColumnWidth = 9
            
            ws.Range("J1").Value = "Yearly Change"
            ws.Columns("J").ColumnWidth = 13
            
            ws.Range("K1").Value = "Percent Change"
            ws.Columns("K").ColumnWidth = 13
            ws.Columns("K").NumberFormat = "0.00%"
            
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Columns("L").ColumnWidth = 17
            
            ws.Range("P1").Value = "Ticker"
            ws.Columns("P").ColumnWidth = 9
            
            ws.Range("Q1").Value = "Value"
            ws.Columns("Q").ColumnWidth = 13
            
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Columns("O").ColumnWidth = 21
            
            
            For i = 2 To LastRow
                    
                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                            ws.Range("J" & summaryTableRow).Value = ws.Cells(i, 3).Value
                            
                            openPrice = ws.Cells(i, 3).Value
           
                    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                            tickerSymbol = ws.Cells(i, 1).Value
                            
                            volTotal = volTotal + ws.Cells(i, 7).Value
                    
                            ws.Range("I" & summaryTableRow).Value = tickerSymbol
                    
                            ws.Range("L" & summaryTableRow).Value = volTotal
                            
                            ws.Range("J" & summaryTableRow).Value = -1 * (Range("J" & summaryTableRow).Value - Cells(i, 6).Value)
                    
                            ws.Range("K" & summaryTableRow).Value = ws.Range("J" & summaryTableRow).Value / openPrice
    
                            summaryTableRow = summaryTableRow + 1
                    
                            volTotal = 0
                    
                    Else
                    
                            volTotal = volTotal + ws.Cells(i, 7).Value
                    
                    End If
            
            Next i
            
            LastRowa = Cells(Rows.Count, 9).End(xlUp).Row
          
' hlprCode          MsgBox (Str(LastRow))

'           Color the boxes
            For i = 2 To LastRowa
            
                    If ws.Cells(i, 10).Value < 0 Then
                                
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                                    
                    ElseIf ws.Cells(i, 10).Value > 0 Then
                            
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                                                
                    End If
                    
            Next i
            
            For i = 2 To LastRowa
            
                    
                    If ws.Cells(i, 11).Value > gratstIncrease Then
                            
                        gratstIncreaseSymbol = ws.Cells(i, 9)
                        gratstIncrease = ws.Cells(i, 11)
                          
                    End If

                    If ws.Cells(i, 11).Value < gratstDecrease Then
                           
                        gratstDecreaseSymbol = ws.Cells(i, 9)
                        gratstDecrease = ws.Cells(i, 11)
                            
                    End If
     
                    If ws.Cells(i, 12).Value > gratstVolume Then
                          
                        gratstVolumeSymbol = ws.Cells(i, 9)
                        gratstVolume = ws.Cells(i, 12)
                            
                    End If
                    
            Next i
'                   Formatting

                    ws.Range("P2").Value = gratstIncreaseSymbol
                    ws.Range("Q2").Value = gratstIncrease
                    ws.Range("Q2").NumberFormat = "0.00%"

                    ws.Range("P3").Value = gratstDecreaseSymbol
                    ws.Range("Q3").Value = gratstDecrease
                    ws.Range("Q3").NumberFormat = "0.00%"
                    
                    ws.Range("P4").Value = gratstVolumeSymbol
                    ws.Range("Q4").Value = gratstVolume
                    
                    ws.Columns(10).NumberFormat = "#,##0.#0"
                    
    Next ws

End Sub
