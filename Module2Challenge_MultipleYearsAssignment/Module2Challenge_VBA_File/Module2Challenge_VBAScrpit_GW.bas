Attribute VB_Name = "Module2Challenge"
Sub Module2Challenge()
            
    For Each ws In Worksheets
        ws.Activate
        ws.Cells(1, "I") = "Ticker"
        ws.Cells(1, "J") = "Yearly Change"
        ws.Cells(1, "K") = "Percent Change"
        ws.Cells(1, "L") = "Total Stock Volume"
         'This is for appropriate adjustments
        ws.Cells(1, "P") = "Ticker"
        ws.Cells(1, "Q") = "Value"
        ws.Cells(2, "O") = "Greatest % Increase"
        ws.Cells(3, "O") = "Greatest % Decrease"
        ws.Cells(4, "O") = "Greatest Total Volume"
        ws.Range("I1 : L1").Font.Bold = True
  
  
        Totalvol = 0
        openprice_pointer = 2
        summery_pointer = 2
    
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To RowCount
            
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
                Totalvol = Totalvol + ws.Cells(i, "G").Value
                openprice = ws.Cells(openprice_pointer, "C").Value
                closeprice = ws.Cells(i, "F").Value
            
                ws.Cells(summery_pointer, "I").Value = ws.Cells(i, "A").Value
                ws.Cells(summery_pointer, "J").Value = closeprice - openprice
                ws.Cells(summery_pointer, "K").Value = (closeprice - openprice) / openprice
                ws.Cells(summery_pointer, "L").Value = Totalvol
                
                    If (closeprice - openprice) < 0 Then
                    ws.Cells(summery_pointer, "J").Interior.Color = RGB(255, 0, 0)
                    
                Else
                    ws.Cells(summery_pointer, "J").Interior.Color = RGB(0, 255, 0)
                
                End If
                
                Totalvol = 0
                openprice_pointer = i + 1
                summery_pointer = summery_pointer + 1
                
            Else
                
                Totalvol = Totalvol + ws.Cells(i, "G").Value
                
            End If
             
        Next i
        
                    GreatestInc = Application.WorksheetFunction.Max(Range("K2:K" & RowCount))
                    increase_ticker = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
                    Range("P2") = Cells(increase_ticker + 1, "I")
                    Range("Q2") = GreatestInc
                    
                    GreatestDec = Application.WorksheetFunction.Min(Range("K2:K" & RowCount))
                    decrease_ticker = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
                    Range("P3") = Cells(decrease_ticker + 1, "I")
                    Range("Q3") = GreatestDec
                    
                    GreatestTotVol = Application.WorksheetFunction.Max(Range("L2:L" & RowCount))
                    TotVol_ticker = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
                    Range("P4") = Cells(TotVol_ticker + 1, "I")
                    Range("Q4") = GreatestTotVol
                    
        'Exit Sub
  Next ws
  
                    MsgBox ("Done")
  
End Sub

