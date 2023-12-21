Attribute VB_Name = "Module1"
Sub CreateOutput():
    For Each ws In Worksheets:
        ' Delete columns
        ' Mostly for resetting if needing to be rerun
        ws.Range("I:I,J:J,K:K,L:L,N:N,O:O,P:P").Delete
        
        ' =======================================
        '               FIRST TABLE
        ' =======================================
    
        ' Variables
        Dim lastRow, outputRow As Integer
        Dim openStart, closeEnd, percentChange, stockVolume As Double
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputRow = 2
        openStart = ws.Range("C2").Value
        closeEnd = 0
        percentChange = 0
        stockVolume = 0
        
        ' Create columns and set titles
        ws.Range("I1:L1").EntireColumn.Insert
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Insert data
        For i = 2 To lastRow:
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Ticker
                ws.Cells(outputRow, 9).Value = ws.Cells(i, 1).Value
                
                ' Yearly Change
                closeEnd = ws.Cells(i, 6).Value
                ws.Cells(outputRow, 10).Value = closeEnd - openStart
                If ws.Cells(outputRow, 10).Value < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
                
                ' Percent Change
                percentChange = (closeEnd - openStart) / openStart
                ws.Cells(outputRow, 11).Value = percentChange
                If percentChange < 0 Then
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 4
                End If
                
                ' Total Stock Volume
                ws.Cells(outputRow, 12).Value = stockVolume
                
                ' Reset values
                openStart = ws.Cells(i + 1, 3).Value
                stockVolume = 0
                
                outputRow = outputRow + 1
            End If
        Next i
        
        ' =======================================
        '               SECOND TABLE
        ' =======================================
        
        ' Variables
        Dim table_lastRow, increase_row, decrease_row, volume_row As Integer
        Dim increase, decrease, volume As Double
        
        table_lastRow = ws.Range("I" & ws.Rows.Count).End(xlUp).Row
        increase = WorksheetFunction.Max(ws.Range("K2:K" & table_lastRow))
        decrease = WorksheetFunction.Min(ws.Range("K2:K" & table_lastRow))
        volume = WorksheetFunction.Max(ws.Range("L2:L" & table_lastRow))
        
        ' Get row numbers of increase, decrease, and volume
        increase_row = ws.Range("K2:K" & table_lastRow).Find(what:=increase).Row
        decrease_row = ws.Range("K2:K" & table_lastRow).Find(what:=decrease).Row
        volume_row = ws.Range("L2:L" & table_lastRow).Find(what:=volume).Row
        
        ' Create columns and set titles
        ws.Range("N1:P1").EntireColumn.Insert
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ' Set appropriate values in table
        ws.Range("P2").Value = increase
        ws.Range("P3").Value = decrease
        ws.Range("P4").Value = volume
        
        ws.Range("O2").Value = ws.Range("I" & increase_row).Value
        ws.Range("O3").Value = ws.Range("I" & decrease_row).Value
        ws.Range("O4").Value = ws.Range("I" & volume_row).Value
        
        
        ' Format columns/cells
        ws.Range("J2:J" & table_lastRow).NumberFormat = "0.00"
        ws.Range("K2:K" & table_lastRow).NumberFormat = "0.00%"
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "0.00E+0"
        ws.Columns("I:L").AutoFit
        ws.Columns("N:P").AutoFit
    Next ws
End Sub

