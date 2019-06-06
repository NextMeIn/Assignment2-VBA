Attribute VB_Name = "Module1"
Sub processhardway()
    'Dishan Wickremasinghe

    Dim lastRow As Long
    Dim totVol As Double
    Dim x, y As Integer
    Dim openPrice, closePrice As Double
    Dim oldTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
        With ws
        
            ws.Range("I:L").ClearContents
            ws.Range("I:L").ClearFormats
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
            .Range("I1").Value = Mid(.Range("A1").Value, 2, Len(.Cells(1, 1)) - 2)
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            
            x = 2
            y = 2
            oldTicker = .Cells(x, 1).Value
            openPrice = .Cells(x, 3).Value
            closePrice = .Cells(x, 6).Value
            .Range("Q2").Value = 0
            .Range("Q3").Value = 0
            .Range("Q4").Value = 0
            
            For x = 2 To lastRow
                If .Cells(x, 1).Value = oldTicker Then
                    totVol = IIf(.Cells(x, 7).Value > 0, totVol + .Cells(x, 7).Value, totVol)
                    closePrice = .Cells(x, 6).Value
                Else
                    .Cells(y, 9).Value = oldTicker
                    .Cells(y, 10).Value = closePrice - openPrice
                    .Cells(y, 11).Value = (closePrice - openPrice) / IIf(openPrice > 0, openPrice, 1)
                    
                    .Range("Q2").Value = IIf(.Range("Q2").Value > .Cells(y, 11).Value, .Range("Q2").Value, .Cells(y, 11).Value)
                    .Range("P2").Value = IIf(.Range("Q2").Value > .Cells(y, 11).Value, .Range("P2").Value, .Cells(y, 9).Value)
                    
                    .Range("Q3").Value = IIf(.Range("Q3").Value < .Cells(y, 11).Value, .Range("Q3").Value, .Cells(y, 11).Value)
                    .Range("P3").Value = IIf(.Range("Q3").Value < .Cells(y, 11).Value, .Range("P3").Value, .Cells(y, 9).Value)
                    
                    .Cells(y, 10).Interior.ColorIndex = IIf(.Cells(y, 10).Value >= 0, 4, 3)
                    .Cells(y, 12).Value = totVol
                    .Range("Q4").Value = IIf(.Range("Q4").Value > .Cells(y, 12).Value, .Range("Q4").Value, .Cells(y, 12).Value)
                    .Range("P4").Value = IIf(.Range("Q4").Value > .Cells(y, 12).Value, .Range("P4").Value, .Cells(y, 9).Value)
                    
                    y = y + 1
                    oldTicker = .Cells(x, 1).Value
                    totVol = .Cells(x, 7).Value
                    openPrice = .Cells(x, 3).Value
                    closePrice = .Cells(x, 6).Value
                End If
            Next x
        End With
        ws.Columns("I:Q").AutoFit
        ws.Columns("K").NumberFormat = "0.00%"
    Next ws
    
End Sub
