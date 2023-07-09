
Sub Stonks():
For Each ws In Worksheets
    Dim lastrow As Double
    Dim stockname As String
    Dim summarytablerow As Double
    Dim BeginningYearPrice As Double
    Dim EndYearPrice As Double
    Dim PriceDiff As Double
    Dim totalvolume As Double
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    

    totalvolume = 0
    summarytablerow = 2
    BeginningYearPrice = ws.Cells(2, 3).Value
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            stockname = ws.Cells(i, 1).Value
            ws.Cells(summarytablerow, 9).Value = stockname
            
            
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            ws.Cells(summarytablerow, 12).Value = totalvolume
            
            EndYearPrice = ws.Cells(i, 6).Value
            ws.Cells(summarytablerow, 10).Value = EndYearPrice - BeginningYearPrice
            ws.Cells(summarytablerow, 10).Value = Format(ws.Cells(summarytablerow, 10), "0.00")
            If (EndYearPrice > BeginningYearPrice) Then
                ws.Cells(summarytablerow, 10).Interior.ColorIndex = 4
            ElseIf (EndYearPrice < BeginningYearPrice) Then
                ws.Cells(summarytablerow, 10).Interior.ColorIndex = 3
            End If
            
            ws.Cells(summarytablerow, 11).Value = FormatPercent(ws.Cells(summarytablerow, 10).Value / BeginningYearPrice, 2)
            
            summarytablerow = summarytablerow + 1
            BeginningYearPrice = ws.Cells(i + 1, 3).Value
            totalvolume = 0
        Else
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If
    Next i

    ws.Cells(2, 17).Value = ws.Cells(2, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(2, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(2, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(2, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(2, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(2, 9).Value
    
    
    
    finalrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For j = 2 To finalrow
        If (ws.Cells(j, 11).Value > ws.Cells(2, 17).Value) Then
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(j, 11).Value
        End If
        If (ws.Cells(j, 11).Value < ws.Cells(3, 17).Value) Then
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(j, 11).Value
        End If
        If (ws.Cells(j, 12).Value > ws.Cells(4, 17).Value) Then
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
        End If
    Next j
    
    ws.Cells(2, 17).Value = FormatPercent(ws.Cells(2, 17).Value, 2)
    ws.Cells(3, 17).Value = FormatPercent(ws.Cells(3, 17).Value, 2)
    
Next ws
End Sub

