Attribute VB_Name = "Module1"
Sub StockAnalysis():


Dim total As Double
Dim i As Long
Dim change As Double
Dim percentChange As Double
Dim start As Long
Dim lastRow As Long
Dim j As Long


Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Ttotal Stock Volume"

j = 0
total = 0
change = 0
start = 2

    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        total = total + Cells(i, 7).Value
        
        If total = 0 Then
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("J" & 2 + j).Value = 0
        Range("K" & 2 + j).Value = "%" & 0
        Range("L" & 2 + j).Value = 0
        
        Else
        If Cells(start, 3) = 0 Then
        For find_value = start To i
            If Cells(find_value, 3).Value <> 0 Then
                start = find_value
                Exit For
            End If
        Next find_value
        End If
        
        change = (Cells(i, 6) - Cells(start, 3))
        percentChange = change / Cells(start, 3)
        
        start = i + 1
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("J" & 2 + j).Value = change
        Range("J" & 2 + j).NumberFormat = "0.00"
        Range("K" & 2 + j).Value = percentChange
        Range("K" & 2 + j).NumberFormat = "0.00%"
        Range("L" & 2 + j).Value = total
        
        Select Case change
            Case Is > 0
                Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
                Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
        
        End If
        
        total = 0
        change = 0
        j = j + 1
        
        Else
        total = total + Cells(i, 7).Value
        
        End If
        Next i

End Sub
