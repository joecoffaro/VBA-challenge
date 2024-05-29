Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim percentChange As Double
    Dim start As Long
    Dim lastRow As Long
    Dim j As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    ' Initialize variables
    j = 0
    total = 0
    change = 0
    start = 2
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Set headers for output columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set headers for summary results
    ws.Range("O1").Value = "Criteria"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            total = total + ws.Cells(i, 7).Value
            
            If total = 0 Then
                ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            Else
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
                change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                If ws.Cells(start, 3).Value <> 0 Then
                    percentChange = change / ws.Cells(start, 3)
                Else
                    percentChange = 0
                End If
                
                ' Update the start point for the next ticker
                start = i + 1
                
                ' Write results to the sheet
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = change
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = total
                
                ' Color formatting
                Select Case change
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
                
                ' Track the greatest percent increase
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ws.Cells(i, 1).Value
                End If
                
                ' Track the greatest percent decrease
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ws.Cells(i, 1).Value
                End If
                
                ' Track the greatest total volume
                If total > greatestVolume Then
                    greatestVolume = total
                    greatestVolumeTicker = ws.Cells(i, 1).Value
                End If
            End If
            
            ' Reset variables for next ticker
            total = 0
            change = 0
            j = j + 1
        Else
            total = total + ws.Cells(i, 7).Value
        End If
    Next i
    
    ' Write summary results
    ws.Range("P2").Value = greatestIncreaseTicker
    ws.Range("P3").Value = greatestDecreaseTicker
    ws.Range("P4").Value = greatestVolumeTicker
    ws.Range("Q2").Value = Format(greatestIncrease, "0.00%")
    ws.Range("Q3").Value = Format(greatestDecrease, "0.00%")
    ws.Range("Q4").Value = greatestVolume
    
        Next ws

End Sub


