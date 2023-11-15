Sub StockTrend()


    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
For Each ws In wb.Worksheets

    
    Dim lastrow As Long
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim currentTicker As String
    Dim currentYear As Integer
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim OutputRow As Long
    Dim i As Long
    Dim RngF As Range
    
    OutputRow = 2
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestIncreaseTicker = ""
    greatestDecreaseTicker = ""
    greatestVolumeTicker = ""

    For i = 2 To lastrow
    If ws.Cells(i, 1).Value = currentTicker Then
        
        yearlyChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value

    Else
        If currentTicker <> "" Then
        ws.Cells(OutputRow, 9).Value = currentTicker
                ws.Cells(OutputRow, 10).Value = yearlyChange
                If ws.Cells(OutputRow, 3).Value <> 0 Then
                    ws.Cells(OutputRow, 11).Value = yearlyChange / ws.Cells(OutputRow, 3).Value
                Else
                    ws.Cells(OutputRow, 11).Value = 0
                End If
                ws.Cells(OutputRow, 12).Value = totalVolume

                
                If yearlyChange > 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

 ' Update summary table
 
                If percentageChange <> 0 And totalVolume <> 0 Then
                    If percentageChange > greatestIncrease Then
                        greatestIncrease = percentageChange
                        greatestIncreaseTicker = currentTicker
                    ElseIf percentageChange < greatestDecrease Then
                        greatestDecrease = percentageChange
                        greatestDecreaseTicker = currentTicker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = currentTicker
                    End If
                End If

                ' Move to the next row in the output
                OutputRow = OutputRow + 2
            
            End If
  
            
            ' Update the current ticker, year, and reset the variables
            currentTicker = ws.Cells(i, 1).Value
            currentYear = CInt(Left(ws.Cells(i, 2).Value, 4))
            yearlyChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
            totalVolume = ws.Cells(i, 7).Value


        End If
    Next i


    Set RngF = ws.Range("K2:K" & Cells(Rows.Count, 11).End(xlUp).Row)
    RngF.NumberFormat = "0.0%"
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
 
 Next
    ' Clean up
    Set ws = Nothing

End Sub

