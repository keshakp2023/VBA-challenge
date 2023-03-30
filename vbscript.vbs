Attribute VB_Name = "Module1"
Sub StockAnalysisReport()

'Variable Declaration
Dim row As Long, rowCount As Long, nextRow As Long
Dim totalStockVolume As Double, openPrice As Double, closePrice As Double
Dim ws As Worksheet

Set ws = ActiveSheet

For Each ws In Worksheets
' Assign initial values to variables
totalStockVolume = 0
nextRow = 2
greatestIncreaseTicker = " "
greatestIncrease = 0
greatestDecreaseTicker = " "
greatestDecrease = 0
greatestTotalTicker = " "
greatestTotalVolume = 0

' Column headers
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

' Get last row of active worksheet
rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

' For loop
For row = 2 To rowCount

    ' To get open Price for new ticker
    If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
        openPrice = ws.Cells(row, 3).Value
        totalStockVolume = 0
    End If
    
    ' Calculate Total Volume
    totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value

    ' To get last price for ticker
    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        ws.Cells(nextRow, 9).Value = ws.Cells(row, 1).Value
        ws.Cells(nextRow, 12).Value = totalStockVolume
        
        If totalStockVolume > greatestTotalVolume Then
            greatestTotalTicker = ws.Cells(row, 1).Value
            greatestTotalVolume = totalStockVolume
        End If
        
        
        closePrice = ws.Cells(row, 6).Value
        ws.Cells(nextRow, 10).Value = closePrice - openPrice
        ws.Cells(nextRow, 11).Value = (closePrice - openPrice) / openPrice
        ws.Cells(nextRow, 11).NumberFormat = "0.00%"
        
        If ws.Cells(nextRow, 11).Value > greatestIncrease Then
            greatestIncreaseTicker = ws.Cells(row, 1).Value
            greatestIncrease = ws.Cells(nextRow, 11).Value
        End If
        
        If ws.Cells(nextRow, 11).Value < greatestDecrease Then
            greatestDecreaseTicker = ws.Cells(row, 1).Value
            greatestDecrease = ws.Cells(nextRow, 11).Value
        End If
        
        If ws.Cells(nextRow, 10).Value >= 0 Then
            ws.Cells(nextRow, 10).Interior.Color = vbGreen
        Else
            ws.Cells(nextRow, 10).Interior.Color = vbRed
        End If
        
        If ws.Cells(nextRow, 11).Value >= 0 Then
            ws.Cells(nextRow, 11).Interior.Color = vbGreen
        Else
            ws.Cells(nextRow, 11).Interior.Color = vbRed
        End If
        nextRow = nextRow + 1
    End If

    Next row
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Range("O3") = "Greatest % Decrease"
    ws.Cells(3, 16).Value = greatestDecreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Range("O4") = "Greatest Total Volume"
    ws.Cells(4, 16).Value = greatestTotalTicker
    ws.Cells(4, 17).Value = greatestTotalVolume
    
    Next
End Sub
