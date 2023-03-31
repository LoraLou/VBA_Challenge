Attribute VB_Name = "Module1"
Sub AnayzeStocks()


Dim Row As Long
Dim RowCount As Long
Dim NextRow As Long
Dim RowNumber As Long

Dim OpenValue As Double
Dim CloseValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double

Dim GreatestPercentIncreaseTicker As String
Dim GreatestPercentIncreaseValue As Double
Dim GreatestPercentDecreaseTicker As String
Dim GreatestPercentDecreaseValue As Double
Dim GreatestTotalVolumeTicker As String
Dim GreatestTotalVolumeValue As Double

Dim ws As Worksheet
For Each ws In Sheets

TotalStockVolume = 0
NextRow = 2

GreatestPercentIncreaseTicker = ""
GreatestPercentIncreaseValue = 0
GreatestPercentDecreaseTicker = ""
GreatestPercentDecreaseValue = 0
GreatestTotalVolumeTicker = ""
GreatestTotalVolumeValue = 0

RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

OpenValue = 0
CloseValue = 0
YearlyChange = 0
PercentChange = 0

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

For Row = 2 To RowCount
If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
TotalStockVolume = 0
OpenValue = ws.Cells(Row, 3).Value
End If

TotalStockVolume = TotalStockVolume + ws.Cells(Row, 7).Value
'row
If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
CloseValue = ws.Cells(Row, 6).Value
ws.Cells(NextRow, 9).Value = ws.Cells(Row, 1).Value
ws.Cells(NextRow, 10).Value = CloseValue - OpenValue
PercentChange = (CloseValue - OpenValue) / OpenValue
ws.Cells(NextRow, 11).Value = PercentChange
ws.Cells(NextRow, 12).Value = TotalStockVolume

ws.Cells(NextRow, 10).NumberFormat = "$0.00"
ws.Cells(NextRow, 11).NumberFormat = "0.00%"
If CloseValue > OpenValue Then
ws.Cells(NextRow, 10).Interior.Color = vbGreen
ws.Cells(NextRow, 11).Interior.Color = vbGreen
Else
ws.Cells(NextRow, 10).Interior.Color = vbRed
ws.Cells(NextRow, 11).Interior.Color = vbRed
End If

If TotalStockVolume > GreatestTotalVolumeValue Then
GreatestTotalVolumeTicker = ws.Cells(Row, 1).Value
GreatestTotalVolumeValue = TotalStockVolume

End If


If PercentChange > GreatestPercentIncreaseValue Then
GreatestPercentIncreaseTicker = ws.Cells(Row, 1).Value
GreatestPercentIncreaseValue = PercentChange

End If


If PercentChange < GreatestPercentDecreaseValue Then
GreatestPercentDecreaseTicker = ws.Cells(Row, 1).Value
GreatestPercentDecreaseValue = PercentChange

End If


NextRow = NextRow + 1
End If
Next Row

ws.Range("P4").Value = GreatestTotalVolumeTicker
ws.Range("Q4").Value = GreatestTotalVolumeValue
ws.Range("P2").Value = GreatestPercentIncreaseTicker
ws.Range("Q2").Value = GreatestPercentIncreaseValue
ws.Range("P3").Value = GreatestPercentDecreaseTicker

ws.Range("P2").Value = GreatestPercentIncreaseTicker
ws.Range("Q2").Value = GreatestPercentIncreaseValue
ws.Range("P3").Value = GreatestPercentDecreaseTicker
ws.Range("Q3").Value = GreatestPercentDecreaseValue
ws.Range("P4").Value = GreatestTotalVolumeTicker
ws.Range("Q4").Value = GreatestTotalVolumeValue



Next ws

End Sub











