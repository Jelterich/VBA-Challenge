Sub Multiple_Quarter_Stock_Data():
 
 ' Declare variables
 
 Dim ws As Worksheet
 
 Dim ticker As String
 
 Dim openValue As Double
 
 Dim closingValue As Double
 
 Dim quarterlyChange As Double
 
 Dim percentChange As Double
 
 Dim totalVolume As Double
 
 Dim lasRow As Long
 
 Dim i As Long
 
 Dim SummaryRow As Integer
 
 ' Loop through all sheets
 For Each ws In ThisWorkbook.Worksheets
 
 ' Initialize summary row
 SummaryRow = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



' Find the last row with data

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Loop through all rows

For i = 2 To lastRow
 
 ' Check if we are still within the same ticker symbol
 
 If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
  
  ' Set the ticker
  
  ticker = ws.Cells(i, 1).Value
  
 ' Set opening price
 
 openingValue = ws.Cells(i, 3).Value
 
 ' Set closing price
 
 closingValue = ws.Cells(i, 6).Value
 
  ' Calculate the quarterly change
  
  quarterlyChange = closingValue - openingValue
  
  ' Calculate the percentage change
  
  If openingValue <> 0 Then
  percentChange = (quarterlyChange / openingValue)
  Else
  percentChange = 0
  End If
  
  ' Calculate the total volume
  
  totalVolume = totalVolume + ws.Cells(i, 7).Value
  
  ' Output the results
  ws.Cells(SummaryRow, 9).Value = ticker
  ws.Cells(SummaryRow, 10).Value = quarterlyChange
  ws.Cells(SummaryRow, 11).Value = percentChange
  ws.Cells(SummaryRow, 12).Value = totalVolume
  
  ' Increment summary row
  SummaryRow = SummaryRow + 1
  
  ' Reset total volume
  
totalVolume = 0
  
Else
  
totalVolume = totalVolume + ws.Cells(i, 7).Value
   
   End If
Next i

Next ws

End Sub
