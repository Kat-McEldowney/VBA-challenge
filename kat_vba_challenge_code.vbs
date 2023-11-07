Sub vba_HW()

Dim ws As worksheet
Dim wb As Workbook


For Each ws In Worksheets

' initilize variable for holding opening price, calculate volume total, get last rows, and
'calculate biggest gains, losses and total volume



Dim opening_price As Double
Dim total_volume As Double
Dim LastRow As Double
Dim Summary_LastRow As Double
Dim g_increase_amt As Double
Dim g_increase_tkr As String
Dim g_decrease_amt As Double
Dim g_decrease_tkr As String
Dim total_vol_amt As Double
Dim total_vol_tkr As String
Dim j As Integer

'get last row to loop from row 2 to the end
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

closing_price = 0
yearly_change = 0
total_volume = 0

'set start values of variables,
j = 2
opening_price = ws.Cells(2, 3).Value
g_increase_amt = ws.Cells(2, 11).Value
g_decrease_amt = ws.Cells(2, 11).Value
total_vol_amt = ws.Cells(2, 12).Value
g_increase_tkr = ws.Cells(2, 1).Value
g_decrease_tkr = ws.Cells(2, 1).Value
total_vol_tkr = ws.Cells(2, 1).Value

'add headings to worksheet
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'iterate through the data to get the required info, calculate the yearly change and total volume traded

For i = 2 To LastRow

  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'record ticker name in new ticker column
    ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
    
    'calculate yearly change
    ws.Cells(j, 10).Value = (ws.Cells(i, 6).Value - opening_price)
    
    'color code increases and decreases
    If ws.Cells(j, 10).Value > 0 Then
      ws.Cells(j, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 10).Value < 0 Then
      ws.Cells(j, 10).Interior.ColorIndex = 3
   
     End If

    'calculate percent change
    ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / opening_price
    ws.Cells(j, 11).Value = FormatPercent(ws.Cells(j, 11))
    
    'record total volume in column L
    ws.Cells(j, 12).Value = total_volume
    'add one to j to advance totals to the next row, reset opening price and total volume counter
    j = j + 1
    ticker_name = ws.Cells(i + 1, 1).Value
    total_volume = 0
    opening_price = ws.Cells(i + 1, 3).Value
    
    
  Else:
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    
    
  End If


Next i

Summary_LastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row


For k = 2 To Summary_LastRow
 If ws.Cells(k, 11).Value > g_increase_amt Then
    g_increase_amt = ws.Cells(k, 11).Value
    g_increase_tkr = ws.Cells(k, 9).Value
  ElseIf ws.Cells(k, 11).Value < g_decrease_amt Then
    g_decrease_amt = ws.Cells(k, 11).Value
    g_decrease_tkr = ws.Cells(k, 9).Value
    End If
  If ws.Cells(k, 12).Value > total_vol_amt Then
    total_vol_amt = ws.Cells(k, 12).Value
    total_vol_tkr = ws.Cells(k, 9).Value
    End If
    
    Next k
    
'print new summary info titles and values
    
ws.Cells(2, 15).Value = "Ticker"
ws.Cells(2, 16).Value = "Value"
ws.Cells(3, 14).Value = "Greatest % Increase"
ws.Cells(4, 14).Value = "Greatest % Decrease"
ws.Cells(5, 14).Value = "Greatest Total Volume"
ws.Cells(3, 15).Value = g_increase_tkr
ws.Cells(4, 15).Value = g_decrease_tkr
ws.Cells(5, 15).Value = total_vol_tkr
ws.Cells(3, 16).Value = g_increase_amt
ws.Cells(3, 16).Value = FormatPercent(ws.Cells(3, 16))
ws.Cells(4, 16).Value = g_decrease_amt
ws.Cells(4, 16).Value = FormatPercent(ws.Cells(4, 16))
ws.Cells(5, 16).Value = total_vol_amt

'autofit data in new columns
ws.Cells.EntireColumn.AutoFit

Next ws

End Sub
