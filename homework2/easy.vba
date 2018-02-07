Sub homework2():

Dim rs As Worksheet

For Each rs In ActiveWorkbook.Worksheets


rs.Cells(1, 9).Value = "Ticker"
rs.Cells(1, 10).Value = "Total Stock Value"


'get the last row of the sheet'
Dim lRow As Long
lRow = rs.Cells(Rows.Count, 1).End(xlUp).Row

'define counters'
Dim tickerCount As Integer
Dim tickerSum As Double

'initial values'
tickerCount = 1
tickerSum = 0

'where the magic happens'
For i = 2 To lRow:

'add up ticker sum'
tickerSum = tickerSum + rs.Cells(i, 7).Value

'CHANGE IN TICKER'
If (rs.Cells(i + 1, 1).Value <> rs.Cells(i, 1).Value) Then
rs.Cells(tickerCount + 1, 10).Value = tickerSum
rs.Cells(tickerCount + 1, 9).Value = rs.Cells(i, 1).Value
tickerCount = tickerCount + 1
tickerSum = 0
End If


Next i

Next rs


End Sub
