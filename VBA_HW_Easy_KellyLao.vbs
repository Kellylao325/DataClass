Sub stock():
  
For Each ws In Worksheets
WorksheetName = ws.Name
  
  'add header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Volume"
    

'set variable holding ticker
Dim ticker_name As String
Dim ticker_total As Double
ticker_total = 0

'start the table
Dim table_row As Double
'start at row 2
table_row = 2

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (last_row)

For i = 2 To last_row
    'if next ticker in A not match with current ticker then print the ticker name
    'in the table row
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker_name = ws.Cells(i, 1).Value
    'add the volume total
   ticker_total = ticker_total + ws.Cells(i, 7).Value
    'print the ticker_name in table
    
    ws.Range("i" & table_row).Value = ticker_name
    'print ticker volume total to table
    ws.Range("j" & table_row).Value = ticker_total
    
    'next row in table and reset the ticker volume
    table_row = table_row + 1
    ticker_total = 0
    
    Else
    ticker_total = ticker_total + ws.Cells(i, 7).Value
    End If
    
Next i

Next ws

    
    
    
End Sub

