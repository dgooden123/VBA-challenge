Attribute VB_Name = "Module1"
Sub wall()
'set  variables
Dim ticker As String
Dim start As Long
Dim change As Double
Dim percent As Double
Dim volume As Long

'set challenge variables
Dim greatestinc As Double
Dim greatestdec As Double
Dim greatestvol As Long

'set variable for summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'set last row variable
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'set volume to zero
volume = 0
start = 2

'loop through worksheets
For Each ws In Worksheets

'need a non zero opening price
If ws.Cells(start, 3) = 0 Then
                   For find_value = start To i
                       If ws.Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               
'review each ticker name
For i = 2 To LastRow

'if the ticker names differ
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'then calculate the values
    
    ticker = ws.Cells(i, 1).Value
    change = ws.Cells(i, 6).Value - start
    percent = ((change) / (start)) * 100
    volume = volume + ws.Cells(i, 7).Value
    
    
    
    ws.Range("I" & Summary_Table_Row) = ticker
    ws.Range("J" & Summary_Table_Row) = change
    ws.Range("K" & Summary_Table_Row) = percent
    ws.Range("L" & Summary_Table_Row) = volume
    
    Summary_Table_Row = Summary_Table_Row + 1
    'add headers
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
   
    'reset values to 0
    change = 0
    percent = 0
    volume = 0
    

'if the ticker name does not change, add the volume and keep going
Else
          volume = volume + ws.Cells(i, 7).Value
       End If

If ws.Range("J" & i).Value >= 0 Then

        ws.Range("J" & i).Interior.ColorIndex = 4

ElseIf ws.Range("J" & i).Value < 0 Then

        ws.Range("J" & i).Interior.ColorIndex = 3
 End If
 
        Next i

   
'find the greatest decrease and decrease, and total volume
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
        
        
increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
        
        
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
Next ws

        
End Sub



