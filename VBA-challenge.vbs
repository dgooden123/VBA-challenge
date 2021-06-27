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
Dim greatestvol As Double

'set variable for summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'set last row variable
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'set volume to zero
volume = 0
start = 2

If Cells(start, 3) = 0 Then
                   For find_value = start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               
For i = 2 To LastRow

'if the ticker names differ
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'then calculate the values
    
    ticker = Cells(i, 1).Value
    change = Cells(i, 6).Value - start
    percent = ((change) / (start)) * 100
    volume = volume + Cells(i, 7).Value
    
    Range("I" & Summary_Table_Row) = ticker
    Range("J" & Summary_Table_Row) = change
    Range("K" & Summary_Table_Row) = percent
    Range("L" & Summary_Table_Row) = volume
    
    'reset values to 0
    change = 0
    percent = 0
    volume = 0
'if the ticker name does not change, add the volume and keep going
Else
          volume = volume + Cells(i, 7).Value
       End If
   Next i
   
End Sub


