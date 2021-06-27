Attribute VB_Name = "Module1"
Sub wall()

'set  variables
Dim ticker As String
Dim opening As Double
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
volume = zero

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ticker = Cells(i, 1).Value
    change = Cells(i, 6).Value - opening
    percent = ((change) / (opening)) * 100
    volume = volume + Cells(i, 7).Value
    
    Range("I" & Summary_Table_Row) = ticker
    Range("J" & Summary_Table_Row) = change
    Range("K" & Summary_Table_Row) = volume
    
    Else: volume = volume + Cells(i, 7).Value
    
    
End If

If change >= 0 Then Range("J", i).Interior.ColorIndex = 3

    Else: Range("J", i).Interior.ColorIndex = 4
End If

Next i


End Sub

