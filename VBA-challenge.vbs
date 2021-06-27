Attribute VB_Name = "Module1"
Sub wall()

'set  variables
Dim ticker As String
Dim opening As Double
Dim change As Double
Dim percent As Double
Dim volume As Long

Dim greatestinc As Double
Dim greatestdec As Double
Dim greatestvol As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
opening = Cells(2, 3)
For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ticker = Cells(i, 1)
    change = Cells(i, 6) - opening
    percent = ((change) / (opening)) * 100
    
    Range("I" & Summary_Table_Row) = ticker
    Range("J" & Summary_Table_Row) = change
    
End If
Next i

End Sub

