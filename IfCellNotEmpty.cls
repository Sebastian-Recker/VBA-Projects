Sub IfCellNotEmpty()
Dim i As Integer

For i = 4 To 3000
    If IsEmpty(Cells(i, "M")) Then
    Else
        Cells(i, "I").Formula = Cells(i - 1, "I").Formula
        Cells(i, "I").Interior.ColorIndex = 7
    End If
Next i

End Sub
