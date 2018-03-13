Sub ClearEmptyLinkedCells()
Dim i As Integer
Dim j As Integer

For i = 9 To 51
    For j = 40 To 52
        If Cells(i, j).Value = 0 Then
            Cells(i, j).Clear
        End If
    Next j
Next i

End Sub
