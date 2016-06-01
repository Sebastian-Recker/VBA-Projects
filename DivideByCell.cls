'This Sub divides all cells which are not SUMs between (5,4) to (197,16) by the cell $C$2
Sub DivideByCell()
Dim i As Integer
Dim j As Integer

For j = 4 To 16
    For i = 5 To 197
        If IsEmpty(Cells(i, j)) Then
        Else
            If Left(Cells(i, j).Formula, 4) = "=SUM" Then
            Else
                Cells(i, j).Formula = "=(" & Right(Cells(i, j).Formula, Len(Cells(i, j).Formula) - 1) & ")/$C$2"
            End If
        End If
    Next i
Next j


End Sub
