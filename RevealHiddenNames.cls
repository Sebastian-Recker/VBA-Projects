'This function reveals all hidden Names, so that you can then delete them.
'Usefull to avoid the "Change name" dialogue box poping up when you duplicate or move a sheet.
Sub RevealHiddenNames()
Dim i As Integer
   For i = 1 To ActiveWorkbook.Names.Count
        ActiveWorkbook.Names(i).Visible = True
    Next i
End Sub
