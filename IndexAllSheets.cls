' This Sub creates a vertical list of the name of all sheets in the worbook on the First sheet.
Sub IndexAllSheets()
Dim i As Integer

For i = 1 To Worksheets.count
 Worksheets(1).Cells(i, 2) = Worksheets(i).Name
Next i

End Sub
