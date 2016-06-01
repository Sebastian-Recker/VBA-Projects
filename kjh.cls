'This Sub renames all sheets to the value of each Sheets Cell A1
Sub RenameSheets()
Dim i as integer
For i =1 to Worksheets.Count
	Worksheets(i).name = Worksheets(i).cells(1,1)
Next i
End Sub
