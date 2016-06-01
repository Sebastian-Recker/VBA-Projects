'This Sub searches for a List of words in the same column over many sheets.
'If it finds a match it writes a value, here "exceptional", in the column I on the same line
'This alows the use of a SUMIF() to sum the value associated with a list of names in a big set of data
Sub FindStrings()
Dim i As Integer
Dim j As Integer
Dim count_hits As Integer
count_hits = 0

For j = 6 To 10
    For i = 5 To 82
        count_hits = 0
        With Worksheets(i).Range("c3:c7000")
            Set c = .Find(Worksheets(4).Cells(j, "B").Value, LookIn:=xlValues)
            If Not c Is Nothing Then
                firstAddress = c.Address
                Do
                    Cells(c.Row, "I").Value = "Exceptional"
                    Set c = .FindNext(c)
                    count_hits = count_hits + 1
                Loop While Not c Is Nothing And c.Address <> firstAddress
            End If
            Worksheets(4).Cells(j, i - 2).Value = count_hits
        End With
    Next i
Next j
End Sub
