'This Sub clears selected cells which formulas do not start by "=SUM"
'This Sub is mainly used to to remove links to other sheets while keeping calculated sub-totals

Sub ClearIfNotSUMinSelection()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim SelRange As Range
Set SelRange = Selection

For Each c In SelRange.Cells
    If Left(c.Formula, 4) = "=SUM" Then
    Else
        c.Clear
    End If
Next

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
