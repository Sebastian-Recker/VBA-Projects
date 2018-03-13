'Template code to apply condition only on cells selected by user

Sub LoopThroughSelectedCells()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = FALSE

Dim SelRange As Range
Set SelRange = Selection

For Each c In SelRange.Cells
    If c.Value = "" Then
        c.clear
    End If
Next

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = TRUE

End Sub
