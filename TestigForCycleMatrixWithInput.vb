Sub TestingForCycleMatrixWithInput()

'Testing For cycles by filling matrixes

'Define variables
Dim i As Integer
Dim iu As Integer
Dim j As Integer
Dim ju As Integer

'Clear all previous cells to easily see outcome
Cells.Select
Range("D4").Activate
Selection.ClearContents

'Ask for User input
ju = InputBox("How many columns?")
iu = InputBox("How many rows?")

'Start the for cycles
For j = 1 To ju
    For i = 1 To iu
        Cells(i, j).Select
        ActiveCell = i
        Next i
Next j
End Sub