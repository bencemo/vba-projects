Sub TestingForCycleMatrix2()

'Testing For cycles by filling matrixes

'Define variables
Dim i As Integer
Dim j As Integer

'Clear all previous cells to easily see outcome
Cells.Select
Range("D4").Activate
Selection.ClearContents

'Start the for cycles
For j = 1 To 5
    For i = 1 To 5
        Cells(i, j).Select
        ActiveCell = i
        Next i
	Next j
End Sub