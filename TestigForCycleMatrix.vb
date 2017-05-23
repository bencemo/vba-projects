Sub TestingForCycleMatrix

'Testing For cycles by filling matrixes

'Define variables
Dim i As Integer
Dim j As Integer
Dim RowNumber As Integer
Dim ColumnNumber As Integer

'Clear all previous cells to easily see outcome
    Cells.Select
    Range("D4").Activate
    Selection.ClearContents

'Start the for cycles

For i = 1 To 5

	RowNumber = i
	ColumnNumber = 1
	Cells(RowNumber, ColumnNumber).Select
	ActiveCell = i
Next i

End Sub