Sub TestingForCyclye1()

'Only for testing, practicing purposes

'Ask for variants

Dim i As Integer
Dim UserInput As Variant
Dim MyColumn As String
Dim MyCell  As String

'Clear all previous cells
    Cells.Select
    Range("D4").Activate
    Selection.ClearContents

'Define variants
MyColumn = "A"

'Ask for user's input
UserInput = InputBox("How many cells?")

'Run the for cycle
For i = 1 To UserInput
    MyCell = MyColumn & i
    Range(MyCell).Select
    ActiveCell = i
    
Next i

'End the Sub
End Sub