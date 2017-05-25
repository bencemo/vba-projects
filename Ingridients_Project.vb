Sub IngridientsCompiler()

'Created for XLnoobie, requesting help on 25/05/2017 at http://bit.ly/2r2f9Gm
'Written by Csaba Bence Molnár - u/DonKajak on 25/05/2017

'Define Variables
Dim i As Integer 'Index
Dim j As Integer 'Numer of added ingridients
Dim LastIng As String 'Last added ingridient
Dim IngOutPut As String 'Final outcome string

'Set Variables to Zero
i = 0
j = 0
LastIng = ""
IngOutPut = ""

'Setting array
Dim Ingridients(1 To 4) As String

Ingridients(1) = "Ham"
Ingridients(2) = "Turkey"
Ingridients(3) = "Salami"
Ingridients(4) = "Roast Beef"

'Starting the for cycle, checking for NOT empty cells in checkbox
'Cells is currently pointing at 'D6', you must relocate it
For i = 1 To 4
    If Not Cells(5 + i, 4) = "" Then
    
        'Adding checked ingridients + "c" to IngOutPut
        IngOutPut = IngOutPut & Ingridients(i) & ", "
        
        'Getting to know to last added ingridient
        LastIng = Ingridients(i)
        
        'Counting added ingridients to correctly adjust format later
        j = j + 1
       
    End If
Next i

'Deleting unneccessary ", " from IngOutPut end
IngOutPut = Left(IngOutPut, Len(IngOutPut) - 2)


If j > 1 Then 'Checking whether there is one/more ingridients
    
    'Correcting format by deleting ", "+LastIng
    IngOutPut = Left(IngOutPut, Len(IngOutPut) - Len(LastIng) - 2)
    
    'Readding LastIng to the end of string
    IngOutPut = IngOutPut & " and " & LastIng
    
    Else 'If there is only one ingridient, IngOutPut is solely LastIng
    
    IngOutPut = LastIng

End If

'Write solution to 'A1' cell
Cells(1, 1) = IngOutPut

End Sub