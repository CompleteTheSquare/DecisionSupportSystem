
Sub AddToListButton()

Dim Task As Variant
Dim Duration As Variant
Dim DueDate As Variant
Dim Importance As Variant


Task = Range("B6").Value
Duration = Range("B7").Value
DueDate = Range("B8").Value
Importance = Range("B9").Value

Dim Placement As Boolean
Dim RowNumber As Integer

Placement = False
RowNumber = 2

Do While Placement = False

If IsEmpty(Cells(RowNumber, 1)) = False Then RowNumber = RowNumber + 1
If IsEmpty(Cells(RowNumber, 1)) = True Then Placement = True


'this method moves the information to the tasklist then clears the fields
MsgBox ("The Cell " & RowNumber & " 1 has info in it")

Loop
End Sub

