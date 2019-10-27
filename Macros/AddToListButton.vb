
Sub AddToListButton()

Dim Task As Variant
Dim Duration As Variant
Dim DueDate As Variant
Dim CurrentDate As Variant
Dim Importance As Variant
Dim Score As Variant

CurrentDate = Date

Worksheets("Input").Activate
Task = Range("B6").Value
Duration = Range("B7").Value
DueDate = Range("B8").Value
Importance = Range("B9").Value



Range("B6").Value = ""
Range("B7").Value = ""
Range("B8").Value = ""
Range("B9").Value = ""




Worksheets("Tasklist").Activate
Dim Placement As Boolean
Dim RowNumber As Integer
Placement = False
RowNumber = 2



Do While Placement = False
If (Cells(RowNumber, 1) <> "") Then
RowNumber = RowNumber + 1
Else
Placement = True
End If
Loop


Score = Importance * Duration * 10 - (DueDate - CurrentDate)

Cells(RowNumber, 1) = Task
Cells(RowNumber, 2) = Duration
Cells(RowNumber, 3) = DueDate
Cells(RowNumber, 4) = Importance
Cells(RowNumber, 5) = Score

ActiveCell.NumberFormat = "@"

'January 1, 2020

End Sub



