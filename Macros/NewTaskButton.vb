
'this displays textboxes for the user to fill in
'if the input is invalid, it asks the user for valid input

Sub AddInformation()

Dim Task As Variant
Dim Duration As Variant
Dim DueDate As Variant
Dim Importance As Variant


Task = InputBox("Input Task Here")
Range("B6").Value = Task

Duration = InputBox("Input Duration Here")
Range("B7").Value = Duration

DueDate = InputBox("Input Due Date Here")
Range("B8").Value = DueDate

Importance = InputBox("Input Importance level (1-10) Here")
Range("B9").Value = Importance

End Sub
