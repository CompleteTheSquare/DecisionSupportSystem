Sub SortByScore()




    Range("A2:E100").Select
    Range("E10").Activate
    ActiveWorkbook.Worksheets("Tasklist").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Tasklist").Sort.SortFields.Add2 Key:=Range( _
        "E2:E10"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Tasklist").Sort
        .SetRange Range("A1:E10")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub

Sub AddToOrder()
'-------------------------------------

Dim PeopleNumber As Integer
Dim EqualDistribution As Double
Dim DurationSum As Integer
Dim RowHolder As Integer 'where you are going to place the selections (row)
RowHolder = 2
Dim ColumnHolder As Integer 'where you are going to place the selections (column)
Dim IsEmpty As Boolean



Worksheets("Input").Activate
PeopleNumber = Range("B15").Value





EqualDistribution = DurationSum / PeopleNumber ' find the number of hours for "equal work distribution"
Worksheets("Tasklist").Activate

myRange = Worksheets("TaskList").Range("B1", "B250")
DurationSum = WorksheetFunction.Sum(myRange)
MsgBox ("the sum of duration is: " & DurationSum) 'total time to divide
MsgBox ("pplnumb " & PeopleNumber)
MsgBox ("RowHolder " & RowHolder)

Dim RowSelected As Integer ' row selected to paste in
RowSelected = 2

Worksheets("TaskList").Activate ' activate the 2nd sheet
Range("A2:E2").Select 'select the first row
Selection.Copy 'copy the damned thing

Worksheets("Order").Activate ' go to the 3rd sheet



'find out where the first space is available 
' and throw info there
MsgBox ("the sum of duration is: " & DurationSum) 'total time to divide
MsgBox ("pplnumb " & PeopleNumber)
MsgBox ("RowHolder " & RowHolder)

Do While IsEmpty = False
If Cells(RowHolder, 1) = "" Then
IsEmpty = True
Else
MsgBox ("RowHolder " & RowHolder)
RowHolder = RowHolder + 1
End If
Loop

MsgBox ("the sum of duration is: " & DurationSum) 'total time to divide
MsgBox ("pplnumb " & PeopleNumber)
MsgBox ("RowHolder " & RowHolder)


Range(1, RowHolder).Select 'BAD
ActiveSheet.Paste

'delete that row entirely, everything shifts up
Worksheets("TaskList").Activate '
Rows("2:2").Select
Selection.Delete Shift:=xlUp







'======================================================
End Sub








