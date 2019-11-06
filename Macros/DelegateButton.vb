Sub SortByScore()


Dim PeopleNumber As Integer

Dim DurationSum As Integer



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

Worksheets("Input").Activate
PeopleNumber = Range("B15").Value

Worksheets("Tasklist").Activate

myRange = Worksheets("TaskList").Range("B1", "B250")
DurationSum = WorksheetFunction.Sum(myRange)
MsgBox ("the sum of duration is: " & DurationSum) 'total time to divide





'equal distribution
Dim EqualDistribution As Double
MsgBox ("pplnumber: " & PeopleNumber)
EqualDistribution = DurationSum / PeopleNumber ' find the number of hours for "equal work distribution"


'====================================================
Dim RowHolder As Integer 'where you are going to place the selections (row)
Dim ColumnHolder As Integer 'where you are going to place the selections (column)

Dim RowSelected As Integer ' row selected to paste in
RowSelected = 2









Do While DurationSum <> 0

DurationSum = WorksheetFunction.Sum(myRange)


Worksheets("TaskList").Activate ' activate the 2nd sheet
Range("A2:E2").Select '
Selection.Copy

Worksheets("Order").Activate

Dim IsEmpty As Boolean
IsEmpty = False

RowHolder = 2

'MsgBox ("RowHolder " & RowHolder)
'Do While IsEmpty = False
If Cells(RowHolder, 1) = "" Then
IsEmpty = True
Else

MsgBox ("RowHolder " & RowHolder)
RowHolder = RowHolder + 1
End If
'Loop

MsgBox ("RowHolder " & RowHolder)

Range(1, RowHolder).Select 'BAD
ActiveSheet.Paste



Worksheets("TaskList").Activate '
Rows("2:2").Select
Selection.Delete Shift:=xlUp


DurationSum = 0

Loop


MsgBox ("yeehaw successful run")
End Sub



























