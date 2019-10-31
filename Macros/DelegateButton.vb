Sub Delegate()

Dim Information(2 To 100, 1 To 5) As Variant
Dim Holder(1 To 1, 1 To 1) As String


Dim Task(2 To 99) As String
Dim Duration(2 To 99) As Integer
Dim DueDate(2 To 99) As Date
Dim CurrentDate(2 To 99) As Date
Dim Importance(2 To 99) As Integer
Dim Score(2 To 99) As Integer
Dim RowNumber(2 To 99) As Integer

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



Worksheets("Input").Activate
PeopleNumber = Cells(2, 15) ' takes the number of ppl in team

Worksheets("Tasklist").Activate

myRange = Worksheets("TaskList").Range("B1", "B250")
DurationSum = WorksheetFunction.Sum(myRange)
MsgBox ("the sum of duration is: " & DurationSum) 'total time to divide





'equal distribution
Dim EqualDistribution As Double

EqualDistribution = DurationSum / PeopleNumber ' find the number of hours for "equal work distribution"

Worksheets("TaskList").Activate
'====================================================
Dim RowHolder As Integer
Dim ColumnHolder As Integer

Dim RowSelected As Integer
RowSelected = 2

Do While DurationSum<> 0 

DurationSum = WorksheetFunction.Sum(myRange)



Range("A2:E2").Select
Selection.Copy
Worksheets("Order").Activate
Range(1,RowSelected)' cell it should be pasted in
ActiveSheet.Paste


DurationSum = 0

Loop


MsgBox ("yeehaw successful run")
End Sub





