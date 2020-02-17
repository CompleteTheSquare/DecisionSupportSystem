Sub SortListAddToPlace()
' this sorts the information in Even and Uneven
' it then places the information into the EvenPlace and UnevenPlace sheets using with the formulas

'sort methods that take each list of data and sort them by Score'
Columns("A:G").Select
ActiveWorkbook.Worksheets("EqualList").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("EqualList").Sort.SortFields.Add2 Key:=Range("G1:G100000") _
, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("EqualList").Sort
.SetRange Range("A1:G100000")
.Header = xlGuess
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With

'sort methods that take each list of data and sort them by Score'
Columns("A:G").Select
ActiveWorkbook.Worksheets("UnequalList").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("UnequalList").Sort.SortFields.Add2 Key:=Range("G1:G100000") _
, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("UnequalList").Sort
.SetRange Range("A1:G100000")
.Header = xlGuess
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With


'get TotPpl: number of people involved in proejct
Worksheets("Menu").Activate
Dim NumberOfPeople As Variant
NumberOfPeople = Range("D18") 'this cell has number of ppl in group
'end get Number of People

'get Taskcount: number of tasks per list
Dim TaskCount As Integer
Worksheets("EqualList").Activate
TaskCount = WorksheetFunction.CountA([A1:A1000])
'end get number of tasks per list

'get equal distribution and Spacing
Dim EqualDistribution As Double
Dim Spacing As Integer
Spacing = 0
Dim SpacingConstant As Integer
EqualDistribution = TaskCount / NumberOfPeople
EqualDistribution = TaskCount / NumberOfPeople
SpacingConstant = Round(EqualDistribution) + 1
'end get equal distribution and spacing

'paste the sorted list data into the unequal list sheet'
Sheets("EqualList").Select
Cells.Select
Selection.Copy
Sheets("UnequalList").Select
Range("A1").Select
ActiveSheet.Paste
Range("A1").Select
Dim Adjustment As Integer
'end paste the sorted list data into the unequal list sheet'

'get the row number:
Adjustment = 1
Dim Origin As Integer
Origin = 1
Dim RowNumber As Integer
RowNumber = Origin * Spacing + Adjustment
Dim i As Integer
Dim j As Integer
i = 1
Dim TaskCountHolder As Integer
TaskCountHolder = 0
Do While TaskCountHolder < TaskCount
For i = 0 To NumberOfPeople - 1
RowNumber = Origin * Spacing + Adjustment
Spacing = Spacing + SpacingConstant
Worksheets("EqualList").Activate
Rows("1:1").Select
Selection.Copy
Sheets("EqualPlace").Select
Cells(RowNumber, 1).Select
ActiveSheet.Paste
Worksheets("EqualList").Activate
Rows("1:1").Select
Selection.Delete Shift:=xlUp
TaskCountHolder = TaskCountHolder + 1
Next i
Adjustment = Adjustment + 1
For j = 0 To NumberOfPeople - 1
Spacing = Spacing - SpacingConstant
RowNumber = Origin * Spacing + Adjustment
Worksheets("EqualList").Activate
Rows("1:1").Select
Selection.Copy
Sheets("EqualPlace").Select
Cells(RowNumber, 1).Select
ActiveSheet.Paste
Worksheets("EqualList").Activate
Rows("1:1").Select
Selection.Delete Shift:=xlUp
TaskCountHolder = TaskCountHolder + 1
Next j
Adjustment = Adjustment + 1


Loop


'get TotPpl: number of people involved in proejct
Worksheets("Menu").Activate
Dim NumberOfPeople2 As Variant
NumberOfPeople2 = Range("D18") 'this cell has number of ppl in group
'end get Number of People


'get Taskcount: number of tasks per list
Dim TaskCount2 As Integer
Worksheets("UnequalList").Activate
TaskCount2 = WorksheetFunction.CountA([A1:A1000])

'end get number of tasks per list

'get equal distribution and Spacing
Dim EqualDistribution2 As Double
Dim SpacingConstant2 As Integer
EqualDistribution2 = TaskCount2 / NumberOfPeople2

EqualDistribution2 = TaskCount2 / NumberOfPeople2

SpacingConstant2 = Round(EqualDistribution2) + 1
'end get equal distribution and spacing

Dim TaskCountHolder2 As Integer
TaskCountHolder2 = 0

Dim Adjustment2 As Integer
'get the row number:
Adjustment2 = 1
Dim Origin2 As Integer
Origin2 = 1
Dim RowNumber2 As Integer
Dim Spacing2 As Integer

Spacing2 = 0
RowNumber2 = Origin2 * Spacing2 + Adjustment2

Do While TaskCountHolder2 < TaskCount2




For i = 0 To NumberOfPeople2 - 1



RowNumber2 = Origin2 * Spacing2 + Adjustment2


Worksheets("UnequalList").Activate
Rows("1:1").Select
Selection.Copy

Worksheets("UnequalPlace").Activate
Cells(RowNumber2, 1).Select
ActiveSheet.Paste



Worksheets("UnequalList").Activate
      Rows("1:1").Select
Selection.Delete Shift:=xlUp


Spacing2 = SpacingConstant2 + Spacing2

TaskCountHolder2 = TaskCountHolder2 + 1


EqualDistribution2 = TaskCount2 / NumberOfPeople2


Next i
Adjustment2 = Adjustment2 + 1
Spacing2 = 0


Loop



End Sub


