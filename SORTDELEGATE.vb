

'for button4
'sorts by score
    Sub SortByScore()
    Range("A1:E90").Select
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Add2 Key:=Range("A1:E90") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet2").Sort
        .SetRange Range("A1:E9")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

    MsgBox ("successful sort")




'get TotPpl: numbe of people involved in proejct
Worksheets("Sheet1").Activate
Dim NumberOfPeople As Variant
NumberOfPeople = Range("E12")
'end get Number of People


'get Tasksum: number of tasks per list
Dim TaskCount As Integer
Worksheets("Sheet2").Activate
TaskCount = WorksheetFunction.CountA([A1:A100])
MsgBox ("TaskCount: " & TaskCount)
'end get number of tasks per list

'get equal distribution and Spacing
Dim EqualDistribution as Double
Dim Spacing as Integer
Spacing = round(EqualDistribution) + 2
'end get equal distribution and spacing


Dim Adjustment As Integer

'get the row number:
Dim Origin As Integer
Origin = 1
Dim RowNumber As Integer
RowNumber = Origin * Spacing + Adjustment

Do While Integer i = 0 
MsgBox ("RowNumber = " & RowNumber)
For integer i < EqualDistribution

RowNumber = Origin * Spacing + Adjustment

Next i 
MsgBox ("RowNumber = " & RowNumber)


End Sub

