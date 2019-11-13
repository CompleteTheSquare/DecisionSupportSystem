

'for button4
'sorts by score
    Sub SortByScore()
    Range("A1:E90").Select
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Add2 Key:=Range("E1:E9") _
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





'get DurationSum
DurationRange = Worksheets("Sheet2").Range("B1", "B250")
DurationSum = WorksheetFunction.Sum(DurationRange)

'get Number Of People
Worksheets("Sheet1").Activate
Dim NumberOfPeople As Variant

NumberOfPeople = Range("E12")

'get work needed by everyone
Equal = DurationSum / NumberOfPeople





'get number of tasks per list
Dim TaskCount As Integer
Worksheets("Sheet2").Activate
TaskCount = WorksheetFunction.CountA([A1:A100])
MsgBox ("TaskCount: " & TaskCount)


'Worksheets("Sheet2").Activate
'    Rows("1:1").Select
'    Selection.Copy
'    Windows("DssProjFix - Copy.xlsm").Activate
'    Sheets("Sheet3").Select
        
 '   Cells(RowHolder, 1).Select
  '  ActiveSheet.Paste
    
    
'get PeopleInteger


Dim Origin As Integer
Origin = 1

Dim Multiply As Integer
Multiply = 0

Dim Adjustment As Integer
Adjustment = 0


Dim TaskCounter As Integer
TaskCounter = 0

Do While TaskCounter <> 10
For Multiply = 0 To NumberOfPeople - 1

RowHolder = 1 + 10 * (Multiply) + Adjustment


Worksheets("Sheet2").Activate

Rows("1:1").Select
Selection.Copy
    
Sheets("Sheet3").Select
Cells(RowHolder, 1).Select
ActiveSheet.Paste
        
  Worksheets("Sheet2").Activate
      Rows("1:1").Select
Selection.Delete Shift:=xlUp
        
        

    
    Next Multiply
Adjustment = Adjustment + 1

TaskCounter = TaskCounter + 1

Loop





'


    'PeopleInteger = PeopleInteger + 1
    'Cells(RowHolder, 1).Select
    'ActiveSheet.Paste



'For i = 1 To TaskCount
'RowHolder = PeopleInteger * 5 + Constant + 1

    
    

  
 'Next i
  
'PeopleInteger = PeopleInteger + 1
  
      MsgBox ("successful placement")

    'Range("A1:E1").Select
    'Selection.Copy
    'Sheets("Sheet3").Select
    'Range("A1").Select
    'ActiveSheet.Paste
  
  
  
  
    
'Worksheets("Sheet3").Activate
'    Selection.Delete Shift:=xlUp
'    Windows("DssProjFix.xlsm").Activate
'    Selection.Delete Shift:=xlUp

End Sub

