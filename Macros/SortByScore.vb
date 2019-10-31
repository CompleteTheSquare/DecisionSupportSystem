Sub SortByScore()

'
    Range("A2:E10").Select
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
