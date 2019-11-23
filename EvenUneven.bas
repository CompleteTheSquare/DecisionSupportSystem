Sub Even()

'clear the list out before pasting anything
    Sheets("Menu").Select
    Range("K4:L172").Select
    Selection.Delete Shift:=xlUp




'goes to the EqualPlace Sheet and counts how many tasks there are
    Sheets("EqualPlace").Select
TaskCount = WorksheetFunction.CountA([A1:A100])

'goes to the cells and copies the list of tasks
    Sheets("EqualPlace").Select
    Range("A1:A1000").Select
    Selection.Copy
    Sheets("Menu").Select
    Range("K5").Select
    ActiveSheet.Paste

'goes to the cells and copies the list of due dates
    Sheets("EqualPlace").Select
    Range("E1:E1000").Select
    Selection.Copy
    Sheets("Menu").Select
    Range("L5").Select
    ActiveSheet.Paste
    
    
    Dim i As Integer
    Dim j As Integer
    

'it also gets rid of blank spaces between dates and then presents the divisions
For i = 4 To TaskCount + 4

If Cells(i, 11) = "" Then
Cells(i, 11).Select
    Selection.Delete Shift:=xlUp
    Cells(i, 11).Select
        Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'recorded macro that takes the range and puts a border right below it
    End If
Next i


'it also gets rid of blank spaces between dates and then presents the divisions
For j = 4 To TaskCount + 4
If Cells(j, 12) = "" Then
Cells(j, 12).Select
    Selection.Delete Shift:=xlUp
        Cells(j, 12).Select
               Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'recorded macro that takes the range and puts a border right below it
        End If
    Next j
  
End Sub



'clear the list out before pasting anything
Sub Uneven()
    Sheets("Menu").Select
    Range("K4:L172").Select
    Selection.Delete Shift:=xlUp
    
    
    
    


'goes to the Unequal Sheet and counts how many tasks there are
    Sheets("UnequalPlace").Select
TaskCount = WorksheetFunction.CountA([A1:A100])

    Sheets("UnequalPlace").Select
    Range("A1:A1000").Select
    Selection.Copy
    Sheets("Menu").Select
    Range("K5").Select
    ActiveSheet.Paste

    Sheets("UnequalPlace").Select
    Range("E1:E1000").Select
    Selection.Copy
    Sheets("Menu").Select
    Range("L5").Select
    ActiveSheet.Paste
    
    
    Dim i As Integer
    Dim j As Integer
    
'goes to the cells and copies the list of tasks
'it also gets rid of blank spaces between dates and then presents the divisions
For i = 4 To TaskCount + 4

If Cells(i, 11) = "" Then
Cells(i, 11).Select
    Selection.Delete Shift:=xlUp
    Cells(i, 11).Select
        Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'recorded macro that takes the range and puts a border right below it
    End If
Next i

'goes to the cells and copies the list of due dates
'it also gets rid of blank spaces between dates and then presents the divisions

For j = 4 To TaskCount + 4
If Cells(j, 12) = "" Then
Cells(j, 12).Select
    Selection.Delete Shift:=xlUp
        Cells(j, 12).Select
               Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'recorded macro that takes the range and puts a border right below it
        End If
    Next j
  



End Sub
