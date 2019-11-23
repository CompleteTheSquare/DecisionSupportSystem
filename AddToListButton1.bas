
Sub AddToListButton1()

Dim Task As Variant
Dim Duration As Variant
Dim Name As Variant
Dim DueDate As Variant
Dim StartDate As Variant
Dim Importance As Variant
Dim Chunks As Integer
Dim Score As Variant
Dim TimePerChunk As Variant

Name = ""

' goes to the first sheet
Worksheets("Menu").Activate

'stores the values in the cells
Task = Range("D23").Value
Duration = Range("D24").Value
StartDate = CDate(Range("D25"))
DueDate = CDate(Range("D26"))
Importance = Range("D27").Value
Chunk = Range("D28").Value

'finds the time and importance for each chunk of work
TimePerChunk = Duration / Chunk
Importance = Importance / Chunk

Dim i As Integer
For i = 1 To Chunk
Duration = TimePerChunk
Worksheets("EqualList").Activate
Dim Placement As Boolean
Dim RowNumber As Integer
Placement = False
RowNumber = 1

Do While Placement = False

'Keep checking the A column to find an empty spot
'if there is no entries in the cell then record the row number
If (Cells(RowNumber, 1) <> "") Then
RowNumber = RowNumber + 1
Else
Placement = True
End If

Loop

Dim CurrentDate As Date
CurrentDate = Date

'now that you know what row is empty, then take the information stored earlier and put it in the cells below
Dim Holder As Variant
Holder = Task & " - part " & i & " of " & Chunk
DueDate = StartDate + i * TimePerChunk
Cells(RowNumber, 1) = Holder ' taskname
Cells(RowNumber, 2) = Name
Cells(RowNumber, 3) = Duration
Cells(RowNumber, 4) = StartDate
Cells(RowNumber, 5) = DueDate
Cells(RowNumber, 6) = Importance ' depends on num of chunks

'make a "score", high importance tasks are given priority
Score = Importance * Duration * 10 * (CurrentDate - DueDate)
Cells(RowNumber, 7) = Score

ActiveCell.NumberFormat = "@"
Next i

Sheets("EqualList").Select
Cells.Select
Selection.Copy
Sheets("UnequalList").Select
Range("A1").Select
ActiveSheet.Paste

Sheets("Menu").Select
End Sub
