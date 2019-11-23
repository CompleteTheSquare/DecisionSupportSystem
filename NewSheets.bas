'majority of code from collaborator Niko T-M

Sub newSheets()
    
    'Dates for the begining and end of the project
    Dim StartDate As Date
    Dim endDate As Date
    
    'variable for the month # for use in calendar later
    Dim monthCount As Integer
    
    'variable for year for use in calendar
    Dim yearCount As Integer
    
    'variable to iterate throught weeks for filling calendar
    Dim weekCount As Integer
    
    'variable for days in a given month for filling calendar
    Dim monthDays As Integer
    
    'integer representing the first day of a given month
    Dim firstDayOfMonth As Integer
    
    
    'Number of sheets that must be created (number of months in between the start and end date
    Dim numSheets As Integer
    
    Sheets("Menu").Activate
    
    'Accepts input for the start/end Dates
    endDate = CDate(Range("D13"))
    StartDate = CDate(Range("D12"))
    
    'assigns an initial value to month as the start month
    monthCount = Month(StartDate)
    yearCount = Year(StartDate)
    
    'Prints start/end dates for debugging purposes
    'Cells(1, 2) = FormatDateTime(StartDate, 2)
    'Cells(1, 3) = FormatDateTime(endDate, 2)
    
    'assigns the number of months between start and end as the numSheets
    numSheets = DateDiff("m", StartDate, endDate) + 1
    
    ' Creates the appropriate number of copies of the Calendar Template
    For i = 1 To numSheets
        
        If monthCount > 12 Then
            monthCount = monthCount - 12
            yearCount = yearCount + 1
        End If
        
        'Copies sheet and gives the correct Name for the title
        Worksheets("Calendar Template").Copy After:=Worksheets(i)
        Sheets(1 + i).Name = MonthName(monthCount) & " " & yearCount

        'Creates correct month title in the calendar on the sheet above ^
        Sheets(1 + i).Cells(2, 2) = MonthName(monthCount) & yearCount
        
        'calculates how many days in the month
        If (monthCount = 2) Then
            'checks if leap year to decide how many days february should have
            If (yearCount Mod 4 = 0) And ((yearCount Mod 100 <> 0) Or (yearCount Mod 400)) Then
                monthDays = 29
            Else
                monthDays = 28
            End If
        ElseIf (monthCount = 9 Or monthCount = 4 Or monthCount = 11 Or monthCount = 6) Then
            monthDays = 30
        Else
            monthDays = 31
        End If
            
            
        'finds first day of given month
        firstDayOfMonth = Weekday(DateSerial(yearCount, monthCount, 1), vbSunday)
        
        
        'loop to fill calendar with days
        weekCount = 0
        For j = 1 To monthDays

            Cells(4 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7)) = j

            If (j + firstDayOfMonth - 1) Mod 7 = 0 Then
                weekCount = weekCount + 1

            End If
            
        Next j
        'iterates to the next month
        monthCount = monthCount + 1
        
    Next i

End Sub

Sub addTasks()
    
    'variable for the name of tne task
    Dim taskName As String
    
    'variable for the date of the task
    Dim taskDate As Date
    
    'Dates for the beggining and end of the project
    Dim StartDate As Date
    Dim endDate As Date
    
    'variable for the sheet number of the task datexsss
    Dim sheetNum As Integer
    
    'integer to shw\ow which day of the week the month starts on
    Dim firstDayOfMonth As Integer
    
    'integer representing the number of tasks
    Dim numTasks As Integer
    
    'variable for the name of a group member
    Dim groupMember As String
    
    
    Sheets("Menu").Activate
    
    'Accepts input for the start/end Dates
    endDate = CDate(Range("D13"))
    StartDate = CDate(Range("D12"))
    
    numTasks = WorksheetFunction.CountA([K5:K1000])
    
    
    'inputs tasks into the calendar
    For i = 1 To numTasks
        
        Sheets("Menu").Activate
        
        'assigns value for the date of the task
        taskName = Range("K" & (i + 4))
        
        'assigns value for the date of the task
        taskDate = CDate(Range("L" & (i + 4)))
        
        groupMember = Range("J" & (i + 4))

        'Check if taskdate is within valid range
        If (StartDate < taskDate And taskDate < endDate) Then
        
            'find the correct sheet for the month of the task and activate it
            sheetNum = DateDiff("m", StartDate, taskDate) + 2
            Worksheets(sheetNum).Activate
            
            'find first day of the task's month
            firstDayOfMonth = Weekday(DateSerial(Year(taskDate), Month(taskDate), 1), vbSunday)
            
            'iterate through days of month to find the row of the task date
            weekCount = 0
            For j = 0 To Day(taskDate)
                If j + firstDayOfMonth = 2 Then
               ElseIf (j + firstDayOfMonth - 2) Mod 7 = 0 Then
                    weekCount = weekCount + 1
                End If
            Next j
            
            
            'nested if statement so that task is printed on top of another task
            If Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
                If Cells(6 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
                    If Cells(7 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
                        If Cells(8 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
                            Cells(9 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = taskName & " - " & groupMember
                        Else
                            Cells(8 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = taskName & " - " & groupMember
                        End If
                    Else
                        Cells(7 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = taskName & " - " & groupMember
                    End If
                Else
                    Cells(6 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = taskName & " - " & groupMember
                End If
            Else
                Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = taskName & " - " & groupMember
            End If
            
        Else
            MsgBox "task date must be between the start date and the end date of the project"
        End If
    
    Next i
    
    
    'input final due date into calendar
    Sheets("Menu").Activate

    'find the correct sheet for the final month of the task and activate it
    sheetNum = DateDiff("m", StartDate, endDate) + 2
    Worksheets(sheetNum).Activate

    'find first day of the final month
    firstDayOfMonth = Weekday(DateSerial(Year(endDate), Month(endDate), 1), vbSunday)

    'iterate through days of final month to find the row of the end date
    weekCount = 0
    For j = 0 To Day(endDate)
        If j + firstDayOfMonth = 2 Then
        ElseIf (j + firstDayOfMonth - 2) Mod 7 = 0 Then
            weekCount = weekCount + 1

        End If
    Next j

    'insert final due date into calendar and makes it RED!
    If Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
        If Cells(6 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
            If Cells(7 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
                If Cells(8 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) <> "" Then
                    Cells(9 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = "Final Due Date"
                    Cells(9 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1).Interior.ColorIndex = 3
                Else
                    Cells(8 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = "Final Due Date"
                    Cells(8 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1).Interior.ColorIndex = 3
                End If
            Else
                Cells(7 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = "Final Due Date"
                Cells(7 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1).Interior.ColorIndex = 3
            End If
        Else
            Cells(6 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = "Final Due Date"
            Cells(6 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1).Interior.ColorIndex = 3
        End If
    Else
        Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = "Final Due Date"
        Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1).Interior.ColorIndex = 3
    End If



    'insert final due date into calendar and makes it RED!
    'Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1) = "Final Due Date"
    'Cells(5 + weekCount * 6, (j + (firstDayOfMonth) - weekCount * 7) - 1).Interior.ColorIndex = 3
    
    
End Sub

