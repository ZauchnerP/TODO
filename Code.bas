Public Const MAX_COL_LETTER As String = "I"  ' "Where"
Public Const MAX_COL As Long = 9 ' Column I

Public Const COL_CATEGORY As Long = 1
Public Const COL_IMP As Long = 2
Public Const COL_TIME As Long = 3
Public Const COL_EMOTION As Long = 4
Public Const COL_DEPENDENCE As Long = 5
Public Const COL_TASK As Long = 6
Public Const COL_WHEN As Long = 7
Public Const COL_HIDE As Long = 8
Public Const COL_WHERE As Long = 9

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Create_To_Do_Sheet()
    ''' Create the to-do list. '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check if sheet is empty
    If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
        MsgBox "The current sheet is not empty. Please create the to-do list on an empty sheet."
        Exit Sub
    End If

    ' Make white background
    Call Background_White

    ' Mncrease first-row height for two-line headers
    ws.Rows("1:1").RowHeight = 36

    ' Freeze the first two rows
    Call Freeze_R_1_2

    ' Fill headers
    Call Create_to_do_Header

    ' Add a bottom border after the second row
    Call AddBottomBorderAfterRow2

    ' Create buttons
    Call Create_All_Buttons

    ' Save numbers as text in E "Dependence" (needed for sorting)
    ws.Columns(COL_DEPENDENCE).NumberFormat = "@"

    ' Activate Today-formatting
    Call Today_Red

    ' Add filter
    ws.Range(ws.Cells(2, 1), ws.Cells(2, MAX_COL)).AutoFilter

End Sub

Sub Create_Today_sheet()
    '''' Create a sheet for today ''''

    ' Make white background
    Call Background_White

    ' Make header
    Call Create_Today_Header
    Call AddBottomBorderAfterRow1
    Call Freeze_R_1

    ' Make button
    Call Create_Clean_Today_Button

    ' Fill time slots
    Call Fill_Time_Slots
    Call Make_Lines_Today


End Sub

Private Sub Create_to_do_Header()
    ''' Fill the headers in row 2 '''

    ' Initialize
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Integer
    Set ws = ActiveSheet

    ' Define the values to write into row 2
    headers = Array("Category", _
                    "Importance" & vbLf & "(1 = important)", _
                    "Time" & vbLf & "needed", _
                    "Emotional" & vbLf & "effort", _
                    "Dependence", _
                    "Task", _
                    "When", _
                    "Hide", _ 
                    "Where")


    For i = 0 To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .WrapText = True
        End With
    Next i

    ' Auto-fit row height to handle line breaks
    ws.Rows(2).EntireRow.AutoFit

    ' Apply smaller font only to "(1 = important)" in column B
    With ws.Cells(2, COL_IMP)
        Dim fullText As String
        fullText = .Value

        Dim startPos As Long
        startPos = InStr(fullText, "(")

        If startPos > 0 Then
            With .Characters(Start:=startPos, Length:=Len("(1 = important)")).Font
                .Size = 8
                .Bold = False ' optional: keep it not bold
            End With
        End If
    End With

    ' Fit width of columns A to max column
    Dim col As Range
    For Each col In ws.Range("A:" & MAX_COL_LETTER).Columns
        col.AutoFit
        col.ColumnWidth = col.ColumnWidth + 3
    Next col

    ' Some columns are a bit wider
    ws.Columns("A").ColumnWidth = 15  ' Category
    ws.Columns("F").ColumnWidth = 60  ' Task
    ws.Columns("E").ColumnWidth = 15  ' Dependence

End Sub

Private Sub Create_Today_Header()
    ''' Fill the headers in row 1 '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Define the values to write into row 1
    Dim headers As Variant
    headers = Array("From", _
                    "To", _
                    "Task", _
                    "Date:")

    Dim i As Integer
    For i = 0 To UBound(headers)
        With ws.Cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .WrapText = True
        End With
    Next i

    ' Task column is a bit wider
    ws.Columns("C").ColumnWidth = 60

End Sub

Private Sub StyleMyShape(shp As Shape)
    ' Button settings

    With shp
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' black text
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' black border
        .Line.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(220, 220, 220) ' light grey

        With .TextFrame2
            .MarginTop = 0
            .MarginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With

    End With

End Sub

Private Sub AddBottomBorderAfterRow1()
    ''' Add a border to the header in the TODAY sheet '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetRange As Range

    Set ws = ActiveSheet
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, 3))

    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0) ' black
    End With
End Sub

Private Sub AddBottomBorderAfterRow2()
    '''Add a bottom border below the second row (the header) in the to-do list '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetRange As Range

    Set ws = ActiveSheet
    Set targetRange = ws.Range(ws.Cells(2, 1), ws.Cells(2, MAX_COL))

    ' Bottom border formatting
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0) ' black
    End With
End Sub

Private Sub Freeze_R_1_2()
    ''' Freeze the first two rows in a to-do sheet '''

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
        .FreezePanes = True
    End With
End Sub

Private Sub Freeze_R_1()
    ''' Freeze the first row in a "Today" sheet'''

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Buttons
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Create_All_Buttons()
    '''' Create all buttons in the worksheet '''

    Call Create_Sort_All_Button
    Call Create_Hide_Low_Button
    Call Create_Sort_Time_Button
    Call Create_Lines_Button
    Call Create_Hide_Dependence_Button
    Call Create_Show_All_Button
    Call Create_Hide_Buttons
    Call Create_MinusPlus_1_Buttons

End Sub

Private Sub Create_Sort_All_Button()
    ''' Create the "sort all" button '''
    ' This button will sort the data in the worksheet based on columns B, C, D, and E

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("A1")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Sort_All"
        .TextFrame2.TextRange.Text = "sort" & vbLf & "document"
        .OnAction = "Main_Sort"
    End With

    ' Apply style
    Call StyleMyShape(shp)
    shp.Fill.ForeColor.RGB = RGB(0, 176, 80) ' green

End Sub

Private Sub Create_Hide_Low_Button()
    ''' Create the "hide low" button '''
    ' This button hides all tasks that are considered less important, meaning those with a value below 100.
    ' You can change this threshold in Hide_Low().

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("B1")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Hide_Low"
        .TextFrame2.TextRange.Text = "hide low"
        .OnAction = "Hide_Low"
    End With

    ' Apply global style
    Call StyleMyShape(shp)

End Sub

Private Sub Create_Lines_Button()
    ''' Create the "lines" button.
    ' This button will make a dotted line between the tasks
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("D1")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Make_Lines_TO_DO"
        .TextFrame2.TextRange.Text = "lines"
        .OnAction = "Make_Lines_TO_DO"
    End With

    ' Apply global style
    Call StyleMyShape(shp)

End Sub

Private Sub Create_Sort_Time_Button()
    ''' Create the "sort time" button.
    ' This button will sort the data in the worksheet based on column C (time)
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("C1")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Sort_Time"
        .TextFrame2.TextRange.Text = "sort" & vbLf & "time"
        .OnAction = "Sort_Time"
    End With

    ' Apply global style
    Call StyleMyShape(shp)
End Sub

Private Sub Create_Hide_Dependence_Button()
    ''' Create the "hide dependence" button '''
    ' This button will hide all rows that are dependent on another action (column E)

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("E1")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Hide_Dependence"
        .TextFrame2.TextRange.Text = "hide" & vbLf & "dependence"
        .OnAction = "Hide_Dependence"
    End With

    ' Apply global style
    Call StyleMyShape(shp)

End Sub

Private Sub Create_Show_All_Button()
    ''' Create the "show all" button '''
    ' This button will reset all filters in the worksheet

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("F1")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Show_All"
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' black border
        .TextFrame2.TextRange.Text = "show all"
        .OnAction = "Reset_Filters"

    End With

    Call StyleMyShape(shp) ' Apply global style
End Sub

Private Sub Create_Clean_Today_Button()
    ''' Create the "clean today" button '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("D2:E2")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Clean_Today"
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' black border
        .TextFrame2.TextRange.Text = "clean today"
        .OnAction = "Clean_Today"

    End With

    ' Apply global style
    Call StyleMyShape(shp) 
End Sub

Private Sub Create_Hide_Buttons()
    ''' Create the "hide" and "set 0" buttons '''

    ' Initialize
    Dim ws As Worksheet
    Dim cell As Range
    Dim topBtn As Shape, bottomBtn As Shape
    Dim cellTop As Double, cellLeft As Double, cellWidth As Double, cellHeight As Double
    Dim halfHeight As Double

    Set ws = ActiveSheet
    Set cell = ws.Range("H1")

    ' Get cell dimensions
    cellTop = cell.Top
    cellLeft = cell.Left
    cellWidth = cell.Width
    cellHeight = cell.Height
    halfHeight = cellHeight / 2

    ' Create top button
    Set topBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop, cellWidth, halfHeight)
    With topBtn
        .Name = "Hide_Hide"
        .TextFrame2.TextRange.Text = "hide"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(220, 220, 220)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .OnAction = "Hide"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' Create bottom button
    Set bottomBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop + halfHeight, cellWidth, halfHeight)
    With bottomBtn
        .Name = "Set0"
        .TextFrame2.TextRange.Text = "set 0"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(220, 220, 220)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .OnAction = "Set_Hide_0"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub


Private Sub Create_MinusPlus_1_Buttons()
    ''' Create Minus_1 and Plus_1 buttons that add and subtract from the importance column '''
    ' Initialize
    Dim ws As Worksheet
    Dim cell As Range
    Dim leftBtn As Shape, rightBtn As Shape
    Dim cellTop As Double, cellLeft As Double, cellWidth As Double, cellHeight As Double
    Dim halfWidth As Double

    Set ws = ActiveSheet
    Set cell = ws.Range("I1")

    ' Get cell dimensions
    cellTop = cell.Top
    cellLeft = cell.Left
    cellWidth = cell.Width
    cellHeight = cell.Height
    halfHeight = cellHeight / 2

    ' Create top button (Plus_1)
    Set topBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop, cellWidth, halfHeight)
    With topBtn
        .Name = "Plus_1_Button"
        .TextFrame2.TextRange.Text = "plus 1"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(220, 220, 220)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .OnAction = "Plus_One"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' Create bottom button
    Set bottomBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop + halfHeight, cellWidth, halfHeight)
    With bottomBtn
        .Name = "Minus_1_Button"
        .TextFrame2.TextRange.Text = "minus 1"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(220, 220, 220)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .OnAction = "Minus_One"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Actions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Main_Sort()
    ''' Main function to sort the document.
    ' This function will sort the data in the worksheet based on columns B, C, D, and E
    '''

    ' Fill column E
    Call Replace_Empty_Dependence

    ' Fill column H
    Call Insert_0_Hide

    ' Sort the to-do list
    Call Sort_To_Do

    ' Colors
    Call Importance_Zero
    Call Color_Category
    Call Color_Importance_Time

End Sub

Private Sub Sort_To_Do()
    ''' Sort the sheet by columns B, C, D, and E
    ' (importance, time, emotion, dependence), ascending. 
    '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Clear existing sort fields
    ws.Sort.SortFields.Clear

    ' Sort by E (dependence)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(3, COL_DEPENDENCE), ws.Cells(lastRow, COL_DEPENDENCE)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by column B (importance)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(3, COL_IMP), ws.Cells(lastRow, COL_IMP)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by column C (time)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(3, COL_TIME), ws.Cells(lastRow, COL_TIME)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by column D (emotion)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(3, COL_EMOTION), ws.Cells(lastRow, COL_EMOTION)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Configure and apply the sort operation
    With ws.Sort
        .SetRange ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, MAX_COL))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub Hide_Dependence()
    ''' Hide all rows that are dependent on another action ("Dependence" column)  '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Filter
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, MAX_COL))
        .AutoFilter Field:=COL_DEPENDENCE , _
        Criteria1:="=", _
        Operator:=xlOr, _
        Criteria2:="."
    End With
End Sub

Private Sub Color_Importance_Time()
    ''' 
    ' Colorize column B (importance) and C (time needed) cells if 
    ' a task is important or quick to do.
    ' Both are just colorized if they do not depend on other tasks.
    ' Time is just colorized when the task does not require too much emotional effort.
    '''

    ' Initialize
    Dim ws As Worksheet    
    Set ws = ActiveSheet

    Dim rng As Range
    Dim cell As Range

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Set to white first
    ws.Columns(COL_IMP).Interior.Color = RGB(255, 255, 255)
    ws.Columns(COL_TIME).Interior.Color = RGB(255, 255, 255)

    ''' COLOR column B: Importance ''''
    Set rng = ws.Range(ws.Cells(3, COL_IMP), ws.Cells(lastRow, COL_IMP))

    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Value = 1 And _
            cell.Offset(0, 3).Value = "." Then  ' No dependence
            cell.Interior.Color = RGB(255, 255, 0) ' Yellow

        ElseIf cell.Value = 2 And _
            cell.Offset(0, 3).Value = "." Then  ' No dependence
                    cell.Interior.Color = RGB(255, 100, 100)  ' Other color

        ElseIf cell.Value > 2 And _
            cell.Offset(0, 3).Value = "." Then  ' No dependence
                    cell.Interior.Color = RGB(255, 255, 255)  ' White

        Else
            ' Clear the interior color (not needed because of interdepence with other functions)
            ' cell.Interior.Color = RGB(255, 255, 255)  ' White
        End If

    Next cell

    ''' COLOR column C: Time ''''
    ' Time is just colorized if it does not take too much emotional effort
    Set rng = ws.Range(ws.Cells(3, COL_TIME), ws.Cells(lastRow, COL_TIME))

    ' Loop through each cell in the range
    For Each cell In rng

        If cell.Offset(0, -1).Value <> 0 Then
            If cell.Value <> "" And _
                IsNumeric(cell.Value) And _
                cell.Value < 6 And _
                ((cell.Offset(0, 1).Value <> "" And _
                IsNumeric(cell.Offset(0, 1).Value) And _
                cell.Offset(0, 1).Value < 5) Or _
                cell.Offset(0, 1).Value = "") And _
                cell.Offset(0, 2).Value = "." Then
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow

            ElseIf cell.Value <> "" And _
                IsNumeric(cell.Value) And _
                cell.Value < 11 And _
                ((cell.Offset(0, 1).Value <> "" And _
                IsNumeric(cell.Offset(0, 1).Value) And _
                cell.Offset(0, 1).Value < 5) Or _
                cell.Offset(0, 1).Value = "") And _
                cell.Offset(0, 2).Value = "." Then

                cell.Interior.Color = RGB(255, 100, 100)  ' Red

            Else
                ' Clear the interior color (not needed, because of interdepence with other functions)
                ' cell.Interior.Color = RGB(255, 255, 255)  ' White
            End If
        End If
    Next cell

End Sub

Private Sub Hide_Low()
    ''' Hide tasks with low importance (<100) '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Apply filter starting at row 3, column 2
    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, MAX_COL)).AutoFilter _
            Field:=COL_IMP, _
            Criteria1:="<100", _
            Operator:=xlAnd, _
            Criteria2:="<>0"

End Sub

Private Sub Make_Lines_TO_DO()
    ''' Clear existing bottom borders and reapply dotted grey ones for non-empty rows '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lastRowDelete As Long
    Dim rng As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastRowDelete = lastRow + 15  ' You can change this number, its just a very conservative assumption of deleted tasks within a short time frame.

    ' Clear all bottom borders in the target range
    For r = 3 To lastRowDelete
        ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL)).Borders(xlEdgeBottom).LineStyle = xlNone
    Next r

    ' Add borders only to non-empty rows
    For r = 3 To lastRow
        If Application.WorksheetFunction.CountA( _
            ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))) > 0 Then
            Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))
            With rng.Borders(xlEdgeBottom)
                .LineStyle = xlDot
                .Weight = xlThin
                .Color = RGB(180, 180, 180) ' light grey
            End With
        End If
    Next r
End Sub

Private Sub Make_Lines_Today()
    ''' Clear existing bottom borders and reapply dotted grey ones for non-empty rows '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lastRowDelete As Long
    Dim rng As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastRowDelete = lastRow + 15  ' You can change this number, its just a very conservative assumption of deleted tasks within a short time frame.

    ' Clear all bottom borders in the target range
    For r = 2 To lastRowDelete
        ws.Range("A" & r & ":C" & r).Borders(xlEdgeBottom).LineStyle = xlNone
    Next r

    ' Add borders only to non-empty rows
    For r = 2 To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("A" & r & ":C" & r)) > 0 Then
            Set rng = ws.Range("A" & r & ":C" & r)
            With rng.Borders(xlEdgeBottom)
                .LineStyle = xlDot
                .Weight = xlThin
                .Color = RGB(180, 180, 180) ' light grey
            End With
        End If
    Next r
End Sub

Private Sub Sort_Time()
    ''' Sort the data in the worksheet based on column C (time) '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Clear any existing sort fields to start fresh
    ws.Sort.SortFields.Clear

    ' Add a sort field for column C (time), from row 2 to the last data row
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range("C3:C" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Configure and apply the sort operation
    With ws.Sort
        .SetRange ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, MAX_COL))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Private Sub Reset_Filters()
    ''' Reset all filters in the worksheet without deleting them '''

    ' Initialize
    Dim ws As Worksheet
    Dim i As Integer

    Set ws = ActiveSheet

    ' Check if the worksheet has an AutoFilter
    If ws.AutoFilterMode Then
        ' Loop through each column with a filter
        With ws.AutoFilter
            For i = 1 To .Filters.Count
                ' Check if there is a filter applied and clear it
                If .Filters(i).On Then
                    ws.AutoFilter.Range.AutoFilter Field:=i
                End If
            Next i
        End With
    End If
End Sub

Private Sub Hide()
    ''' Hide all rows that have the value 1 in the "Hide" column '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, MAX_COL))
        .AutoFilter _
            Field:=COL_HIDE, _
            Criteria1:="<>" & 1
    End With
End Sub

Private Sub Set_Hide_0()
    ''' Set all values in the "Hide" column to 0.
    ' Caution: this only sets values in rows 3 to lastRow to 0 if the row is not hidden.
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Overwrite "Hide" column with 0
    ws.Range(ws.Cells(3, COL_HIDE), ws.Cells(lastRow, COL_HIDE)).Value = "0"
End Sub

Private Sub Color_Category()
    ''' Colorize the rows depending on the categories in column A  '''

    ' Initialize
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    Set rng = Range("A1:A" & lastRow)

    ' First white
    ws.Columns("A").Interior.Color = RGB(255, 255, 255)

    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Value = "Topic1" Then   ' TODO change to your category names
            cell.Interior.Color = RGB(255, 255, 0)  ' Yellow

        ElseIf cell.Value = "Topic2" Then ' TODO change to your category names

                    cell.Interior.Color = RGB(255, 100, 100)  ' Red

        ElseIf cell.Value = "Topic3" Then ' TODO change to your category names

                    cell.Interior.Color = RGB(100, 255, 255)  ' Turquoise

        Else
            ' Same color as the cell right to it
            cell.Interior.Color = cell.Offset(0, 1).Interior.Color
        End If

    Next cell

End Sub

Private Sub Replace_Empty_Dependence()
    ''' Fills column E (dependence) with "." wherever column A (category) has a value.
    ' We need this for sorting column E. 
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For r = 3 To lastRow
        If Trim(ws.Cells(r, 1).Value) <> "" And Trim(ws.Cells(r, COL_DEPENDENCE).Value) = "" Then
            ws.Cells(r, COL_DEPENDENCE).Value = "."
        End If
    Next r
End Sub

Private Sub Insert_0_Hide()
    ''' Fills column H ("Hide") with a 0 in every row
    ' where column A ("Category") is not empty.
    ' We need this for sorting column H. """

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For r = 3 To lastRow
        If Trim(ws.Cells(r, 1).Value) <> "" _
            And Trim(ws.Cells(r, COL_HIDE).Value) = "" Then
            ws.Cells(r, COL_HIDE).Value = "0"
        End If
    Next r
End Sub

Private Sub Today_Red()
    ''' Apply conditional formatting to column G ("When") in the active to-do sheet.
    ' After this procedure is run once within Create_To_Do_Sheet, 
    ' any cell in column G that contains today's date will be 
    ' highlighted automatically when entered.
    '''

    ' Initialize
    Dim fc As FormatCondition

    With ActiveSheet.Columns("G")
        .FormatConditions.Delete
        
        Set fc = .FormatConditions.Add( _
            Type:=xlCellValue, _
            Operator:=xlEqual, _
            Formula1:="=" & CLng(Date))

        fc.Font.Color = -16383844
        fc.Interior.Color = 13551615
    End With

End Sub

Private Sub Importance_Zero()
    ''' Colors an entire table row in light grey if the value in the
    ' "Importance" column is equal to 0. '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For r = 3 To lastRow
        If ws.Cells(r, COL_IMP).Value = 0 And ws.Cells(r, COL_IMP).Text <> "" Then
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))
                .Interior.Color = RGB(248, 248, 248) ' Light grey background
                .Font.Color = RGB(100, 100, 100)     ' Medium grey text
            End With
        Else
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))
                .Interior.Color = RGB(255, 255, 255) ' White background
                .Font.Color = RGB(0, 0, 0)     ' Black text
            End With
        End If
    Next r

End Sub

Private Sub Background_White()
    ''' Set background color of all used cells to white '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Apply white background to entire used range
    ws.Cells.Interior.Color = RGB(255, 255, 255)
End Sub

Private Sub Clean_Today()
    ''' Clean the "Today" sheet.
    ' Removes all entries and formatting from the "Today" sheet.
    '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' White background
    Call Background_White
    
    ' Black font
    With ws.Rows("2:37").Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
    End With

    ' Clean content
    ws.Range("C2:C37").ClearContents

    ' Delete date
    ws.Range("E1").ClearContents

    ' Fill time and draw lines again
    Call Fill_Time_Slots
    Call Make_Lines_Today

End Sub

Private Sub Fill_Time_Slots()
    ''' Fill time slots in the "Today" sheet '''

    ' Initialize
    Dim ws As Worksheet
    Dim startTimeA As Date
    Dim startTimeB As Date
    Dim row As Long

    Set ws = ActiveSheet
    startTimeA = TimeValue("08:00")
    startTimeB = TimeValue("08:30")
    row = 2

    Do While startTimeA <= TimeValue("23:40")
        ws.Cells(row, 1).Value = Format(startTimeA, "hh:mm")
        ws.Cells(row, 2).Value = Format(startTimeB, "hh:mm")

        startTimeA = startTimeA + TimeSerial(0, 30, 0)
        startTimeB = startTimeB + TimeSerial(0, 30, 0)
        row = row + 1
    Loop
End Sub

Private Sub Plus_One()
    ''' Add 1 to each cell in column B (importance), excluding cells with a value of 0'''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim val As Variant
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row


    For r = 3 To lastRow
        val = ws.Cells(r, COL_IMP).Value

        If IsNumeric(val) And val <> 0 Then
            ws.Cells(r, COL_IMP).Value = val + 1
        End If
    Next r
End Sub


Private Sub Minus_One()
    ''' Subtract 1 from each cell in column B (importance), excluding cells with a value of 0 or 1 '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim val As Variant  

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row


    For r = 3 To lastRow
        val = ws.Cells(r, COL_IMP).Value

        If IsNumeric(val) And val <> 1 And val <> 0 Then
            ws.Cells(r, COL_IMP).Value = val - 1
        End If
    Next r
End Sub
