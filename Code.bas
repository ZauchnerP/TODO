Public Const DELETE_VALUE As String = "x"
Public Const FOCUS_VALUE As String = "check" ' TODO insert description to README

Public Const COLOR_BG As Long = 16777215 ' White
Public Const COLOR_TEXTRANGE As Long = 0 ' Black
Public Const COLOR_BG_UNIMP As Long = 16316664 ' Light grey
Public Const COLOR_TEXTRANGE_UNIMP As Long = 6579300 ' Medium grey
Public Const COLOR_BUTTON_BG_COLOR As Long = 14474460  ' 14474460 light grey ' 15129810 light blue
Public Const COLOR_BUTTON_TEXTFRAME As Long = 0 ' Black
Public Const COLOR_BUTTON_TEXTRANGE As Long = 0 ' Black
Public Const COLOR_LINES_SUB As Long = 11842740 ' Light grey
Public Const COLOR_LINES_MAIN As Long = 0 ' Black

' To-do (T) sheet constants
Public Const MAX_COL_LETTER As String = "J"  ' "Where"
Public Const MAX_COL As Long = 10 ' Column J

Public Const COL_DEL As Long = 1
Public Const COL_CATEGORY As Long = 2
Public Const COL_IMP As Long = 3
Public Const COL_TIME As Long = 4
Public Const COL_EMOTION As Long = 5
Public Const COL_DEP As Long = 6
Public Const COL_TASK As Long = 7
Public Const COL_WHEN As Long = 8
Public Const COL_HIDE As Long = 9
Public Const COL_WHERE As Long = 10

Public Const ROW_CONTENT_START_T As Long = 3
Public Const ROW_HEADER_T As Long = 2 ' Replace

' Buttons in the to-do sheet
Public Const SORT_BUTTON As Long = 2
Public Const HIDE_LOW_BUTTON As Long = 3
Public Const SORT_TIME_BUTTON As Long = 4
Public Const LINES_BUTTON As Long = 5
Public Const HIDE_DEP_BUTTON As Long = 6
Public Const SHOW_ALL_BUTTON As Long = 7
Public Const HIDE_BUTTON As Long = 9  ' Includes set0
Public Const PLUS_1_BUTTON As Long = 10  ' Includes minus 1

' Day (D) sheet constants
Public Const ROW_HEADER_D As Long = 2
Public Const ROW_CONTENT_START_D As Long = 3

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
    Call Set_Background_White

    ' Mncrease first-row height for two-line headers
    ws.Rows("1:1").RowHeight = 36

    ' Freeze the first two rows
    Call Freeze_Header_Row_2

    ' Fill headers
    Call Create_to_do_Header

    ' Add a bottom border after the second row
    Call Set_Header_Border_T

    ' Create buttons
    Call Create_All_Buttons

    ' Save numbers as text in E "Dependence" (needed for sorting)
    ws.Columns(COL_DEP).NumberFormat = "@"

    ' Activate Today-formatting
    Call Today_Red

    ' Add filter (Header in row 2, so filter starts there)
    ws.Range(ws.Cells(ROW_CONTENT_START_T-1, 1), _
             ws.Cells(ROW_CONTENT_START_T-1, MAX_COL)).AutoFilter

End Sub

Sub Create_Day_Sheet()
    '''' Create a sheet for a day ''''

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check if sheet is empty
    If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
        MsgBox "The current sheet is not empty. Please create the Today sheet on an empty sheet."
        Exit Sub
    End If

    ' Make white background
    Call Set_Background_White

    ' Make header
    Call Create_Header_D
    Call Set_Header_Border_D
    Call Freeze_Header_Row

    ' Make button
    Call Create_Clean_Button_D
    Call Create_Time_Button_D
    Call Create_Time_Undo_Button_D

    ' Fill time slots
    Call Create_Time_Slots
    Call Set_Lines_Today

End Sub

Private Sub Create_to_do_Header()
    ''' Fill the headers in row 2 '''

    ' Initialize
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Integer
    Set ws = ActiveSheet

    ' Define the values to write into row 2
    headers = Array( _
                    "Del", _
                    "Category", _
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

    ' Apply smaller font only to "(1 = important)" in column COL_IMP
    With ws.Cells(2, COL_IMP)
        Dim fullText As String
        fullText = .Value

        Dim startPos As Long
        startPos = InStr(fullText, "(")

        If startPos > 0 Then
            With .Characters(Start:=startPos, _
                             Length:=Len("(1 = important)")).Font
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
    ws.Columns(COL_CATEGORY).ColumnWidth = 15  ' Category
    ws.Columns(COL_TASK).ColumnWidth = 60  ' Task
    ws.Columns(COL_DEP).ColumnWidth = 15  ' Dependence

End Sub

Private Sub Create_Header_D()
    ''' Fill the headers in row 1 '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim i As Integer
    Dim headers As Variant

    ' Define headers
    headers = Array("From", _
                    "To", _
                    "Task")

    ' Fill headers
    For i = 0 To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .WrapText = True
        End With
    Next i

    ' Task column is a bit wider
    ws.Columns("C").ColumnWidth = 60

    ' Date is in the previous row
    With ws.Cells(1, 4)
        .Value = "Date:"
        .Font.Bold = True
        .WrapText = True
    End With

End Sub

Private Sub StyleMyShape(shp As Shape)
    ' Button settings

    With shp
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_BUTTON_TEXTRANGE
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .Line.Visible = msoTrue
        .Fill.ForeColor.RGB = COLOR_BUTTON_BG_COLOR

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

Private Sub Set_Header_Border_D()
    ''' Add a border to the header in the TODAY sheet '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetRange As Range

    Set ws = ActiveSheet
    Set targetRange = ws.Range(ws.Cells(ROW_HEADER_D, 1), _
                               ws.Cells(ROW_HEADER_D, 3))

    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_LINES_MAIN
    End With
End Sub

Private Sub Set_Header_Border_T()
    '''Add a bottom border below the header row in the to-do list '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetRange As Range

    Set ws = ActiveSheet
    Set targetRange = ws.Range(ws.Cells(2, 1), _
                      ws.Cells(2, MAX_COL))

    ' Bottom border formatting
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = COLOR_LINES_MAIN
    End With
End Sub

Private Sub Freeze_Header_Row_2()
    ''' Freeze the first two rows in a to-do sheet '''

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
        .FreezePanes = True
    End With
End Sub

Private Sub Freeze_Header_Row()
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
    Set targetCell = ws.Cells(1, SORT_BUTTON)

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                targetCell.Left, _
                targetCell.Top, _
                targetCell.Width, _
                targetCell.Height)

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
    Set targetCell = ws.Cells(1, HIDE_LOW_BUTTON)

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)

    With shp
        .Name = "Hide_Low"
        .TextFrame2.TextRange.Text = "hide low"
        .OnAction = "Hide_Low"
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
    Set targetCell = ws.Cells(1, SORT_TIME_BUTTON)

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)

    With shp
        .Name = "Sort_Time"
        .TextFrame2.TextRange.Text = "sort" & vbLf & "time"
        .OnAction = "Sort_Time"
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
    Set targetCell = ws.Cells(1, LINES_BUTTON)

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)

    With shp
        .Name = "Make_Lines_TO_DO"
        .TextFrame2.TextRange.Text = "lines"
        .OnAction = "Make_Lines_TO_DO"
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
    Set targetCell = ws.Cells(1, HIDE_DEP_BUTTON)

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)

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
    Set targetCell = ws.Cells(1, SHOW_ALL_BUTTON)

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)

    With shp
        .Name = "Show_All"
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .TextFrame2.TextRange.Text = "show all"
        .OnAction = "Reset_Filters"

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
    Set cell = ws.Cells(1, HIDE_BUTTON)

    ' Get cell dimensions
    cellTop = cell.Top
    cellLeft = cell.Left
    cellWidth = cell.Width
    cellHeight = cell.Height
    halfHeight = cellHeight / 2

    ' Create top button
    Set topBtn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        cellLeft, _
        cellTop, _
        cellWidth, _
        halfHeight)

    With topBtn
        .Name = "Hide_Hide"
        .TextFrame2.TextRange.Text = "hide"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_BUTTON_TEXTRANGE
        .Fill.ForeColor.RGB = COLOR_BUTTON_BG_COLOR
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .OnAction = "Hide_HiddenTasks"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' Create bottom button
    Set bottomBtn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        cellLeft, _
        cellTop + halfHeight, _
        cellWidth, _
        halfHeight)
    With bottomBtn
        .Name = "Set0"
        .TextFrame2.TextRange.Text = "set 0"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_BUTTON_TEXTRANGE
        .Fill.ForeColor.RGB = COLOR_BUTTON_BG_COLOR
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
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
    Dim topBtn As Shape
    Dim bottomBtn As Shape
    Dim cellTop As Double, cellLeft As Double, cellWidth As Double, cellHeight As Double
    Dim halfWidth As Double

    Set ws = ActiveSheet
    Set cell = ws.Cells(1, PLUS_1_BUTTON)

    ' Get cell dimensions
    cellTop = cell.Top
    cellLeft = cell.Left
    cellWidth = cell.Width
    cellHeight = cell.Height
    halfHeight = cellHeight / 2

    ' Create top button (Plus_1)
    Set topBtn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        cellLeft, _
        cellTop, _
        cellWidth, _
        halfHeight)

    With topBtn
        .Name = "Plus_1_Button"
        .TextFrame2.TextRange.Text = "plus 1"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_BUTTON_TEXTRANGE
        .Fill.ForeColor.RGB = COLOR_BUTTON_BG_COLOR
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .OnAction = "Plus_One"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' Create bottom button
    Set bottomBtn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        cellLeft, _
        cellTop + halfHeight, _
        cellWidth, _
        halfHeight)

    With bottomBtn
        .Name = "Minus_1_Button"
        .TextFrame2.TextRange.Text = "minus 1"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_BUTTON_TEXTRANGE
        .Fill.ForeColor.RGB = COLOR_BUTTON_BG_COLOR
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .OnAction = "Minus_One"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

End Sub

' TODAY buttons
Private Sub Create_Clean_Button_D()
    ''' Create the "clean day" button '''

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("A1:B1")

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)

    With shp
        .Name = "Clean_Today"
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .TextFrame2.TextRange.Text = "clean day"
        .OnAction = "Clean_Today"

    End With

    ' Apply global style
    Call StyleMyShape(shp)
End Sub

Private Sub Create_Time_Button_D()
    ''' Create the "time filter" button '''
    'TODO noch in README eingeben!


    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("C1:C1")

    ' Add a button
    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width / 2, _
        targetCell.Height)

    With shp
        .Name = "Select_Time"
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .TextFrame2.TextRange.Text = "time filter"
        .OnAction = "Select_Time"

    End With

    ' Apply global style
    Call StyleMyShape(shp)
End Sub

Private Sub Create_Time_Undo_Button_D()
    ''' Create the "Reset time filter" button '''
    'TODO noch in README eingeben! Besseren Namen finden

    ' Initialize
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape
    Dim baseBtn As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("C1:C1")
    Set baseBtn = ws.Shapes("Select_Time")

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        baseBtn.Left + baseBtn.Width, _
        targetCell.Top, _
        targetCell.Width / 4, _
        targetCell.Height)

    With shp
        .Name = "Undo_Time_Filter"
        .Line.ForeColor.RGB = COLOR_BUTTON_TEXTFRAME
        .TextFrame2.TextRange.Text = "filter off" ' TODO find better name
        .AlternativeText = "Reset time filter"  ' TODO does not work yet
        .OnAction = "Undo_Time_Filter"
    End With

    ' Apply global style
    Call StyleMyShape(shp)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Actions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Main_Sort()
    ''' Main function to sort the document.
    ' This function will sort the data in the worksheet based on columns B, C, D, and E.
    ' Is run by "sort document" button.
    '''

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp

    ' Fill column E
    Call Replace_Empty_Dependence

    ' Fill column H
    Call Insert_0_Hide

    ' Sort the to-do list
    Call Sort_To_Do

    ' Colors
    ' Call Color_WhiteFirst  ' TODO repair
    Call Color_Category
    Call Color_Importance_Time
    Call Importance_Zero

    ' Delete rows
    Call Delete_Rows

    ' Focus rows
    Call Focus_Rows

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Private Sub Delete_Rows()
    ''' Delete rows with DELETE_VALUE in column A '''

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet

    ' Find last used row
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).Row

    ' Loop from bottom to top
    For r = lastRow To 1 Step -1

        If LCase(Trim(ws.Cells(r, 1).Value)) = DELETE_VALUE Then
            ws.Rows(r).Delete Shift:=xlUp
        End If

    Next r

End Sub

Private Sub Focus_Rows()
    ''' Hide all rows if column A contains values with FOCUS_VALUE,
    ' and show only those with FOCUS_VALUE in column A
    '''

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet

    ' Find last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).Row
    Debug.Print "Last row in Focus_Rows: " & lastRow

    ' Check if any cell in column A contains FOCUS_VALUE
    Dim hasFocus As Boolean
    hasFocus = False
    For r = 1 To lastRow
        If LCase(Trim(ws.Cells(r, 1).Value)) = LCase(FOCUS_VALUE) Then
            hasFocus = True
            Exit For
        End If
    Next r

    If hasFocus Then
        ' Hide all rows first
        ws.Rows("3:" & lastRow).Hidden = True

        ' Unhide rows with FOCUS_VALUE in column A
        For r = 1 To lastRow
            If LCase(Trim(ws.Cells(r, 1).Value)) = LCase(FOCUS_VALUE) Then
                ws.Rows(r).Hidden = False
            End If
        Next r
    Else
        ws.Rows.Hidden = False
    End If


End Sub



Private Sub Sort_To_Do()
    ''' Sort the sheet by columns importance, time, emotion, dependence, ascending.
    '''

    Debug.Print "Sorting..."

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    ' Clear existing sort fields
    ws.Sort.SortFields.Clear

    ' Sort by dependence
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_DEP), _
                    ws.Cells(lastRow, COL_DEP)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal


    ' Then by importance
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_IMP), _
                    ws.Cells(lastRow, COL_IMP)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by time
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_TIME), _
                    ws.Cells(lastRow, COL_TIME)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by emotion
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_EMOTION), _
                    ws.Cells(lastRow, COL_EMOTION)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Configure and apply the sort operation
    With ws.Sort
        .SetRange ws.Range(ws.Cells(ROW_HEADER_T, 1), _
                  ws.Cells(lastRow, MAX_COL))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin

        .Apply

    End With

End Sub

Private Sub Hide_Dependence()
    ''' Hide all rows that are dependent on another action
    ' ("Dependence" column)
    '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    ' Filter
    With ws.Range( _
         ws.Cells(ROW_HEADER_T, 1), _
         ws.Cells(lastRow, MAX_COL))
        .AutoFilter Field:=COL_DEP , _
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
    lastRow = ws.Cells(ws.Rows.Count, COL_CATEGORY).End(xlUp).row

    ' Set to white first
    ws.Columns(COL_IMP).Interior.Color = COLOR_BG
    ws.Columns(COL_TIME).Interior.Color = COLOR_BG

    ''' COLOR column B: Importance ''''
    Set rng = ws.Range( _
        ws.Cells(ROW_CONTENT_START_T, COL_IMP), _
        ws.Cells(lastRow, COL_IMP))

    ' Loop through each cell in the range
    For Each cell In rng

        If ws.Cells(cell.row, COL_DEP).Value = "." Then  ' No dependence

            If cell.Value = 1 Then
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow

            ElseIf cell.Value = 2 Then
                cell.Interior.Color = RGB(255, 100, 100)  ' Other color

            ElseIf cell.Value > 2 Then
                cell.Interior.Color = RGB(255, 255, 255)  ' White

            Else
                ' Clear the interior color (not needed because of interdepence with other functions)
                ' cell.Interior.Color = RGB(255, 255, 255)  ' White
            End If
        End If

    Next cell

    ''' COLOR column C: Time ''''
    ' Time is just colorized if it does not take too much emotional effort
    Set rng = ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_TIME), _
                       ws.Cells(lastRow, COL_TIME))

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
    ''' Hide tasks with low importance (>=100) '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    ' Apply filter starting at row 3, column 2
    ' TODO Why does this work although there is a 1?
    ws.Range(ws.Cells(ROW_CONTENT_START_T, 1), _
             ws.Cells(lastRow, MAX_COL)).AutoFilter _
        Field:=COL_IMP, _
            Criteria1:="<100", _
        Operator:=xlAnd, _
            Criteria2:="<>0"

End Sub

Private Sub Make_Lines_TO_DO()
    ''' Add dotted bottom borders to all table rows '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lastRowDelete As Long
    Dim rng As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row
    lastRowDelete = lastRow + 15  ' You can change this number, its just a very conservative assumption of deleted tasks within a short time frame.

    ' Clear all bottom borders in the target range
    For r = ROW_CONTENT_START_T To lastRowDelete
        ws.Range(ws.Cells(r, 1), _
        ws.Cells(r, MAX_COL)).Borders(xlEdgeBottom).LineStyle = xlNone
    Next r

    ' Add borders only to non-empty rows
    For r = ROW_CONTENT_START_T To lastRow
        If Application.WorksheetFunction.CountA( _
            ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))) > 0 Then
            Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))
            With rng.Borders(xlEdgeBottom)
                .LineStyle = xlDot
                .Weight = xlThin
                .Color = COLOR_LINES_SUB
            End With
        End If
    Next r
End Sub

Private Sub Set_Lines_Today()
    ''' Clear existing bottom borders and reapply dotted grey ones for non-empty rows '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim lastRowDelete As Long
    Dim rng As Range

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, COL_CATEGORY).End(xlUp).row
    lastRowDelete = lastRow + 15  ' You can change this number, its just a very conservative assumption of deleted tasks within a short time frame.

    ' Clear all bottom borders in the target range
    For r = ROW_CONTENT_START_D To lastRowDelete
        ws.Range("A" & r & ":C" & r).Borders(xlEdgeBottom).LineStyle = xlNone
    Next r

    ' Add borders only to non-empty rows
    For r = ROW_CONTENT_START_D To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("A" & r & ":C" & r)) > 0 Then
            Set rng = ws.Range("A" & r & ":C" & r)
            With rng.Borders(xlEdgeBottom)
                .LineStyle = xlDot
                .Weight = xlThin
                .Color = COLOR_LINES_SUB
            End With
        End If
    Next r
End Sub

Private Sub Sort_Time()
    ''' Sort the table by the Time column '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    ' Clear any existing sort fields to start fresh
    ws.Sort.SortFields.Clear

    ' Add a sort field for column C (time)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_TIME), _
              ws.Cells(lastRow, COL_TIME)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

    ' Configure and apply the sort operation
    With ws.Sort
        .SetRange ws.Range(ws.Cells(ROW_CONTENT_START_T, 1), _
                           ws.Cells(lastRow, MAX_COL))
        .Header = xlYes
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

Private Sub Hide_HiddenTasks()
    ''' Hide all rows that have the value 1
    ' in the "Hide" column
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    With ws.Range( _
        ws.Cells(ROW_HEADER_T, 1), _
        ws.Cells(lastRow, MAX_COL))
        .AutoFilter _
        Field:=COL_HIDE, _
        Criteria1:="<>1"
    End With
End Sub

Private Sub Set_Hide_0()
    ''' Set all values in the "Hide" column to 0.
    ' Caution: this only sets values in the "Hide" column if the row is not hidden.
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    ' Overwrite "Hide" column with 0
    ws.Range(ws.Cells(ROW_CONTENT_START_T, COL_HIDE), _
             ws.Cells(lastRow, COL_HIDE)).Value = "0"
End Sub

Private Sub Color_Category()
    ''' Colorize the rows depending on the categories  '''

    ' Initialize
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row
    Set rng = Range("A1:A" & lastRow)

    ' Reset colors first
    rng.Interior.Color = COLOR_BG

    ' Loop through each cell in the range
    For Each cell In rng

        Select Case Trim(cell.Value)

            Case "Topic1"
                cell.Interior.Color = RGB(255, 255, 0)   ' Yellow

            Case "Topic2"
                cell.Interior.Color = RGB(255, 100, 100) ' Red

            Case "Topic3"
                cell.Interior.Color = RGB(100, 255, 255) ' Turquoise

        End Select

    Next cell

End Sub

Private Sub Replace_Empty_Dependence()
    ''' Fills column "Dependence" with "." wherever column "Category" has a value.
    ' We need this for sorting the column.
    '''

    ' Initialize
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_CATEGORY).End(xlUp).row

    For r = ROW_CONTENT_START_T To lastRow
        If Trim(ws.Cells(r, COL_TASK).Value) <> "" And Trim(ws.Cells(r, COL_DEP).Value) = "" Then
            ws.Cells(r, COL_DEP).Value = "."
        End If
    Next r
End Sub

Private Sub Insert_0_Hide()
    ''' Fills column "Hide" with 0
    ' where the category column is not empty
    '''

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    For r = ROW_CONTENT_START_T To lastRow
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

    With ActiveSheet.Columns(COL_WHEN)
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
    lastRow = ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).row

    For r = ROW_CONTENT_START_T To lastRow
        If ws.Cells(r, COL_IMP).Value = 0 And _
            ws.Cells(r, COL_IMP).Text <> "" Then

            With ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_COL))
                .Interior.Color = COLOR_BG_UNIMP
                .Font.Color = COLOR_TEXTRANGE_UNIMP
            End With
        Else
            With ws.Range(ws.Cells(r, 1), _
                 ws.Cells(r, MAX_COL))

                .Interior.Color = COLOR_BG
                .Font.Color = COLOR_TEXTRANGE
            End With
        End If
    Next r

End Sub

Private Sub Set_Background_White()
    ''' Set background color of all used cells to white '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Apply white background to entire used range
    ws.Cells.Interior.Color = COLOR_BG
End Sub

Private Sub Clean_Today()
    ''' Clean the "Today" sheet.
    ' Removes all entries and formatting from the "Today" sheet.
    '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' White background
    Call Set_Background_White

    ' Black font
    With ws.Rows(ROW_HEADER_D + 1 & ":38").Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
    End With

    ' Clean content
    ws.Range("C" & ROW_HEADER_D + 1 & " :C38").ClearContents

    ' Delete date
    ws.Range("E1").ClearContents

    ' Delete time
    ws.Range("C1").ClearContents

    ' Unhide everything
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False

    ' Fill time and draw lines again
    Call Create_Time_Slots
    Call Set_Lines_Today

End Sub

Private Sub Select_Time()
    ''' Show rows values from now onwards in the DAY sheet.
    ' Hides all rows for vergangene hours '''

    ' Initialize
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim nowTime As Date
    Dim lastRow As Long
    Dim r As Long
    Dim cellTime As Variant

    ' Show the current time in C1
    nowTime = Time
    With ws.Range("C1")
        .Value = Format(nowTime, "HH:MM")
        .HorizontalAlignment = xlRight
    End With

    ' Hide all previous times
    ' Find last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    For r = ROW_CONTENT_START_D To lastRow

        cellTime = ws.Cells(r, "B").Value

        If IsNumeric(cellTime) Then
            If cellTime < TimeValue(nowTime) Then
                ws.Rows(r).Hidden = True
            Else
                ws.Rows(r).Hidden = False
            End If
        End If

    Next r
    Application.ScreenUpdating = True

End Sub

Private Sub Undo_Time_Filter()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Rows.Hidden = False
    ws.Range("C1").ClearContents
End Sub

Private Sub Create_Time_Slots()
    ''' Fill time slots in the "Today" sheet '''

    ' Initialize
    Dim ws As Worksheet
    Dim startTimeA As Date
    Dim startTimeB As Date
    Dim row As Long

    Set ws = ActiveSheet
    startTimeA = TimeValue("08:00")
    startTimeB = TimeValue("08:30")
    row = ROW_CONTENT_START_D

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
    lastRow = ws.Cells(ws.Rows.Count, COL_CATEGORY).End(xlUp).Row


    For r = ROW_CONTENT_START_T To lastRow
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
    lastRow = ws.Cells(ws.Rows.Count, COL_CATEGORY).End(xlUp).row

    For r = ROW_CONTENT_START_T To lastRow
        val = ws.Cells(r, COL_IMP).Value

        If IsNumeric(val) And val <> 1 And val <> 0 Then
            ws.Cells(r, COL_IMP).Value = val - 1
        End If
    Next r
End Sub


