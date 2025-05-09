''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Create_TODO_sheet()
    '''' Create the TODO list ''''
    ' this function runs all the functions below. You only need to run this one.

    ' Make white background
    Call Background_White

    ' Make first row to a two-liner
    Rows("1:1").RowHeight = 36

    ' Freeze the first two rows
    Call Freeze_R_1_2

    ' Fill headers
    Call Create_TODO_Header

    ' Add a bottom border after the second row
    Call AddBottomBorderAfterRow2

    ' Create buttons
    Call Create_All_Buttons

    ' Add filter
    Range("A2:H2").Select
    Range("H2").Activate
    Selection.AutoFilter

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

Private Sub Create_TODO_Header()
    ''' Fill the headers in row 2 '''

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Define the values to write into row 2
    Dim headers As Variant
    headers = Array("Category", _
                    "Importance" & vbLf & "(1 = important)", _
                    "Time" & vbLf & "needed", _
                    "Emotional" & vbLf & "effort", _
                    "Dependence", _
                    "Task", _
                    "When", _
                    "Hide")

    Dim i As Integer
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
    With ws.Cells(2, 2)
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

    ' Fit width of columns A to H
    For Each col In ws.Range("A:H").Columns
        col.AutoFit
        col.ColumnWidth = col.ColumnWidth + 3
    Next col

    ' Some columns are a bit wider
    ws.Columns("A").ColumnWidth = 15  ' Category
    ws.Columns("F").ColumnWidth = 60  ' Task
    ws.Columns("E").ColumnWidth = 15  ' Dependence

End Sub

Private Sub Create_Today_Header()
    ''' Fill the headers in row 2 '''

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

Private Sub AddBottomBorderAfterRow2()
    ''' Add a bottom border after second row (header) in the TODO list'''

    Dim ws As Worksheet
    Dim lastCol As Long
    Dim targetRange As Range

    Set ws = ActiveSheet

    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Set targetRange = ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))

    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0) ' black
    End With
End Sub

Private Sub AddBottomBorderAfterRow1()
    ''' Add a border to the header in the TODAY sheet '''

    Dim ws As Worksheet
    Dim lastCol As Long
    Dim targetRange As Range

    Set ws = ActiveSheet

    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, 3))

    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0) ' black
    End With
End Sub

Private Sub Freeze_R_1_2()
    ''' Freeze the first two rows '''

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
        .FreezePanes = True
    End With
End Sub

Private Sub Freeze_R_1()
    ''' Freeze the first row '''

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

' Create buttons

Private Sub Create_All_Buttons()
    '''' Create all buttons in the worksheet '''
    Call Create_Sort_All_Button
    Call Create_Lines_Button
    Call Create_Sort_Time_Button
    Call Create_Hide_Dependence_Button
    Call Create_Show_All_Button
    Call Create_Hide_Buttons

End Sub

Private Sub Create_Sort_All_Button()
    ''' Create the "sort all" button '''
    ' This button will sort the data in the worksheet based on columns B, C, D, and E

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("A1")

    ' Delete existing shape if it exists
    On Error Resume Next
    ws.Shapes("Sort_All").Delete
    On Error GoTo 0

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

Private Sub Create_Lines_Button()
    ''' Create the "lines" button '''
    ' This button will make a dotted line between the tasks

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("B1")

    ' Delete existing shape if it exists
    On Error Resume Next
    ws.Shapes("Make_Lines").Delete
    On Error GoTo 0

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Make_Lines"
        .TextFrame2.TextRange.Text = "lines"
        .OnAction = "Make_Lines"
    End With

    ' Apply global style
    Call StyleMyShape(shp)

End Sub

Private Sub Create_Sort_Time_Button()
    ''' Create the "sort time" button '''
    ' This button will sort the data in the worksheet based on column C (time)

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("C1")

    ' Delete existing shape if it exists
    On Error Resume Next
    ws.Shapes("Sort_Time").Delete
    On Error GoTo 0

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

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("E1")

    ' Delete existing shape if it exists
    On Error Resume Next
    ws.Shapes("Hide_Dependence").Delete
    On Error GoTo 0

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

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("F1")

    ' Delete existing shape if it exists
    On Error Resume Next
    ws.Shapes("Show_All").Delete
    On Error GoTo 0

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

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim shp As Shape

    Set ws = ActiveSheet
    Set targetCell = ws.Range("D2:E2")

    ' Delete existing shape if it exists
    On Error Resume Next
    ws.Shapes("Clean_Today").Delete
    On Error GoTo 0

    ' Add a button
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

    With shp
        .Name = "Clean_Today"
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' black border
        .TextFrame2.TextRange.Text = "clean today"
        .OnAction = "Clean_Today"

    End With

    Call StyleMyShape(shp) ' Apply global style
End Sub

Private Sub Create_Hide_Buttons()
    ''' Create the "hide" and "set 0" buttons '''

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

    ' Delete old buttons if they exist
    On Error Resume Next
    ws.Shapes("Hide_Hide").Delete
    ws.Shapes("Set0").Delete
    On Error GoTo 0

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
        .OnAction = "Hide_Set0"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Actions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Main_Sort()
    ''' Main function to sort the document '''
    ' This function will sort the data in the worksheet based on columns B, C, D, and E

    EnableEvents = True
    Call Replace_Empty_Dependence
    Call Insert_0_Hide
    Call Sort_TODO
    Call Importance_Zero
    Call Color_Category
    Call Color_Importance
    Call Today_Red

End Sub

Private Sub Sort_TODO()
    ''' Sort the data in the worksheet based on columns B C D and I '''
    ' I. e. importance, time, emotion, dependence in ascending order

    ' Set worksheet and lastRow
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Clear existing sort fields
    ws.Sort.SortFields.Clear

    ' Sort by E (dependence)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range("E3:E" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by column B (importance)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range("B3:B" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by column C (time)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range("C3:C" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Then by column D (emotion)
    ws.Sort.SortFields.Add2 _
        Key:=ws.Range("D3:D" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' Configure and apply the sort operation
    With ws.Sort
        .SetRange ws.Range("A3:H" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub Hide_Dependence()
    ''' Hide all rows that are dependent on another action ("Dependence" column)  '''

    ' Set worksheet and lastRow
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Filter
    With ActiveSheet.Range("A2:J" & lastRow)
        .AutoFilter Field:=5, _
        Criteria1:="=", _
        Operator:=xlOr, _
        Criteria2:="."
    End With
End Sub

Private Sub Color_Importance()
    ''' Colorize column B (importance) and C (time needed) cells if a task is important or quick to do'''
    ' Both are just colorized if they do not depend on other tasks.
    ' Time is just colorized when the taask does not require too much emotional effort.

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Set ws = ActiveSheet

    ' Set to white first
    ws.Columns("B").Interior.Color = RGB(255, 255, 255)
    ws.Columns("C").Interior.Color = RGB(255, 255, 255)

    ''' COLOR column B: Importance ''''
    Set rng = ws.Range("B2:B60")

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
            ' Clear the interior color (not needed, because of interdepence with other functions)
            ' cell.Interior.Color = RGB(255, 255, 255)  ' White
        End If

    Next cell

    ''' COLOR column C: Time ''''
    ' Time is just colorized if it does not take too much emotional effort
    Set rng = ws.Range("C2:C60")

    ' Loop through each cell in the range
    For Each cell In rng

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
    Next cell

End Sub

Sub Make_Lines()
    ''' Clear existing bottom borders and reapply dotted grey ones for non-empty rows '''

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim rng As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastRowDelete = lastRow + 15  ' You can change this number, its just a very conservative assumption of deleted tasks within a short time frame.

    ' Step 1: Clear all bottom borders in the target range
    For r = 3 To lastRowDelete
        ws.Range("A" & r & ":H" & r).Borders(xlEdgeBottom).LineStyle = xlNone
    Next r

    ' Step 2: Add borders only to non-empty rows
    For r = 3 To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("A" & r & ":H" & r)) > 0 Then
            Set rng = ws.Range("A" & r & ":H" & r)
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

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim rng As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastRowDelete = lastRow + 15  ' You can change this number, its just a very conservative assumption of deleted tasks within a short time frame.

    ' Step 1: Clear all bottom borders in the target range
    For r = 2 To lastRowDelete
        ws.Range("A" & r & ":C" & r).Borders(xlEdgeBottom).LineStyle = xlNone
    Next r

    ' Step 2: Add borders only to non-empty rows
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

Sub Sort_Time()
    ''' Sort the data in the worksheet based on column C (time) '''

    ' Set worksheet and lastRow
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Select the range (these columns will be sorted)
    ws.Columns("A:H").Select

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
        .SetRange ws.Range("A3:H" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub Reset_Filters()
    ''' Reset all filters in the worksheet without deleting them '''

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

Sub Hide()
    ''' Hide all rows that have the value 0 in the "Hide" column '''

    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row

    With ActiveSheet.Range("A2:H" & lastRow)  ' TODO check if correct
        .AutoFilter _
            Field:=8, _
            Criteria1:="<>=", _
            Operator:=xlAnd, _
            Criteria2:="0"
    End With
End Sub

Sub Hide_Set0()
    ''' Set all 1 in the "Hide" column to 0  '''
    ' Caution: this will only set the value 1 to 0 if it is not hidden '''

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ws.Range("H3:H" & lastRow).Value = "0"
End Sub

Sub Color_Category()
    ''' Colorize the rows depending on the categories in column A  '''

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
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
    ' We need this for sorting column E

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 3 To lastRow
        If Trim(ws.Cells(r, "A").Value) <> "" And Trim(ws.Cells(r, "E").Value) = "" Then
            ws.Cells(r, "E").Value = "."
        End If
    Next r
End Sub

Private Sub Insert_0_Hide()
    ''' Fills column H (hide) with "0" wherever column A (category) has a value.
    ' We need this for sorting this column

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 3 To lastRow
        If Trim(ws.Cells(r, "A").Value) <> "" _
            And Trim(ws.Cells(r, "H").Value) = "" Then
            ws.Cells(r, "H").Value = "0"
        End If
    Next r
End Sub

Private Sub Today_Red()
    ''' Colorize the rows in column G (When) that are equal to today '''

    Dim Today As Date
    Today = Date ' This is needed so that it works with all language settings

    With ActiveSheet.Columns("G")
        ' Set format condition to today
        .FormatConditions.Add _
            Type:=xlCellValue, _
            Operator:=xlEqual, _
            Formula1:=Today

        ' Text color
        With .FormatConditions(.FormatConditions.Count).Font
            .Color = -16383844 ' Dark red color
        End With

        ' Background cell color
        With .FormatConditions(.FormatConditions.Count).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615  ' Light red/pinkish
        End With

    End With
End Sub

Private Sub Importance_Zero()
    ''' Colorize the rows when importance is 0 to light grey '''

    ' Set worksheet and lastRow
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 3 To lastRow
        If ws.Cells(r, "B").Value = 0 And ws.Cells(r, "B").Text <> "" Then
            With ws.Range("A" & r & ":H" & r)
                .Interior.Color = RGB(248, 248, 248) ' Light grey background
                .Font.Color = RGB(100, 100, 100)     ' Medium grey text
            End With
        Else
            With ws.Range("A" & r & ":H" & r)
                .Interior.Color = RGB(255, 255, 255) ' White background
                .Font.Color = RGB(0, 0, 0)     ' Black text
            End With
        End If
    Next r

End Sub

Private Sub Background_White()
    ''' Set background color of all used cells to white '''

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Apply white background to entire used range
    ws.Cells.Interior.Color = RGB(255, 255, 255)
End Sub



Sub Clean_Today()
    ''' Clean the "Today" sheet '''
    ' Removes all entries and formatting from the "Today" sheet

    Rows("2:28").Select

    ' White background
    Call Background_White

    ' Black font
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
    End With

    'Clean content
    Range("C2:C28").Select
    Selection.ClearContents

    ' Delete date
    Range("E1").Select
    Selection.ClearContents

End Sub

Private Sub Fill_Time_Slots()
    ''' Fill time slots in the "Today" sheet '''

    Dim startTimeA As Date
    Dim startTimeB As Date
    Dim row As Long

    startTimeA = TimeValue("08:00")
    startTimeB = TimeValue("08:30")
    row = 2

    Do While startTimeA <= TimeValue("23:40")
        Cells(row, 1).Value = Format(startTimeA, "hh:mm")
        Cells(row, 2).Value = Format(startTimeB, "hh:mm")

        startTimeA = startTimeA + TimeSerial(0, 30, 0)
        startTimeB = startTimeB + TimeSerial(0, 30, 0)
        row = row + 1
    Loop
End Sub



