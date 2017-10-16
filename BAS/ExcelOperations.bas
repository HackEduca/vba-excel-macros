Attribute VB_Name = "ExcelOperations"
' Collection of Excel-centric general applications

Sub VerticalOps()   ' Convert downloaded ops into vertical format

    Dim r, c, OpCount, RowPerOp, i As Integer
    RowPerOp = 5
    c = 1 + RowPerOp
    r = 1
    
    OpCount = Application.WorksheetFunction.RoundUp((ActiveSheet.UsedRange.Columns.Count) / RowPerOp, 0)

    Do While OpCount <> 1
        Do While IsEmpty(Cells(r, 4)) = False Or IsEmpty(Cells(r, 3)) = False Or IsEmpty(Cells(r + 1, 4)) = False
            r = r + 1
        Loop
        i = 1
        Do While IsEmpty(Cells(i, c + 2)) = False Or IsEmpty(Cells(i, c + 3)) = False
            Cells(r + i, 1).Formula = Cells(i, c).Formula
            Cells(r + i, 2).Formula = Cells(i, c + 1).Formula
            Cells(r + i, 3).Formula = Cells(i, c + 2).Formula
            Cells(r + i, 4).Formula = Cells(i, c + 3).Formula
            Cells(r + i, 5).Formula = Cells(i, c + 4).Formula
            i = i + 1
        Loop
        c = c + RowPerOp
        OpCount = OpCount - 1
    Loop
    Range(Cells(1, 6), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).Delete Shift:=xlUp
    MsgBox "Done."
End Sub

Function IsEmptySheet(ByVal shit As Worksheet) As Boolean    ' Checks if the input sheet is empty
    If WorksheetFunction.CountA(shit.Cells) = 0 Then
        IsEmptySheet = True
    Else
        IsEmptySheet = False
    End If
End Function

Sub Concen()    ' Join multi-line strings into SAP Text Editor ready format.
    mystr = ""
    For Each item In Selection
        mystr = mystr & " " & item.text
    Next
    mystr = Trim(mystr)
    LTRearrange mystr, ActiveCell.row, ActiveCell.Column
End Sub

Sub Concen2()   ' Join multi-line strings into a single-line string.
    mystr = ""
    For Each item In Selection
        mystr = mystr & " " & item.text
    Next
    mystr = Trim(mystr)
    WriteIntoCell mystr, ActiveCell.row, (ActiveCell.Column + 1)
End Sub

Function LTRearrange(ByVal mystr As String, i As Integer, j As Integer)
    ' Function of Concen().
    ' Join words after string as long as len() <= 71
    Dim words() As String
    Dim NewStr As String
    NewStr = ""
    words() = Split(mystr, " ")
    For Each word In words()
        If Len(NewStr & " " & word) <= 71 Then
            NewStr = Trim(NewStr & " " & word)
        Else
            WriteIntoCell NewStr, i, (j + 1)
            i = i + 1
            NewStr = word
        End If
    Next
    WriteIntoCell NewStr, i, (j + 1)
End Function
Sub CompareRow_Callable()
    Dim sorted, exact As Boolean
    
    m = CInt(InputBox("First Row (Master)"))        ' The Row to be highlighted
    s = CInt(InputBox("Second Row (Slave)"))        ' The Row to be searched
    
    colour = InputBox("Color")
    If colour = "" Then
        colour = 5296274
    Else
        colour = CLng(colour)
    End If
        
    If MsgBox("Exact match?", vbYesNo) = vbYes Then
        exact = True
    Else
        exact = False
    End If
    
    If MsgBox("Sorted match?", vbYesNo) = vbYes Then
        sorted = True
    Else
        sorted = False
    End If
    
    CompareRow m, s, colour, exact, sorted
End Sub

Sub CompareRow(ByVal m As Integer, ByVal s As Integer, Optional ByVal colour As Long = 5296274, Optional exact, Optional sorted)
    Dim Datatange, area As Range

    With ActiveSheet
        Datarange = .UsedRange.Value
        MaxM = UBound(Datarange)
        
        Do While IsEmpty(Datarange(MaxM, m))
            MaxM = MaxM - 1
        Loop
        
        MaxS = UBound(Datarange)
        
        Do While IsEmpty(Datarange(MaxS, s))
            MaxS = MaxS - 1
        Loop
        
        LastFound = 1
        
        For i = 1 To MaxM
            j = LastFound
            If Datarange(i, m) <> "" Then
                Do While j <= MaxS
                    If (exact And Datarange(i, m) = Datarange(j, s)) Xor (Not exact And InStr(1, Datarange(i, m), Datarange(j, s)) <> 0) Then
                        Debug.Print "i = " & i & ", j = " & j & " --- " & Datarange(i, m) & " - " & Datarange(j, s) & " : OK!"
                        If area Is Nothing Then
                            Set area = .Cells(i, m)
                        Else
                            Set area = Union(area, .Cells(i, m))
                        End If
                        If sorted Then
                            LastFound = j
                        Else
                            LastFound = 1
                        End If
                        Exit Do
                    Else
                        Debug.Print "i = " & i & ", j = " & j & " --- " & Datarange(i, m) & " - " & Datarange(j, s)
                    End If
                    j = j + 1
                Loop
            End If
        Next
        
        If Not area Is Nothing Then
            area.Interior.color = colour
        End If
        
    End With
End Sub

Function WriteIntoCell(ByVal mystr As String, row As Integer, Column As Integer)
    ' Experimental
    Cells(row, Column).Value = mystr
End Function

Sub ExtractSUBTASK()
    Dim i, j As Integer
    Dim del_area As Range
    i = 1
    j = InputBox("The Column with ""SUBTASK"" inside.")
    Set del_area = Nothing

    For Each entry In ActiveSheet.UsedRange.Rows
        'If InStr(1, entry.Cells(1, j), "SUBTASK") <> 0 And (Left(entry.Cells(1, j), 1) = "F" Or Left(entry.Cells(1, j), 1) = "S" Or Left(entry.Cells(1, j), 1) = "<" Or Left(entry.Cells(1, j), 4) = "Ref." Or Left(entry.Cells(1, j), 3) = "REF") Then
        If InStr(1, entry.Cells(1, j), "SUBTASK") <> 0 And InStr(1, entry.Cells(1, j), " 70-") = 0 And InStr(1, entry.Cells(1, j), "-") <> 0 Then
        Else
            If del_area Is Nothing Then
                Set del_area = entry
            Else
                Set del_area = Union(del_area, entry)
            End If
        End If
    Next
    
    del_area.Delete
    
    i = 1
    If IsEmpty(Cells(i, j)) Then
        Exit Sub
    End If
    
    Do While Not IsEmpty(Cells(i, j))
        Cells(i, j) = Trim(Mid(Cells(i, j), InStr(1, Cells(i, j), "SUBTASK")))
        i = i + 1
    Loop
    
    'Sort Marco
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add key:=Cells(1, j), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveSheet.Sort
        .SetRange ActiveSheet.UsedRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort Marco
    
    i = 1
    Do While i <= ActiveSheet.UsedRange.Rows.Count
        If Mid(Cells(i, j), 18, 1) = "8" Or Mid(Cells(i, j), 9, 2) = "70" Then
            Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
    
    i = 2
    Do While i <= ActiveSheet.UsedRange.Rows.Count
        If Cells(i, j) = Cells(i - 1, j) Then
            Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
End Sub

Sub ExtractString(Optional ByVal j As Integer = 0, Optional ByVal mystr As String = "")
    If mystr = "" Then
        mystr = InputBox("String to extract.")
    End If
    If j = 0 Then
        j = InputBox("The Column you would like to serach.")
    End If

    Dim area As Range
    
    For Each entry In ActiveSheet.UsedRange.Rows
        If ActiveSheet.UsedRange.Rows.Count = 1 And Len(entry.Cells(1, j)) < 1 Then
            Exit For
        ElseIf InStr(1, entry.Cells(1, j), mystr) <> 0 Then
            If area Is Nothing Then
                Set area = entry
            Else
                Set area = Union(area, entry)
            End If
        End If
    Next
    
    If Not area Is Nothing Then
        area.Delete
    End If

End Sub

Sub Extract_highlighted(ByVal col As Integer, ByVal colour As Long, ByVal TrueKeepFalseDelete As Boolean, Optional ByVal StartRow As Integer = 2)
    Dim area As Range
    Dim cell As Range
    i = StartRow
    Do While Not IsEmpty(Cells(i, 2))
        Set cell = Cells(i, col)
        If (TrueKeepFalseDelete Xor (cell.Interior.color = colour)) Then
            If area Is Nothing Then
                Set area = cell
            Else
                Set area = Union(area, cell)
            End If
        End If
        i = i + 1
    Loop
    
    If Not area Is Nothing Then
        area.Delete xlUp
    End If
End Sub

Sub ExtractRMR()
    'For spreadsheet downloaded from SAP: ia17.
    'Extract long text lines with TV reference.
    
    Dim olds, news As Object
    Set olds = ActiveSheet
    Sheets.Add After:=Sheets(Sheets.Count)
    Set news = ActiveSheet
    news.Name = "Result"
    Dim i, j, k As Long
    Dim GrpCtr As String
    
    news.Cells(1, 1) = "Long Text"
    news.Cells(1, 2) = "Short Text"
    news.Cells(1, 3) = "GrpCtr"
    news.Cells(1, 4) = "Operation No."
    i = 1
    j = 2
    mystr = "TV"

    Do While i <= olds.UsedRange.Rows.Count
        If Not IsEmpty(olds.Cells(i, 5)) Then
            GrpCtr = olds.Cells(i, 5).text
        End If
        If InStr(1, olds.Cells(i, 3), mystr) <> 0 And Mid(olds.Cells(i, 3), InStr(1, olds.Cells(i, 3), mystr) + 2, 1) <> " " Then
            news.Cells(j, 1) = olds.Cells(i, 3)
            k = i - 1
            Do While Mid(olds.Cells(k, 12).text, 3, 1) <> "-"
                k = k - 1
            Loop
            news.Cells(j, 2) = olds.Cells(k, 12)
            news.Cells(j, 3) = GrpCtr
            news.Cells(j, 4) = olds.Cells(k, 3)
            j = j + 1
        End If
        i = i + 1
    Loop
    news.UsedRange.AutoFilter
    news.UsedRange.Columns.AutoFit
    With news.Range("A1:D1").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    news.Range("A1:D1").Font.Bold = True
    MsgBox ("Done.")
End Sub

Sub Extract_String_of_certain_Length()
    start = InputBox("Start of the String")
    Leng = InputBox("length of the string")
    If Leng = "" Or Not IsNumeric(Leng) Then
        For Each sentence In Selection
            sentence.Value = Mid(sentence, InStr(1, sentence, start))
        Next
    Else
        For Each sentence In Selection
            sentence.Value = Mid(sentence, InStr(1, sentence, start), CStr(Leng))
        Next
    End If
    
End Sub

Sub HighlightOPs_callable()
    Dim col As Integer
    Dim strS As String
    Dim colour As Long
    col = InputBox("col")
    strS = InputBox("String")
    colour = InputBox("Colour")
    If MsgBox("All sheets [Yes] / Local sheet [No] ?", vbYesNo) = vbNo Then
        localonly = True
    Else
        localonly = False
    End If
    For Each item In ActiveWorkbook.Sheets
        If localonly Then
            Logic = (item.Name = ActiveSheet.Name)
        Else
            Logic = True
        End If
        If Logic Then
            HighlightOPs item, col, strS, colour
        End If
    Next
End Sub

Sub HighlightOPs(ByRef shit As Variant, col As Integer, word As String, colour As Long, Optional highlight_numeric As Boolean = True, Optional ExactMatch As Boolean = False)
    word = UCase(word)
    Dim Logic As Boolean
    Dim area As Range
    Set area = Nothing
    
    If word = "INVALID" Then
        For Each cell In shit.UsedRange.Columns(col).Cells
            pos = InStr(1, UCase(cell), word)
            If pos <> 0 Then
                If InStr(pos, UCase(cell), "T") = 0 And InStr(1, UCase(cell), "(T)") = 0 Then
                    If area Is Nothing Then
                        Set area = cell
                    Else
                        Set area = Union(area, cell)
                    End If
                End If
            End If
        Next
    Else
        IsWordANumber = IsNumeric(word)
        For Each cell In shit.UsedRange.Columns(col).Cells
            If ExactMatch Then
                Logic = (UCase(cell) = word)
            Else
                Logic = InStr(1, UCase(cell), word) <> 0
            End If
            If Logic Then
                If IsWordANumber And IsNumeric(cell) And Not highlight_numeric Then
                Else
                    If area Is Nothing Then
                        Set area = cell
                    Else
                        Set area = Union(area, cell)
                    End If
                End If
            End If
        Next
    End If
    If Not (area Is Nothing) Then
        area.Interior.color = colour
    End If
End Sub

Sub M_HighlightSUBTASKs(ByRef shit As Variant)
    HighlightOPs shit, 1, "OMat", 65535, False
    HighlightOPs shit, 1, "tape", 65535, False
    HighlightOPs shit, 1, "-230", 255, True
    HighlightOPs shit, 1, "SUBTASK", 255, True
    HighlightOPs shit, 4, "INVALID", 255, True
    HighlightOPs shit, 4, "DO NOT USE", 255, True
    HighlightOPs shit, 4, "DELETED", 255, True
    HighlightOPs shit, 4, "VOID", 255, True
End Sub

Sub M_HighlightOPs(ByRef shit As Variant)
    HighlightOPs shit, 1, "OP", 65535, False
    HighlightOPs shit, 1, "70-", 65535, False
    HighlightOPs shit, 4, "INVALID", 255, True
    HighlightOPs shit, 4, "DO NOT USE", 255, True
    HighlightOPs shit, 4, "DELETED", 255, True
    HighlightOPs shit, 4, "VOID", 255, True
    HighlightOPs shit, 8, " ", 65535, True
End Sub

Sub M_HighlightTrent(ByRef shit As Variant)
    HighlightOPs shit, 4, "TRENT", 65535, True
    HighlightOPs shit, 4, "T8", 65535, True
    HighlightOPs shit, 4, "T7", 65535, True
    HighlightOPs shit, 4, "T5", 65535, True
End Sub

Sub M_HighlightCustomize(ByRef shit As Variant)
'    HighlightOPs shit, 4, "INVALID", 255, True
'    HighlightOPs shit, 4, "DO NOT USE", 255, True
'    HighlightOPs shit, 4, "DELETED", 255, True
'    HighlightOPs shit, 4, "VOID", 255, True
    HighlightOPs shit, 7, "CMM", 65535, True
End Sub

Sub mass_highlight()
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name <> "Result" Then
            'Application.Run "M_HighlightOPs", sheet
            'Application.Run "HighlightOPs", sheet, 7, "GRD", 65535, False
            'Application.Run "M_HighlightSUBTASKs", sheet
            'Application.Run "M_HighlightTrent", sheet
            Application.Run "M_HighlightCustomize", sheet
            'Application.Run "HighlightOPs", sheet, 1, "Touch", 65535, False
            'Application.Run "HighlightOPs", sheet, 6, "Touch", 65535, False
            
            Debug.Print sheet.Name & " OK"
            
        End If
    Next
End Sub

Sub mass_unhighlight()
    For Each sheet In ActiveWorkbook.Sheets
        With sheet.Cells.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Next
End Sub

Sub mass_extract()
    VBATurboMode True
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name <> "Result" Then
            sheet.Activate
            Application.Run "M_Extract_highlighted"
            Debug.Print sheet.Name & " OK"
        End If
    Next
    VBATurboMode False
End Sub

Sub M_Extract_highlighted()
    Extract_highlighted 1, 255, True
End Sub
Sub format_lul(Optional ByRef shit As Variant = Nothing)
    If shit Is Nothing Then
        Set shit = ActiveSheet
    End If
    
    shit.ListObjects.Add(xlSrcRange, shit.UsedRange, , xlYes).Name = "Table" & shit.Name
    shit.ListObjects("Table" & shit.Name).TableStyle = "TableStyleLight1"
End Sub

Sub mass_format()
    For Each shit In ActiveWorkbook.Sheets
        Application.Run "format_lul", shit
        shit.Columns.AutoFit
    Next
End Sub

Sub assigndate()
    Dim day As Date
    day = Date
    For Each entry In Selection
        If entry.row Mod 150 = 0 Then
            day = day + 1
            If Weekday(day) = vbSaturday Then
                day = day + 2
            End If
        End If
        If entry.row = 1 Then
        Else
            entry.Cells(1, 1) = day
        End If
    Next
End Sub

Sub assigntoday()
Attribute assigntoday.VB_ProcData.VB_Invoke_Func = "T\n14"
    For Each item In Selection
        item.FormulaR1C1 = Date
    Next
End Sub

Sub focus_on_abnormal_height()
Attribute focus_on_abnormal_height.VB_ProcData.VB_Invoke_Func = "h\n14"
    For i = Selection.row To ActiveSheet.UsedRange.Rows.Count
        If ActiveSheet.Rows(i).RowHeight > 20 Then
            ActiveSheet.Rows(i).Select
            Do While Range("A" & CStr(i)) <> "Task list"
                i = i - 1
            Loop
            j = 1
            Do While Left(Cells(i, j), 2) <> "H0"
                j = j + 1
            Loop
            MsgBox "Plan = " & Cells(i, j)
            Exit For
        End If
    Next
End Sub

Function RemoveDuplicates(ByVal sheet As Worksheet, ByVal col As Integer)
    sheet.UsedRange.RemoveDuplicates Columns:=col, Header:=xlYes
End Function

Sub mass_RevDup()
    For Each item In ActiveWorkbook.Sheets
        If item.Name <> "Result" Then
            RemoveDuplicates item, 3
        End If
    Next
End Sub

Sub GroupResults()

    Dim type_e As Collection
    Set type_e = New Collection
    With type_e
        .Add "H00", "T800"
        .Add "H01", "T700"
        .Add "H03", "524GHT"
        .Add "H04", "535E4"
        .Add "H07", "T900"
        .Add "H08", "T500"
    End With
    
    For Each shit In ActiveWorkbook.Sheets
        If Left(shit.Name, 1) = "H" And Right(shit.Name, 3) <> "099" Then
            shit.Activate
            shit.Range(Cells(2, 1), Cells(get_lastrow(2, shit, 1), 6)).copy
            Sheets(shit.index + 1).Activate
            Sheets(shit.index + 1).Cells(get_lastrow(2, Sheets(shit.index + 1), 1) + 1, 1).Select
            Sheets(shit.index + 1).Paste
            Application.DisplayAlerts = False
            shit.Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    For Each shit In ActiveWorkbook.Sheets
        shit.Name = Replace(shit.Name, "099", "999")
'        DeleteEmptyRows shit
'        For Each engine In type_e
'            If Left(shit.Name, 3) = engine Then
'                MsgBox type_e(engine.Index).Key     'broken as fuck
'                shit.Name = type_e(engine).Key      'broken as fuck
'            End If
'        Next
    Next
End Sub
Function get_lastrow(ByVal start As Integer, ByVal shit As Worksheet, ByVal col As Integer) As Integer
    i = start
    Do
        i = i + 1
    Loop Until IsEmpty(shit.Cells(i, col))
    get_lastrow = i - 1
End Function

Sub highlight_allcolumn()
    Dim Plan As String
    Plan = InputBox("Plan")
    For Each Column In ActiveSheet.UsedRange.Columns
        Application.Run "HighlightOps", Column.Column, Plan, 60000
    Next
End Sub

Sub ChangePrefixSign()                  'Useful shits are always simple and clear
    criteria = Split("REF FRS SB TV TASK SUBTASK")
    With ActiveSheet
        op = 0
        Do While Not IsEmpty(.Cells(1, 4 + (op * 5)))
            i = 2
            Do While Not (IsEmpty(.Cells(i, 3 + (op * 5))) And IsEmpty(.Cells(i, 4 + (op * 5))))
                For Each item In criteria
                    If .Cells(i, 3 + (op * 5)) = "/:" And Left(.Cells(i, 4 + (op * 5)), Len(item)) = item Then
                        .Cells(i, 3 + (op * 5)) = "*"
                        Exit For
                    End If
                Next
                i = i + 1
            Loop
            op = op + 1
        Loop
    End With
End Sub

Sub ConvertXLStoXLSM(Optional DPath As String = "NULL")
    If DPath = "NULL" Then
        DPath = get_folder()
    End If
    
    Dim list() As String
    ReDim list(0)
    writefiles DPath, list
    ReDim Preserve list(0 To (UBound(list) - 1))
        
    For iter = 0 To UBound(list)
        file = list(iter)
        If Right(file, 4) = ".xls" Then     'Skipping non .xls files
            SupressWarning False
            Workbooks.Open (file)
            ActiveWorkbook.SaveAs filename:=file & "m", FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
            ActiveWorkbook.Close
            SupressWarning True
            Kill file
        End If
    Next
End Sub

Sub SupressWarning(State As Boolean)
    Application.DisplayAlerts = State
    Application.AskToUpdateLinks = State
End Sub

Sub VBATurboMode(Enab As Boolean)
    If Enab Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.DisplayAlerts = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        Application.DisplayAlerts = True
    End If
End Sub

Function getsheet(sName As String, Optional CreateNonExistentSheet As Boolean = True) As Object
    On Error Resume Next
        exist = (Sheets(sName).Name <> "")
        If exist Then
            Set getsheet = ActiveWorkbook.Sheets(sName)
        ElseIf CreateNonExistentSheet Then
            Set getsheet = ActiveWorkbook.Sheets.Add(, Worksheets(Worksheets.Count))
            getsheet.Name = sName
        Else
            Set getsheet = Nothing
        End If
    On Error GoTo 0
End Function

Sub OCD()
    Set Oshet = ActiveSheet
    VBATurboMode True
    For Each sheet In ActiveWorkbook.Sheets
        sheet.Activate
        sheet.Cells(1, 1).Select
    Next
    Oshet.Activate
    Set Oshet = Nothing
    VBATurboMode False
    MsgBox "Fixed your OCD."
End Sub

Sub SortCells(SortRange As Range, SortColumn As Integer, HasHeader As Boolean)
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add _
        key:=SortRange.Columns(SortColumn), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
        .SetRange SortRange
        If HasHeader Then
            .Header = xlYes
        Else
            .Header = xlNo
        End If
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
