Attribute VB_Name = "The_Grand_Search"
' Standalone version is more updated

'Land of Hope and Glory,
'Mother of the Free;
'How shall we extol thee
'Who are born of thee?
'Wider still and wider
'Shall thy bounds be set;
'God, who made thee mighty,
'Make thee mightier yet!
'God, who made thee mighty,
'Make thee mightier yet.

' Main sear algo for generating search result tables using snapshots dumped
' Spaghetti codes but hey it gets the job done

Sub ia17_xlsx_search_main()
    'Search strings in the ActiveSheet to all XLSX files in a folder.
    Dim a, result As Object
    Dim master As Workbook
    Dim parent, path, file, strFinal As String
    Dim find_str() As String
    Dim srh_ary() As Variant
    
    ReDim find_str(0)
    
    path = get_folder() & "\"
    parent = path & "..\"
    
    file = Dir(path & "*.xlsx")
    
    If file = "" Then
        MsgBox "No xlsx files present. Aborting", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Selection(Yes)/Manual Input(No)?", vbYesNo) = vbYes Then
        For Each num In Selection
            find_str(UBound(find_str)) = num.text
            ReDim Preserve find_str(UBound(find_str) + 1)
        Next
        ReDim Preserve find_str(UBound(find_str) - 1)
    Else
        find_str(0) = InputBox("Please enter the search string you want to find.")
        
addmore:
        If MsgBox("Do you want to add more?", vbYesNo) = vbYes Then
            cin = InputBox("Please enter the search string you want to find.")
            If cin <> "" Then
                ReDim Preserve find_str(UBound(find_str) + 1)
                find_str(UBound(find_str)) = cin
            End If
            GoTo addmore
        End If
    End If
    
    If UBound(find_str) = 0 And find_str(0) = "" Then
        MsgBox "No valid keyword entered.", vbExclamation, "Tosser!"
        Exit Sub
    End If
    
    If MsgBox("Putting search results together?", vbYesNo) = vbYes Then
        ReDim srh_ary(0 To 0)
        srh_ary(0) = find_str
    Else
        ReDim srh_ary(LBound(find_str) To UBound(find_str))
        For i = LBound(find_str) To UBound(find_str)
            Dim sub_ary(0) As String
            sub_ary(0) = find_str(i)
            srh_ary(i) = sub_ary
        Next
    End If
    
    For Each item In srh_ary
        file = Dir(path & "*.xlsx")
        
        Set master = Workbooks.Add
        Set result = master.Worksheets(1)
        result.Name = "Result"
        
        VBATurboMode True
        
        master.Worksheets(3).Delete
        master.Worksheets(2).Delete
        
        Do While file <> ""
            Set a = Workbooks.Open(path & file, , True)
            
            ia17_xlsx_extract_nondatarange master, a, item
            'ia17_xlsx_extract master, a, item               'rekt
    
            a.Close
            file = Dir()
            Set a = Nothing
        Loop
        
        result_generate master, result, item
        
        strFinal = ""
        For Each word In item
            strFinal = strFinal & word & ", "
        Next
        strFinal = Left(strFinal, Len(strFinal) - 2)
        
        If Len(strFinal) >= 50 Then strFinal = Left(strFinal, InStrRev(strFinal, ",")) & "& etc"
        
        master.Sheets("Result").Move Before:=Sheets(1)
        VBATurboMode False
        master.SaveAs (parent & "Search Result of " & filename_normalize(strFinal) & " (" & format(Date, "yyyy-mm-dd") & ").xlsx")
        master.Close
        Set result = Nothing
        Set master = Nothing
    Next
    
    Shell "C:\Windows\explorer.exe """ & parent & "", vbNormalFocus
End Sub

Function finalize(ByVal sheet As Worksheet)
    If IsEmptySheet(sheet) Then
        Application.DisplayAlerts = False
        sheet.Delete
        Application.DisplayAlerts = True
        Exit Function
    End If
    
    With sheet
        .Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("A1:H1") = Split("Search Hits,Line,Plan,Plan Name,Op,Op Short Text,Workctr.,Package Selected", ",")
        .Rows(1).Font.Bold = True
        .Rows(1).Font.Size = 14
        sheet.Activate
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        .Columns.AutoFit
    End With
End Function

Function result_generate(master As Workbook, result As Worksheet, ByRef item As Variant)
    With result
        .Cells(1, 1).Value = "SEARCH RESULT of"
        For Each word In item
            strF = strF & word & vbCrLf
        Next
        .Cells(1, 2).Value = Left(strF, Len(strF) - 1)
        
        .Range("A2:D2") = Split("Plans,Hit Counts,Op. Counts,Plan Counts", ",")
        i = 3
        For Each sheet In master.Sheets
            If sheet.Name <> "Result" Then
                .Cells(i, 1) = sheet.Name
                If IsEmpty(sheet.Cells(1, 2)) Then
                    .Cells(i, 2) = 0
                    .Cells(i, 3) = 0
                    .Cells(i, 4) = 0
                    With .Rows(i).Font
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.499984740745262
                    End With
                Else
                    .Cells(i, 2) = sheet.UsedRange.Rows.Count
                    plan_count = 0
                    op_count = 0
                    count_inv sheet, plan_count, op_count, 0, 1
                    .Cells(i, 3) = op_count
                    .Cells(i, 4) = plan_count
                    .Hyperlinks.Add .Cells(i, 1), "", "'" & sheet.Name & "'!A1", , sheet.Name
                End If
                i = i + 1
                finalize sheet
            End If
        Next
        .Cells(i, 1) = "Total"
        .Cells(i, 2).Formula = "=SUM(B3:B" & i - 1 & ")"
        .Cells(i, 3).Formula = "=SUM(C3:C" & i - 1 & ")"
        .Cells(i, 4).Formula = "=SUM(D3:D" & i - 1 & ")"
        .Rows(2).Font.Bold = True
        .Rows(i).Font.Bold = True
        .Rows(1).Font.Size = 20
        .Activate
        .Range(Cells(1, 3), Cells(1, 4)).Font.ThemeColor = xlThemeColorDark1
        For Each sheet In master.Sheets
            format_lul sheet
            If sheet.Name <> "Result" Then
                SortTable sheet
            End If
        Next
        .Columns.AutoFit
    End With
End Function
Sub SortTable(ByRef sheet As Variant)
    With sheet.ListObjects(1).Sort
        .SortFields.Clear
        .SortFields.Add key:=Range("Table" & Replace(sheet.Name, "-", "_") & "[[#Headers],[Line]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Function ia17_xlsx_extract(master As Workbook, ByVal data As Workbook, find_str As Variant)     'Rekt by memory overflow, outdated. DO NOT USE
    Dim j As Long
    Dim o, n, r As Object
    Set o = data.Sheets(1)
    Set n = master.Sheets.Add
    n.Name = o.Name
    j = 1
    Dim hitrange As Range
    For Each item In find_str
        With o.UsedRange
            Set r = .Find(item, LookIn:=xlValues, LookAt:=xlPart, searchorder:=xlByRows, MatchCase:=False)
            If Not r Is Nothing Then
                firstaddress = r.Address
                Do
'                    n.Cells(j, 1) = r.Text
'                    n.Cells(j, 2) = r.row
                    hitrange(j, 1) = r.text
                    hitrange(j, 2) = r.row
                    
                    Set r = .FindNext(r)
                    j = j + 1
                Loop While Not r Is Nothing And r.Address <> firstaddress
            End If
        End With
    Next
    n.Range("A1") = hitrange
    
    n.Columns(5).NumberFormat = "@"
    
    Datarange = o.UsedRange.Value           'By accessing the whole range only once,                        'REKT
    MaxR = UBound(Datarange)                'the time needed for searching should be vastly reduced.        'REKT
                                            
    w_n = 3
    x_n = 3
    y_v = 1
    y_n = 7
    z_n = 6
    v_v = 1     'v = Workcentre-related
    v_n = 6
    u_n = 5     'u = Maint-related
    u_v = 1
    
    Do While IsEmpty(Datarange(3, x_n))
        x_n = x_n + 1
    Loop
    Do While Len(Datarange(3, z_n)) < 3
        z_n = z_n + 1
    Loop
    Do While Datarange(y_v, 2) <> "Operation"
        y_v = y_v + 1
    Loop
    Do While Len(Datarange(y_v, y_n)) < 3
        y_n = y_n + 1
    Loop
    Do While IsEmpty(Datarange(y_v, w_n))
        w_n = w_n + 1
    Loop
    Do While Datarange(v_v, x_n) <> "Work center"
        v_v = v_v + 1
    Loop
    Do While IsEmpty(Datarange(v_v, v_n))
        v_n = v_n + 1
    Loop
    
    Do While u_v <= UBound(Datarange)
        If Datarange(u_v, 3) = "MntPack." Then
            Do While IsEmpty(Datarange(u_v, u_n))
                u_n = u_n + 1
            Loop
            Exit Do
        End If
        u_v = u_v + 1
    Loop
    
    MaxL = n.UsedRange.Rows.Count
    Editrange = Range(n.Cells(1, 1), n.Cells(MaxL, 8))
    
    x = 1
    Do While Not IsEmpty(Editrange(x, 2))
    
        Z = Editrange(x, 2)
        Do While IsEmpty(Datarange(Z, 1))
            Z = Z - 1
        Loop
        
        y = Editrange(x, 2)
        Do While Datarange(y, 2) <> "Operation" And y >= Z
            y = y - 1
        Loop
            
        V = y + 1
        Do While IsEmpty(Datarange(V, v_n)) Or Len(Datarange(V, v_n)) > 8
            V = V + 1
        Loop
        Editrange(x, 3) = Datarange(Z, x_n)
        Editrange(x, 4) = Datarange(Z, z_n)
        Editrange(x, 5) = format(Datarange(y, w_n), "0000")
        Editrange(x, 6) = Datarange(y, y_n)
        Editrange(x, 7) = Datarange(V, v_n)
        
        If x = 1 Then
            SameOp = False
        Else
            SameOp = (Editrange(x, 3) = Editrange(x - 1, 3)) And (Editrange(x, 4) = Editrange(x - 1, 4)) And (Editrange(x, 5) = Editrange(x - 1, 5))
        End If
        
        If SameOp Then
            Editrange(x, 8) = maint
        Else
            maint = ""
            u = 5
            Do While Datarange(V + u, 3) = "MntPack."
                maint = maint & Datarange(V + u, u_n) & " "
                maint = maint & Datarange(V + u, v_n) & vbCrLf
                u = u + 1
                If (V + u) > MaxR Then Exit Do
            Loop
            If maint <> "" Then maint = Left(maint, Len(maint) - 1)
            Editrange(x, 8) = maint
        End If
        x = x + 1
        If x > MaxL Then Exit Do
    Loop
    
    Range(n.Cells(1, 1), n.Cells(n.UsedRange.Rows.Count, 8)) = Editrange
    
    n.UsedRange.Columns.AutoFit
    n.UsedRange.Rows.AutoFit
    
    Set o = Nothing
    Set n = Nothing
End Function

Function ia17_xlsx_extract_nondatarange(master As Workbook, ByVal data As Workbook, find_str As Variant)
    Dim j As Long
    Dim o, n, r As Object
    Set o = data.Sheets(1)
    Set n = master.Sheets.Add
    n.Name = o.Name
    j = 1
    
    Dim textrow() As String
    Dim numrow() As Long
    ReDim textrow(1 To 1)
    ReDim numrow(1 To 1)
    
    For Each item In find_str
        With o.UsedRange
            Set r = .Find(item, LookIn:=xlValues, LookAt:=xlPart, searchorder:=xlByRows, MatchCase:=False)
            If Not r Is Nothing Then
                firstaddress = r.Address
                Do
                    textrow(j) = r.text
                    numrow(j) = r.row
                    Set r = .FindNext(r)
                    j = j + 1
                    ReDim Preserve textrow(1 To j)
                    ReDim Preserve numrow(1 To j)
                Loop While Not r Is Nothing And r.Address <> firstaddress
                n.Cells(1, 1).Resize(j - 1, 1) = Application.Transpose(textrow)
                n.Cells(1, 2).Resize(j - 1, 1) = Application.Transpose(numrow)
            End If
        End With
    Next
    
    Erase textrow, numrow
    
    n.Columns(5).NumberFormat = "@"
    
    MaxR = o.UsedRange.Rows.Count
    MaxL = j - 1
                                            
    w_n = 3
    x_n = 3
    y_v = 1
    y_n = 7
    z_n = 6
    v_v = 1     'v = Workcentre-related
    v_n = 6
    u_n = 5     'u = Maint-related
    u_v = 1
    
    Do While IsEmpty(o.Cells(3, x_n))
        x_n = x_n + 1
    Loop
    Do While Len(o.Cells(3, z_n)) < 3
        z_n = z_n + 1
    Loop
    Do While o.Cells(y_v, 2) <> "Operation"
        y_v = y_v + 1
    Loop
    Do While Len(o.Cells(y_v, y_n)) < 3
        y_n = y_n + 1
    Loop
    Do While IsEmpty(o.Cells(y_v, w_n))
        w_n = w_n + 1
    Loop
    Do While o.Cells(v_v, x_n) <> "Work center"
        v_v = v_v + 1
    Loop
    Do While IsEmpty(o.Cells(v_v, v_n))
        v_n = v_n + 1
    Loop
    
    Do While u_v <= MaxR
        If o.Cells(u_v, 3) = "MntPack." Then
            HavePack = True
            Exit Do
        End If
        u_v = u_v + 1
    Loop
    
    If HavePack Then
        Do While IsEmpty(o.Cells(u_v, u_n))
            u_n = u_n + 1
        Loop
    End If
    
    
    If MaxL <> 0 Then
        Editrange = Range(n.Cells(1, 1), n.Cells(MaxL, 8))
        For x = 1 To MaxL
            
            Z = Editrange(x, 2)
            Do While IsEmpty(o.Cells(Z, 1))
                Z = Z - 1
            Loop
            
            y = Editrange(x, 2)
            Do While o.Cells(y, 2) <> "Operation" And y >= Z
                y = y - 1
            Loop
            
            Editrange(x, 3) = o.Cells(Z, x_n)
            Editrange(x, 4) = o.Cells(Z, z_n)
            Editrange(x, 5) = format(o.Cells(y, w_n), "0000")
            
            If x = 1 Then
                SameOp = False
            Else
                SameOp = (Editrange(x, 3) = Editrange(x - 1, 3)) And (Editrange(x, 4) = Editrange(x - 1, 4)) And (Editrange(x, 5) = Editrange(x - 1, 5))
            End If
            
            If SameOp Then
                Editrange(x, 6) = Editrange(x - 1, 6)
                Editrange(x, 7) = Editrange(x - 1, 7)
                Editrange(x, 8) = maint
            Else
            
                Editrange(x, 6) = o.Cells(y, y_n)
                
                V = y + 1
                Do While IsEmpty(o.Cells(V, v_n))
                    V = V + 1
                Loop
                Editrange(x, 7) = o.Cells(V, v_n)
                
                maint = ""
                u = 5
                Do While o.Cells(V + u, 3) = "MntPack."
                    If maint <> "" Then maint = maint & vbCrLf
                    maint = maint & o.Cells(V + u, u_n) & " " & o.Cells(V + u, v_n)
                    u = u + 1
                Loop
                If maint <> "" Then Editrange(x, 8) = maint
            End If
        Next
        Range(n.Cells(1, 1), n.Cells(MaxL, 8)) = Editrange
    End If
    
    n.UsedRange.Columns.AutoFit
    n.UsedRange.Rows.AutoFit
    
    Set o = Nothing
    Set n = Nothing
End Function

Sub dumbfire()
    ia17_xlsx_search_main
    Shutdown
End Sub
