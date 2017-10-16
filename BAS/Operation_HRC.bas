Attribute VB_Name = "Operation_HRC"
' One off shits for searching HRC
' It's ok we now have snapshot searching argo. No more hassles

Sub Excel_HRC_Sorter1()
Attribute Excel_HRC_Sorter1.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim i, j As Integer
    i = 1
    j = 1
    Do While IsEmpty(Cells(i, 1)) = False
    
        If i > 1 Then
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                Cells(j, 2).Value = Cells(i, 1).Value
                j = j + 1
            End If
        ElseIf i = 1 Then
            Cells(j, 2).Value = Cells(i, 1).Value
            j = j + 1
        End If
        
        i = i + 1
    Loop
End Sub

Sub Excel_HRC_Sorter2()
    Dim MaxR, i, strStart, strEnd As Integer
    
    MaxR = ActiveSheet.UsedRange.Rows.Count
    i = 1
    
    Do While i <= MaxR
        strStart = InStr(1, Cells(i, 10), "HRC")
        If strStart <> 0 Then
            strEnd = strStart + 1
            Do While strEnd <= Len(Cells(i, 10))
                If IsAlphaNumeric(Mid(Cells(i, 10), strEnd, 1)) = False Then
                    Exit Do
                End If
                strEnd = strEnd + 1
            Loop
            Cells(i, 11) = Mid(Cells(i, 10), strStart, (strEnd - strStart))
        End If
        i = i + 1
    Loop
End Sub

Sub Excel_HRC_Master_Comparer()
    Dim i, j, MaxR As Integer
    MaxR = ActiveSheet.UsedRange.Rows.Count
    i = 1
    Do While i <= MaxR
        j = 1
        Do While IsEmpty(Worksheets("Master").Cells(j, 1)) = False
            If ActiveSheet.Cells(i, 1) = Worksheets("Master").Cells(j, 1) Then
                Cells(i, 2) = "HRC found in Master List"
                Exit Do
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop
    MsgBox ("Done!")
End Sub

Sub Excel_HRC_Extractor()
    Dim MaxR, i, strStart, j As Integer
    Dim behind, HRC As String

    MaxR = ActiveSheet.UsedRange.Rows.Count
    i = 1
    j = 1
    Do While i <= MaxR
        strStart = 1
        Do While InStr(strStart, Cells(i, 10), "HRC") <> 0
            strStart = InStr(strStart, Cells(i, 10), "HRC")
            Cells(j, 18) = Mid(Cells(i, 10), strStart, 9)
            If Len(Cells(j, 18)) = 9 And IsAlphaNumeric(Mid(Cells(j, 18), 9, 1)) = False Then
                Cells(j, 18) = Left(Cells(j, 18), 8)
            End If
            j = j + 1
            strStart = strStart + 8
        Loop
        i = i + 1
    Loop
    MsgBox "Done!"
End Sub

Sub apr26_1_extract_tooling()
    Dim r, j As Integer
    j = 1
    For r = 1 To ActiveSheet.UsedRange.Rows.Count
        If InStr(1, Cells(r, 1), "HRC") <> 0 Then
            Cells(j, 2) = Cells(r, 1)
            j = j + 1
        End If
    Next
End Sub

Sub apr26_2_ExtractHRC()
    'For spreadsheet downloaded from SAP: ia17.
    'Extract long text lines with TV reference.
    
    Dim olds, news As Object
    Set olds = ActiveSheet
    Sheets.Add After:=Sheets(Sheets.Count)
    news.Name = "Result"
    Set news = ActiveSheet
    Dim i, j, k As Long
    Dim GrpCtr As String
    
    news.Cells(1, 1) = "Long Text"
    news.Cells(1, 2) = "Short Text"
    news.Cells(1, 3) = "GrpCtr"
    news.Cells(1, 4) = "Operation No."
    i = 1
    j = 2
    mystr = "HRC"

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


Sub apr26_4_search()
    'Search strings in the ActiveSheet to all XLSX files in a folder.
    Dim a, master As Object
    Dim path, file As String
    'Set m = ActiveWorkbook.ActiveSheet
    path = get_folder() & "\"
    file = Dir(path & "*.xlsx")
    
    Set master = ActiveWorkbook
    
    Do While file <> ""
        Workbooks.Open (path & file)
        Set a = Workbooks(file)

        apr26_extract master, "HKE*CMM"
        apr26_Excel_HRC_Extractor
'        apr26_deletedups
'        apr26_deleteerrsheet
'        apr26_compare_master master
'        apr26_4_deletesheet
        
'        a.Save
        a.Close savechanges:=False
        file = Dir()
        Set a = Nothing
    Loop
End Sub


Sub apr26_extract(master As Object, find_str As String)
    Dim o, n, r, l, k As Object
    Set o = ActiveSheet
    master.Sheets.Add After:=Sheets(Sheets.Count)
    Set n = master.ActiveSheet
    n.Name = o.Name
        
    Dim r_max, c_max, i, j As Long
    Dim Plan, grp, op As String
    
'    r_max = o.UsedRange.Rows.count
'    c_max = o.UsedRange.Columns.count
    
    j = 1
    With o.UsedRange
        Set r = .Find(find_str, LookIn:=xlValues)
        If Not r Is Nothing Then
            firstaddress = r.Address
            Do
                n.Cells(j, 1) = r.text
                n.Cells(j, 2) = r.row
                
                Set r = .FindNext(r)
                j = j + 1
            Loop While Not r Is Nothing And r.Address <> firstaddress
        End If
    End With
    
    For x = 1 To n.UsedRange.Rows.Count Step 1
        If IsEmpty(n.Cells(x, 2)) = True Then
            Exit For
        End If
        y = n.Cells(x, 2).Value
        
        Do While o.Cells(y, 2) <> "Operation"
            y = y - 1
        Loop
        
        Z = n.Cells(x, 2).Value
        Do While o.Cells(Z, 1) <> "Task list"
            Z = Z - 1
        Loop
        
        y_n = 5
        z_n = 5
        Do While Len(Cells(Z, z_n + 1)) < 3
            z_n = z_n + 1
        Loop
        Do While Len(Cells(y, y_n + 2)) < 3
            y_n = y_n + 1
        Loop
        
        n.Cells(x, 3).Value = o.Range("A" & CStr(Z)).Offset(0, 3).Value
        n.Cells(x, 4).Value = o.Range("A" & CStr(Z)).Offset(0, z_n).Value
        n.Cells(x, 5).Value = o.Range("B" & CStr(y)).Offset(0, 2).Value
        n.Cells(x, 6).Value = o.Range("B" & CStr(y)).Offset(0, y_n).Value
    Next x
    
    Set o = Nothing
    Set n = Nothing
End Sub

Sub apr26_Excel_HRC_Extractor()
    Dim MaxR, i, strStart, j As Integer
    Dim behind, HRC As String
    Dim olds, news As Object
    
    Set olds = ActiveSheet
    MaxR = ActiveSheet.UsedRange.Rows.Count
    Sheets.Add After:=Sheets(Sheets.Count)
    Set news = ActiveSheet
    
    i = 1
    j = 1
    Do While i <= MaxR
        strStart = 1
        Do While InStr(strStart, olds.Cells(i, 1), "HRC") <> 0
            strStart = InStr(strStart, olds.Cells(i, 1), "HRC")
            news.Cells(j, 1) = Mid(olds.Cells(i, 1), strStart, 9)
            If Len(news.Cells(j, 1)) = 9 And IsAlphaNumeric(Mid(news.Cells(j, 1), 9, 1)) = False Then
                news.Cells(j, 1) = Left(news.Cells(j, 1), 8)
            End If
            j = j + 1
            strStart = strStart + 8
        Loop
        i = i + 1
    Loop
    
End Sub

Sub apr26_deleteerrsheet()
    ActiveSheet.Delete
End Sub

Sub apr26_deletedups()
    k = 2
    Dim news As Object
    Set news = ActiveSheet
    
    Range("A1").Select
    news.Sort.SortFields.Clear
    news.Sort.SortFields.Add key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With news.Sort
        .SetRange Range("A1:A" & news.UsedRange.Rows.Count)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Do While IsEmpty(news.Cells(k, 1)) = False
        If news.Cells(k, 1).text = news.Cells(k - 1, 1).text Then
            news.Rows(k).Delete
        Else
            k = k + 1
        End If
    Loop
End Sub

Sub apr26_4_deletesheet()
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Sheet1").Delete
    ActiveWorkbook.Sheets("Sheet2").Delete
    Application.DisplayAlerts = True
End Sub

Sub apr29_Excel_HRC_Extractor_samebook()
    Dim MaxR, i, strStart, j As Integer
    Dim behind, HRC As String
    
    For Each item In ActiveWorkbook.Sheets
        If item.Name <> "Master" And item.Name <> "Summary" Then
            i = 1
            j = 1
            MaxR = item.UsedRange.Rows.Count
            Do While i <= MaxR
                strStart = 1
                Do While InStr(strStart, item.Cells(i, 1), "HRC") <> 0
                    strStart = InStr(strStart, item.Cells(i, 1), "HRC")
                    item.Cells(j, 10) = Trim(Mid(item.Cells(i, 1), strStart, 9))
                    If Len(item.Cells(j, 10)) = 9 Then
                        If IsAlphaNumeric(Mid(item.Cells(j, 10).text, 9, 1)) = False Then
'                   If Len(item.Cells(j, 1)) = 9 And IsAlphaNumeric(Mid(item.Cells(j, 1), 9, 1)) = False Then
                            item.Cells(j, 10) = Trim(Left(item.Cells(j, 10), 8))
                        End If
                    End If
                    item.Cells(j, 11) = item.Cells(i, 3)
                    item.Cells(j, 12) = item.Cells(i, 4)
                    item.Cells(j, 13) = item.Cells(i, 5)
                    item.Cells(j, 14) = item.Cells(i, 6)
                    j = j + 1
                    strStart = strStart + 8
                Loop
                i = i + 1
            Loop
        End If
    Next item
End Sub

Sub apr29_console()
    For Each item In ActiveWorkbook.Sheets
        If item.Name <> "Master" Then
            For x = 1 To item.UsedRange.Rows.Count Step 1
                item.Cells(x, 10) = Trim(item.Cells(x, 10))
                item.Cells(x, 10) = Replace(item.Cells(x, 10), " ", "")
                item.Cells(x, 10) = Replace(item.Cells(x, 10), ",", "")
                item.Cells(x, 10) = Replace(item.Cells(x, 10), "-", "")
                item.Cells(x, 10) = Replace(item.Cells(x, 10), "(", "")
                item.Cells(x, 10) = Replace(item.Cells(x, 10), ")", "")
                item.Cells(x, 10) = Replace(item.Cells(x, 10), ".", "")
                item.Cells(x, 10) = Replace(item.Cells(x, 10), """", "")
            Next
        End If
    Next item
End Sub

Sub apr29_replace_1st()
    For Each item In ActiveWorkbook.Sheets
        If item.Name <> "Master" Then
            For x = 1 To item.UsedRange.Rows.Count Step 1
                item.Cells(x, 1) = Trim(item.Cells(x, 1))
                item.Cells(x, 1) = Replace(item.Cells(x, 1), " ", "")
                item.Cells(x, 1) = Replace(item.Cells(x, 1), ",", "")
                item.Cells(x, 1) = Replace(item.Cells(x, 1), "-", "")
                item.Cells(x, 1) = Replace(item.Cells(x, 1), "(", "")
                item.Cells(x, 1) = Replace(item.Cells(x, 1), ")", "")
                item.Cells(x, 1) = Replace(item.Cells(x, 1), ".", "")
                item.Cells(x, 1) = Replace(item.Cells(x, 1), """", "")
            Next
        End If
    Next item
End Sub

Sub apr29_bulk_compare()
    Dim master As Object
    Set master = ActiveWorkbook.Sheets("Master")
    Dim Summary As Object
    Set Summary = ActiveWorkbook.Sheets("Summary")
    Summary.Cells(1, 7).Value = 0
    Summary.Cells(2, 7).Value = 0
    Summary.Cells(3, 7).Value = 0
    Summary.Cells(4, 7).Value = 0
    Summary.Cells(1, 9).Interior.color = 255
    Summary.Cells(1, 11).Interior.color = 65535
    For Each item In ActiveWorkbook.Sheets
        If item.Name <> "Master" And item.Name <> "Summary" Then
            item.Activate
            
'            apr29_nofill
            apr29_compare_master master
            apr29_compare_master_addon
            apr29_count
            apr29_count2 Summary
            apr29_itemlist Summary
        End If
    Next
End Sub

Sub apr29_compare_master(master As Object)
    'modified version of CompareRow
    Dim list As Object
    Set list = ActiveSheet
    Dim i1, i2, ver1, ver2 As Integer
    
    list.Cells(1, 15).Interior.color = 5296274
    list.Cells(1, 16).Value = " = HRC found in master list"
    list.Cells(2, 15).Interior.color = 255
    list.Cells(2, 16).Value = " = ATAF updated, PMO not updated"
    list.Cells(3, 15).Interior.color = 65535
    list.Cells(3, 16).Value = " = PMO updated, ATAF not updated"
    list.Cells(4, 16).Value = " = HRC not found in master list"
    
    i1 = 1
    Do While IsEmpty(list.Cells(i1, 10)) = False
        ver1 = get_ver(list.Cells(i1, 10))
        i2 = 1
        Do While IsEmpty(master.Cells(i2, 1)) = False
            ver2 = get_ver(master.Cells(i2, 1))
            If list.Cells(i1, 10) = master.Cells(i2, 1) Then
                list.Cells(i1, 10).Interior.color = 5296274
                Exit Do
            ElseIf Mid(list.Cells(i1, 10), 1, 8) = Mid(master.Cells(i2, 1), 1, 8) And ver1 > ver2 Then
                list.Cells(i1, 10).Interior.color = 65535
            ElseIf Mid(list.Cells(i1, 10), 1, 8) = Mid(master.Cells(i2, 1), 1, 8) And ver1 < ver2 Then
                list.Cells(i1, 10).Interior.color = 255
            End If
            i2 = i2 + 1
        Loop
        i1 = i1 + 1
    Loop
End Sub

Function get_ver(ByVal mystr As String) As Integer
    If Len(mystr) = 9 Then
        get_ver = Asc(Right(mystr, 1))
        Exit Function
    End If
    get_ver = 0
End Function

Sub apr29_compare_master_addon()
    'modified version of CompareRow
    Dim list As Object
    Set list = ActiveSheet
    Dim i1, i2, m, s As Integer
    
    list.Cells(1, 15).Interior.color = 5296274
    list.Cells(1, 16).Value = " = HRC found in master list"
    list.Cells(2, 15).Interior.color = 255
    list.Cells(2, 16).Value = " = HRC suffix not updated"
    list.Cells(3, 16).Value = " = HRC not found in master list"
    list.Columns.AutoFit
    list.Columns("A:A").ColumnWidth = 20
    list.Columns("B:I").EntireColumn.Hidden = True
End Sub

Sub apr29_count()
    Dim list As Object
    Set list = ActiveSheet
    Dim i1, g, r, w, y  As Integer
    
    g = 0
    r = 0
    w = 0
    y = 0
    i1 = 1
    Do While IsEmpty(list.Cells(i1, 10)) = False
        If list.Cells(i1, 10).Interior.color = 5296274 Then
            g = g + 1
        ElseIf list.Cells(i1, 10).Interior.color = 255 Then
            r = r + 1
        ElseIf list.Cells(i1, 10).Interior.color = 65535 Then
            y = y + 1
        Else
            w = w + 1
        End If
        i1 = i1 + 1
    Loop
    list.Cells(1, 20).Value = "Count:"
    list.Cells(2, 20).Value = "Count:"
    list.Cells(3, 20).Value = "Count:"
    list.Cells(4, 20).Value = "Count:"
    list.Cells(1, 21).Value = g
    list.Cells(2, 21).Value = r
    list.Cells(3, 21).Value = y
    list.Cells(4, 21).Value = w
End Sub

Sub apr29_count2(Summary As Object)
    Dim list As Object
    Set list = ActiveSheet
    Dim i1, g, r, w As Integer
    Summary.Cells(1, 7).Value = Summary.Cells(1, 7).Value + list.Cells(1, 21).Value
    Summary.Cells(2, 7).Value = Summary.Cells(2, 7).Value + list.Cells(2, 21).Value
    Summary.Cells(3, 7).Value = Summary.Cells(3, 7).Value + list.Cells(3, 21).Value
    Summary.Cells(4, 7).Value = Summary.Cells(4, 7).Value + list.Cells(4, 21).Value
End Sub

Sub apr29_nofill()
    Dim list As Object
    Set list = ActiveSheet
    
    With list.Range("J:J").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub apr29_itemlist(Summary As Object)
    Dim r, y As Integer
    Dim list As Object
    Set list = ActiveSheet
    
    i = 1
    
    r = 1
    Do While IsEmpty(Summary.Cells(r, 10)) = False
        r = r + 1
    Loop
    
    y = 1
    Do While IsEmpty(Summary.Cells(y, 12)) = False
        y = y + 1
    Loop
            
    Do While IsEmpty(list.Cells(i, 10)) = False
        If list.Cells(i, 10).Interior.color = 255 Then
            Summary.Cells(r, 10) = list.Cells(i, 10)
            r = r + 1
        ElseIf list.Cells(i, 10).Interior.color = 65535 Then
            Summary.Cells(y, 12) = list.Cells(i, 10)
            y = y + 1
        End If
        i = i + 1
    Loop
End Sub

Sub no_dups()
    Dim i, j As Integer
    
    i = CInt(InputBox("Row"))
    j = 2
    Do While IsEmpty(Cells(j, i)) = False
        If Cells(j, i) = Cells(j - 1, i) Then
            Cells(j, i).Delete Shift:=xlUp
        Else
            j = j + 1
        End If
    Loop
    
End Sub

