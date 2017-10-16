Attribute VB_Name = "Snapshot_Dump"
' Dumping snapshots since 2016. Satisfcation guaranteed. (tm)

Function dump_H0plans(ByRef mysap As Variant, ByVal path As String)

    On Error Resume Next
    
    Application.DisplayAlerts = False

'    Dim type_e As Collection
'    Set type_e = New Collection
'    With type_e
'        .Add "H00", "T800"
'        .Add "H01", "T700"
'        .Add "H03", "524GHT"
'        .Add "H04", "535E4"
'        .Add "H07", "T900"
'        .Add "H08", "T500"
'    End With
    
    Dim plan1, plan2 As String
    
    For Each engine In type_e
        For i = 0 To 900 Step 100
            j = i + 99
            plan1 = engine & format(CStr(i), "000")
            plan2 = engine & format(CStr(j), "000")
            With mysap
                '.FindById("wnd[0]").resizeWorkingPane 133, 46, False
                .SendCommand ("/nia17")
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = plan1
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-HIGH").text = plan2
                .FindById("wnd[0]/usr/ctxtPN_WERKS-LOW").text = "HK01"
                .FindById("wnd[0]").SendVKey 8
                
                .FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
                If Err.Number = 0 Then
                    .FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                    .FindById("wnd[1]/usr/ctxtDY_PATH").text = path
                    .FindById("wnd[1]/usr/ctxtDY_FILENAME").text = plan1 & "-" & plan2 & ".XLS"
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                Else
                    Err.Number = 0
                End If
            End With
        Next i
    Next
    
    Application.DisplayAlerts = True
End Function

Function dump_HQplans(ByRef mysap As Variant, ByVal path As String)

    On Error Resume Next
    
    Application.DisplayAlerts = False
    
'    Dim type_e As Collection
'    Set type_e = New Collection
'    With type_e
'        .Add "HQ0", "EOH1"
'        .Add "HQ1", "EOH2"
'        '.Add "HQ00", "EOH3"
'    End With
    
    Dim plan1, plan2 As String
    
    For Each engine In type_e
        For i = 0 To 900 Step 100
            j = i + 99
            plan1 = engine & format(CStr(i), "000")
            plan2 = engine & format(CStr(j), "000")
            With mysap
                '.FindById("wnd[0]").resizeWorkingPane 133, 46, False
                .SendCommand ("/nia17")
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = plan1
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-HIGH").text = plan2
                .FindById("wnd[0]/usr/ctxtPN_WERKS-LOW").text = "HK01"
                .FindById("wnd[0]").SendVKey 8
                
                .FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
                If Err.Number = 0 Then
                    .FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                    .FindById("wnd[1]/usr/ctxtDY_PATH").text = path
                    .FindById("wnd[1]/usr/ctxtDY_FILENAME").text = plan1 & "-" & plan2 & ".XLS"
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                Else
                    Err.Number = 0
                End If
            End With
        Next i
    Next
    
    Application.DisplayAlerts = True
End Function

Function dump_HIplans(ByRef mysap As Variant, ByVal path As String)

    On Error Resume Next
    
    Application.DisplayAlerts = False

    With mysap
        '.FindById("wnd[0]").resizeWorkingPane 133, 46, False
        .SendCommand ("/nia17")
        .FindById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = "HI*"
        .FindById("wnd[0]/usr/ctxtPN_PLNNR-HIGH").text = ""
        .FindById("wnd[0]/usr/ctxtPN_WERKS-LOW").text = "HK01"
        .FindById("wnd[0]").SendVKey 8
        
        .FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
        If Err.Number = 0 Then
            .FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
            .FindById("wnd[1]/tbar[0]/btn[0]").press
            .FindById("wnd[1]/usr/ctxtDY_PATH").text = path
            .FindById("wnd[1]/usr/ctxtDY_FILENAME").text = "HI_.XLS"
            .FindById("wnd[1]/tbar[0]/btn[0]").press
        Else
            Err.Number = 0
        End If
    End With

    Application.DisplayAlerts = True
End Function

Function dump_plans(mysap, ItemList, path)

    On Error Resume Next
    Application.DisplayAlerts = False
    
    For Each item In ItemList
        For i = 0 To 900 Step 100
            j = i + 99
            If item = "HI" Then
                plan1 = "HI*"
                plan2 = ""
                filename = "HI_.XLS"
                i = 900
            Else
                plan1 = item & format(CStr(i), "000")
                plan2 = item & format(CStr(j), "000")
                filename = plan1 & "-" & plan2 & ".XLS"
            End If
            
            With mysap
                .SendCommand ("/nia17")
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = plan1
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-HIGH").text = plan2
                .FindById("wnd[0]/usr/ctxtPN_WERKS-LOW").text = "HK01"
                .FindById("wnd[0]").SendVKey 8
                .FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
                If Err.Number = 0 Then
                    .FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                    .FindById("wnd[1]/usr/ctxtDY_PATH").text = path
                    .FindById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                Else
                    Err.Number = 0
                End If
            End With
        Next i
    Next
    
    Application.DisplayAlerts = True
    On Error GoTo 0
    
End Function

Sub main()
    Dim path, file As String
    
    path = get_folder()
    xls_path = path & "\xls"
    xlsx_path = path & "\xlsx"
    
    If Len(Dir(xls_path, vbDirectory)) = 0 Then
        MkDir xls_path
    End If
    
    file = Dir(xls_path & "\*.XLS?")

    If file <> "" Then
        MsgBox ".xls or .xlsx files exist at the directory: " & xls_path & "." & vbCrLf & "Aborting.", vbExclamation
        Exit Sub
    End If
    
    Do While file <> ""
        file = Dir()
    Loop
    
    If MsgBox("This action will take extremely long time (~90mins),  are you sure to proceed?", vbOKCancel) = vbOK Then
        Set mysap = getSession()
        'dump_HQplans mysap, xls_path
'        dump_H0plans mysap, xls_path
'        dump_HIplans mysap, xls_path
        dump_plans mysap, Split("H00 H01 H03 H04 H07 H08"), xls_path
        dump_plans mysap, Split("HI"), xls_path
        xls_to_xlsx path, xls_path, xlsx_path
        
        Kill xls_path & "\*.*"
        RmDir xls_path
        'Combine_xlsx xlsx_path         'Combined file too large. Excel failed to save.
        Shell "C:\Windows\explorer.exe """ & path & "", vbNormalFocus
    Else
        MsgBox "Action aborted.", vbExclamation
        Exit Sub
    End If
End Sub

Function xls_to_xlsx(ByVal path As String, ByVal xls_path As String, ByVal xlsx_path As String)
    Dim file As String
    Dim a As Object
    
    If Len(Dir(path & "\xlsx", vbDirectory)) = 0 Then
        MkDir path & "\xlsx"
    End If
    
    file = Dir(xls_path & "\*.XLS")
    
    Do While file <> ""
        Workbooks.Open (xls_path & "\" & file)
        Set a = ActiveWorkbook
        NewFile = Left(file, InStr(file, ".")) & "xlsx"
        
        a.SaveAs filename:=(xlsx_path & "\" & NewFile), FileFormat:=xlOpenXMLWorkbook
        a.Close
        file = Dir()
        Set a = Nothing
    Loop
    
End Function

Sub Combine_xlsx(ByVal xlsx_path As String)
    VBATurboMode True
    file = Dir(xlsx_path & "\*.XLSX")
    Dim db As Workbook
    Dim ds As Workbook
    Dim sheet As Worksheet
    Set db = Workbooks.Add
    Do While file <> ""
        Set ds = Workbooks.Open(xlsx_path & "\" & file, , True)
        ds.Sheets(1).copy After:=db.Sheets(db.Sheets.Count)
        ds.Close
        Set ds = Nothing
        Set sheet = Nothing
        file = Dir()
    Loop
    
    db.Sheets("Sheet1").Delete
    db.Sheets("Sheet2").Delete
    db.Sheets("Sheet3").Delete
    db.SaveAs xlsx_path & "\db.xlsx"
    db.Close
    Set db = Nothing
    VBATurboMode False
End Sub
Sub combine_dumb()
    Combine_xlsx "C:\Users\orix.auyeung\Desktop\SNAPSHOTS\H0\20161011(test)\xlsx"
End Sub

Sub shutdown_after_dump()   'Does what the title say
    main
    Shutdown
End Sub
