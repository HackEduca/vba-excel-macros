Attribute VB_Name = "HellOnEarth"
'Collections of one-off codes written for updating the GHT outdated repair schemes.
'No longer in use, just for references.

Sub HellOnEarth_SAP_cv04n_Search()
Attribute HellOnEarth_SAP_cv04n_Search.VB_ProcData.VB_Invoke_Func = " \n14"

    On Error Resume Next

    Set SAPsession = getSession()
    
    Dim i, j As Integer
    i = 1
    j = 1
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Result_Neu"
    
    SAPsession.FindById("wnd[0]").resizeWorkingPane 117, 31, False

    
    Do While IsEmpty(Worksheets("Sort").Cells(i, 1).Value) = False
        SAPsession.SendCommand ("/ncv04n")
        Worksheets("Result_Neu").Cells(j, 1).Value = Worksheets("Sort").Cells(i, 1).Value
        j = j + 1
        ActiveCell.Offset(1, 0).Range("A1").Select
'        Worksheets("Result").Cells(j, 1).Value = "20SEP14 Hits"
'        Worksheets("Result").Cells(j, 2).Value = "Total Hits"
'        j = j + 1
        SAPsession.FindById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/ssubSCR_MAIN:SAPLCV100:0402/txtSTDKTXT-LOW").text = "*G*" & Worksheets("Sort").Cells(i, 1).Value & "*"
        SAPsession.FindById("wnd[0]").SendVKey 8
        
        If SAPsession.FindById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/ssubSCR_MAIN:SAPLCV100:0402/txtSTDKTXT-LOW").text = "" Then
            SAPsession.FindById("wnd[0]").SendVKey 9
    '        SAPsession.findById("wnd[1]/tbar[0]/btn[9]").press
            SAPsession.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    '        SAPsession.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
            SAPsession.FindById("wnd[1]/tbar[0]/btn[0]").press
        
            ActiveSheet.Paste
            
            
            Do While 1 = 1
                j = j + 1
                ActiveCell.Offset(1, 0).Range("A1").Select
                If Left(Worksheets("Result_Neu").Cells(j, 1).text, 4) = "|  0" And Left(Worksheets("Result_Neu").Cells(j + 1, 1).text, 3) = "---" Then
                    Exit Do
                End If
            Loop
            j = j + 1
            ActiveCell.Offset(1, 0).Range("A1").Select
    '        SAPsession.findById("wnd[0]").sendVKey 71
    '        SAPsession.findById("wnd[1]/usr/chkGS_SEARCH-SHOW_HITS").Selected = True
    '        SAPsession.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "20SEP14"
    '        SAPsession.findById("wnd[1]").sendVKey 0
    '        Worksheets("Result").Cells(j, 1).Value = Right(SAPsession.findById("wnd[1]/usr/txtGS_SEARCH-SEARCH_INFO").Text, 1)
    '        SAPsession.findById("wnd[1]/usr/chkGS_SEARCH-SHOW_HITS").Selected = True
    '        SAPsession.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").Text = Worksheets("Sort").Cells(i, 1).Value
    '        SAPsession.findById("wnd[1]").sendVKey 0
    '        Worksheets("Result").Cells(j, 2).Value = Right(SAPsession.findById("wnd[1]/usr/txtGS_SEARCH-SEARCH_INFO").Text, 1)
    '        SAPsession.findById("wnd[1]").Close
            SAPsession.FindById("wnd[0]/tbar[0]/btn[3]").press
            
        End If
        
        j = j + 2
        ActiveCell.Offset(2, 0).Range("A1").Select
        i = i + 1
    Loop
    
    Cells.Columns.AutoFit
    MsgBox "Done"
End Sub

Sub HellOnEarth_SortFields()
    Dim i As Integer
    i = 1
    Do While ((((IsEmpty(Cells(i, 1)) = False) Or (IsEmpty(Cells(i + 1, 1)) = False)) Or (IsEmpty(Cells(i + 2, 1)) = False)) Or (IsEmpty(Cells(i + 3, 1)) = False))
        If Left(Cells(i, 1).text, 4) = "|  0" Then
            Cells(i, 2).Value = Mid(Cells(i, 1).text, InStr(1, Cells(i, 1).text, "G"), (InStr(InStr(1, Cells(i, 1).text, "G"), Cells(i, 1).text, "|") - InStr(1, Cells(i, 1).text, "G")))
            Cells(i, 2).Value = Trim(Cells(i, 2).Value)
        End If
        i = i + 1
    Loop
End Sub

Sub HellOnEarth_CheckAvailability()
    Dim fso As Object
    Dim path As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i, j As Integer
    i = 1
    Do While ((((IsEmpty(Cells(i, 1)) = False) Or (IsEmpty(Cells(i + 1, 1)) = False)) Or (IsEmpty(Cells(i + 2, 1)) = False)) Or (IsEmpty(Cells(i + 3, 1)) = False))
        If Left(Cells(i, 1).text, 4) = "|  0" Then
            path = "Z:\TechData\G\PCdrawing\RB211_524GHT\" & Cells(i, 2).text & ".pdf"
            If fso.FileExists(path) Or InStr(1, Cells(i, 2).Value, "DELETED") <> 0 Then
                Cells(i, 3).Value = ""
            Else
                Cells(i, 3).Value = "NO ACCESS"
            End If
            If InStr(1, Cells(i, 1).Value, "20SEP14") = 0 And (InStr(1, Cells(i, 2).Value, "DELETED") = 0 Or InStr(1, Cells(i, 2).Value, "INVALID") = 0) Then
                Cells(i, 4).Value = "DATE MISMATCH"
            Else
                Cells(i, 4).Value = ""
            End If
            Cells(i, 5).Value = Right(Cells(i, 2).Value, 7)
            If (VarType(Cells(i, 5)) = 7 And Cells(i, 5).Value < #9/20/2014#) Then
                Cells(i, 6).Formula = "=""" & Mid(Cells(i, 1).Value, 4, 2) & """"
                Cells(i, 7).Formula = "=""" & Mid(Cells(i, 1).Value, 7, 11) & """"
                Cells(i, 8).Value = Left(Cells(i, 2).Value, Len(Cells(i, 2).Value) - 7) & "20SEP14"
            ElseIf (VarType(Cells(i, 5)) = 7 And Cells(i, 5).Value > #9/20/2014#) Then
                Cells(i, 4).Value = ""
            End If
        End If
        If InStr(1, Cells(i, 1).Value, "FRS") <> 0 And Left(Cells(i, 1).Value, 1) <> "F" Then
        j = i - 1
            Do While Left(Cells(j, 1).Value, 1) <> "F"
                j = j - 1
            Loop
            If Cells(j, 3).Value = "TRANSFORMED!" And IsNumeric(Left(Right(Cells(i, 2).Value, 11), 3)) Then
                Cells(i, 3).Value = "TRANSFORMED!"
            End If
        End If
        i = i + 1
    Loop
End Sub

Sub HellOnEarth_SAP_cv02n_UpdateDate()
    
    On Error Resume Next

    Set mysap = getSession()
    
    Dim i As Integer
    
    Do While ((((IsEmpty(Cells(i, 1)) = False) Or (IsEmpty(Cells(i + 1, 1)) = False)) Or (IsEmpty(Cells(i + 2, 1)) = False)) Or (IsEmpty(Cells(i + 3, 1)) = False))
        If IsEmpty(Cells(i, 8)) = False Then
            With mysap
              .FindById("wnd[0]").resizeWorkingPane 117, 31, False
              .SendCommand ("/nCV02n")
              .FindById("wnd[0]/usr/ctxtDRAW-DOKNR").text = Cells(i, 7).Value
              .FindById("wnd[0]/usr/ctxtDRAW-DOKAR").text = "DAT"
              .FindById("wnd[0]/usr/ctxtDRAW-DOKTL").text = "000"
              .FindById("wnd[0]/usr/ctxtDRAW-DOKVR").text = Cells(i, 6).Value
              .FindById("wnd[0]/usr/ctxtDRAW-DOKVR").SetFocus
              .FindById("wnd[0]/usr/ctxtDRAW-DOKVR").CaretPosition = 2
              .FindById("wnd[0]").SendVKey 0
              .FindById("wnd[0]/tbar[1]/btn[20]").press
              .FindById("wnd[0]/tbar[1]/btn[20]").press
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/ctxtTDWST-STABK").text = "ie"
              .FindById("wnd[0]").SendVKey 0
              .FindById("wnd[1]/usr/txtDRAP-PROTF").text = "1"
              .FindById("wnd[1]/usr/txtDRAP-PROTF").CaretPosition = 1
              .FindById("wnd[1]").SendVKey 0
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").selectNode "          3"
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/txtDRAT-DKTXT").text = Cells(i, 8).Value
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_DELETE").press
              .FindById("wnd[1]/usr/btnSPOP-OPTION1").press
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_CREATE").press
              .FindById("wnd[1]/usr/ctxtDRAW-DAPPL").text = "PDF"
              .FindById("wnd[1]/usr/ctxtDRAW-DTTRG").text = "H-GENERIC"
              .FindById("wnd[1]/usr/ctxtDRAW-FILEP").text = "PCdrawing\RB211_524GHT\" & Cells(i, 8).Value & ".pdf"
              .FindById("wnd[1]/usr/ctxtDRAW-FILEP").SetFocus
              .FindById("wnd[1]/usr/ctxtDRAW-FILEP").CaretPosition = 86
              .FindById("wnd[1]/tbar[0]/btn[0]").press
              .FindById("wnd[0]/tbar[1]/btn[20]").press
              .FindById("wnd[1]/usr/btnSPOP-OPTION1").press
              .FindById("wnd[0]/tbar[1]/btn[20]").press
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/ctxtTDWST-STABK").text = "FR"
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/ctxtTDWST-STABK").SetFocus
              .FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/ctxtTDWST-STABK").CaretPosition = 2
              .FindById("wnd[0]").SendVKey 0
              .FindById("wnd[1]/usr/txtDRAP-PROTF").text = "1"
              .FindById("wnd[1]/usr/txtDRAP-PROTF").CaretPosition = 1
              .FindById("wnd[1]").SendVKey 0
              .FindById("wnd[0]/tbar[0]/btn[11]").press
            
            End With
        End If
        i = i + 1
    Loop
    
End Sub

Sub Runtime()
    HellOnEarth_SAP_cv04n_Search
    ActiveWorkbook.Save
    Shutdown
End Sub

Sub HellOnEarth_SAP_sq01_DB_Search()

    On Error Resume Next

    Set SAPsession = getSession()
    
    Dim i, j, k As Integer
    i = 1
    j = 1
    
    SAPsession.FindById("wnd[0]").resizeWorkingPane 117, 31, False
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Plan_Neu"
    
    Do While IsEmpty(Worksheets("Sort").Cells(i, 1).Value) = False
    
        SAPsession.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "HK01"
        SAPsession.FindById("wnd[0]/usr/txtSP$00001-LOW").text = "*" & Worksheets("Sort").Cells(i, 1).Value & "*"
        Worksheets("Plan_Neu").Cells(j, 1).Value = Worksheets("Sort").Cells(i, 1).Value
        
        SAPsession.FindById("wnd[0]").SendVKey 8
        j = j + 1
        
        If SAPsession.FindById("wnd[0]/usr/ctxtPLANT-LOW").text <> "HK01" Then
            
            SAPsession.FindById("wnd[0]/tbar[1]/btn[6]").press
    
            k = 0
            Do While SAPsession.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0," & CStr(k) & "]").text <> "____________"
                
                If k = 27 Then
                    SAPsession.FindById("wnd[0]/tbar[0]/btn[82]").press
                    k = 0
                End If
                
                If Left(SAPsession.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0," & CStr(k) & "]").text, 3) = "H03" Then
                    Worksheets("Plan_Neu").Cells(j, 1).Value = SAPsession.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0," & CStr(k) & "]").text
                    Worksheets("Plan_Neu").Cells(j, 2).Value = SAPsession.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-KZYK1[1," & CStr(k) & "]").text
                    Worksheets("Plan_Neu").Cells(j, 3).Value = SAPsession.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-KTEX1[2," & CStr(k) & "]").text
                    j = j + 1
                End If
                k = k + 1
            Loop
            SAPsession.FindById("wnd[0]/tbar[0]/btn[3]").press
            SAPsession.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
        
        j = j + 1
        i = i + 1
    Loop
    
    Cells.Columns.AutoFit
    MsgBox "Done"
End Sub

Sub HellOnEarth_SAP_ip03_Search()
    On Error Resume Next

    Set SAPsession = getSession()
    
    Dim i, j, k As Integer
    i = 1
    j = 1
    
    SAPsession.FindById("wnd[0]").resizeWorkingPane 117, 31, False
    
    Do While (((IsEmpty(Cells(i, 1)) = False) Or (IsEmpty(Cells(i + 1, 1)) = False)) Or (IsEmpty(Cells(i + 2, 1)) = False))
    
        If Left(Cells(i, 1).Value, 3) = "H03" And (Cells(i, 1).Value <> Cells(i - 1, 1).Value) Then
            SAPsession.SendCommand ("/nip03")
            SAPsession.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Cells(i, 1).Value & "/1"
            SAPsession.FindById("wnd[0]").SendVKey 0
            Cells(i, 4).Value = "Active Group Counter:"
            Cells(i, 5).Value = SAPsession.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text
        End If
        i = i + 1
    Loop
    Cells.Columns.AutoFit
    MsgBox "Done"
End Sub

Sub HellOnEarth_FetchFileNames()
    Dim fso, folder, item As Object
    Dim path As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = "Z:\TechData\G\PCdrawing\RB211_524GHT\"
    Set folder = fso.GetFolder(path)
    Set item = folder.Files

    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "FileNames_Neu"
    
    Dim j As Integer
    j = 1
    
    For Each i In item
        If InStr(1, i.Name, "20SEP14") <> 0 Then
            Cells(j, 1).Value = i.Name
            j = j + 1
        End If
    Next
    
End Sub

Sub HellOnEarth_FindValidFiles()

    Dim i, j As Integer
    Dim FRS As String
    j = 1

    Do While (((IsEmpty(Cells(j, 1)) = False) Or (IsEmpty(Cells(j + 1, 1)) = False)) Or (IsEmpty(Cells(j + 2, 1)) = False))
        If (Cells(j, 3).Value = "NO ACCESS" And Right(Cells(j, 2).Value, 7) = "20SEP14") And (IsNumeric(Left(Right(Cells(j, 2).Value, 11), 3)) Or Left(Right(Cells(j, 2).Value, 12), 4) = "LIST" Or Left(Right(Cells(j, 2).Value, 12), 4) = "List" Or Left(Right(Cells(j, 2).Value, 12), 4) = "list") Then
            FRS = Mid(Cells(j, 2).Value, InStr(1, Cells(j, 2).Value, "FRS"), 7)
            
            i = 1
            Do While IsEmpty(Worksheets("FileNames_Neu").Cells(i, 1)) = False
                If InStr(1, Worksheets("FileNames_Neu").Cells(i, 1).Value, FRS) <> 0 And Left(Right(Worksheets("FileNames_Neu").Cells(i, 1).Value, 11), 3) = Left(Right(Cells(j, 2).Value, 11), 3) Then
                    Cells(j, 6).Formula = "=""" & Mid(Cells(j, 1).Value, 4, 2) & """"
                    Cells(j, 7).Formula = "=""" & Mid(Cells(j, 1).Value, 7, 11) & """"
                    Cells(j, 8).Value = Worksheets("FileNames_Neu").Cells(i, 1).Value
                    Exit Do
                End If
                i = i + 1
            Loop
            
        End If
        j = j + 1
    Loop
End Sub

Sub HellOnEarth_CheckIfTransformed()
    Dim i, j As Integer
    Dim FRS As String
    j = 1
    i = 1
    Do While (((IsEmpty(Cells(j, 1)) = False) Or (IsEmpty(Cells(j + 1, 1)) = False)) Or (IsEmpty(Cells(j + 2, 1)) = False))
        Do While Worksheets("Plan_03DEC15").Cells(i, 1).Value <> Cells(j, 1).Value
            i = i + 1
        Loop
        If Worksheets("Plan_03DEC15").Cells(i + 1, 5).Value > 10 Then
            Cells(j, 3).Value = "TRANSFORMED!"
        End If
        j = j + 1
        Do While Left(Cells(j, 1), 1) <> "F" And (((IsEmpty(Cells(j, 1)) = False) Or (IsEmpty(Cells(j + 1, 1)) = False)) Or (IsEmpty(Cells(j + 2, 1)) = False))
            j = j + 1
        Loop
    Loop
End Sub
