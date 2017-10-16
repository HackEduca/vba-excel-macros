Attribute VB_Name = "zl07print_WASH_NDT"
' Selecting plan numbers (I.e. H0 or HX plans, w/ or w/o "/1" is ok) and print out WASH/INSPECT PACK within the plan.
' By default GenerateList() will return a list of cells of selection in the active worksheet.

Sub main()
    PrintPlans GenerateList()
End Sub

Sub PrintPlans(plans() As Variant, Optional DeleteAfterPrint As Boolean = True)             ' DeleteAfterPrint is off in XWB WASH/INSPECT PACK printing
    Dim r As String
    'Dim plans() As Variant
    'plans = GenerateList()
    
    If HasValue(plans) Then
        For Each Plan In plans
            If Not IsPlan(Plan) Then
                GoTo Escape
            End If
            plan_txt = plan_txt & Plan & ", "
        Next
        plan_txt = Left(plan_txt, Len(plan_txt) - 2)
        
        FunLoc = InputBox("Please etner the Functional Location which you want to create WASH/NDT PACKs in.")
        
        If MsgBox("Are you sure to create WASH/NDT PACKs of plan:" & vbCrLf & plan_txt & vbCrLf & "for Functional Location " & FunLoc & "?", vbYesNo) = vbNo Then GoTo Escape
        
        form_zl07print_WASH_NDT.Show  'get chkWASH & chkNDT
        
        If (Not form_zl07print_WASH_NDT.chkWASH) And (Not form_zl07print_WASH_NDT.chkNDT) Then GoTo Escape
        
        Set mysap = getSession()
        With mysap
            .SendCommand ("/nzl07")
            .FindById("wnd[0]/usr/ctxtTPLNR-LOW").text = FunLoc
            .FindById("wnd[0]").SendVKey 8
            
            For Each Plan In plans
                If Right(Plan, 2) <> "/1" Then Plan = Plan & "/1"
                .FindById("wnd[0]/usr/chk[0,4]").Selected = True
                .FindById("wnd[0]/mbar/menu[4]/menu[0]/menu[1]").Select
                .FindById("wnd[1]/usr/lbl[0,3]").SetFocus
                .FindById("wnd[1]").SendVKey 2
                .FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = Plan
                .FindById("wnd[1]/tbar[0]/btn[0]").press
                
                Set sapPage = getPackageSelection(mysap)
                AbsRow = 0
                PackageFound = False
                Do
                    MaxRow = (sapPage.Children.Count / 3) - 1
                    For i = 0 To MaxRow
                        package_txt = UCase(sapPage.GetCell(i, 1).text)
                        
                        LogicWASH = (InStr(package_txt, "WASH") <> 0)
                        LogicNDT = (InStr(package_txt, "INSPECT") <> 0)
                        
                        If (form_zl07print_WASH_NDT.chkWASH And LogicWASH) Or (form_zl07print_WASH_NDT.chkNDT And LogicNDT) Then
                            sapPage.GetAbsoluteRow(i).Selected = True
                            PackageFound = True
                        End If
                        AbsRow = AbsRow + 1
                    Next
                    If MaxRow = 7 Then
                        .FindById("wnd[2]/tbar[0]/btn[82]").press
                        Set sapPage = getPackageSelection(mysap)
                    End If
                Loop While MaxRow = 7
                
                If Not PackageFound Then
                    .FindById("wnd[2]/tbar[0]/btn[12]").press
                    .FindById("wnd[0]/tbar[0]/btn[3]").press
                    .FindById("wnd[1]/usr/btnSPOP-OPTION2").press
                Else
                    .FindById("wnd[2]/tbar[0]/btn[0]").press
                    .FindById("wnd[0]/tbar[0]/btn[11]").press
                    
                    ClearWindows mysap
                    
                    For Each item In mysap.FindById("wnd[0]/usr").Children
                        If item.text = Plan Then                                            'Find the row with item.text matches the Plan Name
                            r = Replace(Mid(item.ID, InStr(1, item.ID, ",") + 1), "]", "")  'Extract row number from item.ID
                            Exit For
                        End If
                    Next
                    .FindById("wnd[0]/usr/chk[0," & r & "]").Selected = True
                    .FindById("wnd[0]/mbar/menu[4]/menu[1]").Select
                    
                    'Extract PMO No. for deletion after printing
                    PMO = .FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR").text
                    
                    'Print PMO
                    .FindById("wnd[0]/tbar[1]/btn[26]").press
                    .FindById("wnd[1]/usr/radPMWO-FDWS").Select
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                    Set sapPage = .FindById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS")
                    For j = 0 To 1
                        For k = 6 To 8
                            sapPage.GetCell(j, k).Selected = True
                        Next
                    Next
                    Do While 1
                        If sapPage.GetCell(1, 2).text = "" Or errorstate Then
                            If outdevice = "" Or errorstate Then
                                outdevice = InputBox("Please input the printing device ID.", "No printer found")
                            End If
                            sapPage.GetCell(1, 2).text = outdevice
                        End If
                        .FindById("wnd[1]/tbar[0]/btn[8]").press
                        If .Children.Count = 3 Then
                            .FindById("wnd[2]/tbar[0]/btn[0]").press
                            errorstate = True
                            Set sapPage = .FindById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS")
                        Else
                            errorstate = False
                            Exit Do
                        End If
                    Loop
                    
                    ClearWindows mysap
                    
                    'Deleting the Order
                    If DeleteAfterPrint Then
                        For Each item In mysap.FindById("wnd[0]/usr").Children
                            If item.text = PMO Then         'Extracted PMO no. is used here and here only
                                r = Replace(Mid(item.ID, InStr(1, item.ID, ",") + 1), "]", "")
                                Exit For
                            End If
                        Next
                        .FindById("wnd[0]/usr/chk[0," & r & "]").Selected = True
                        .FindById("wnd[0]/mbar/menu[4]/menu[5]").Select
                        .FindById("wnd[1]/usr/btnBUTTON_1").press
                        .FindById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                End If
            Next
            
        End With
        
        Set mysap = Nothing
    Else
        MsgBox ("Please select the area containing plans which you would like to print.")
        GoTo Escape
    End If
    MsgBox ("Created.")
    Exit Sub
    
Escape:
    escmsg
    Exit Sub
End Sub
Sub ClearWindows(mysap As Variant)
    Do While mysap.Children.Count = 2
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
    Loop
End Sub

Function getPackageSelection(mysap As Variant) As Object
    Set getPackageSelection = mysap.FindById("wnd[2]/usr/tblSAPLIPM5TCTRL_0100")
End Function

Function GenerateList() As Variant()
    Dim tempARY() As Variant
    i = 0
    For Each cell In Selection
        plan_ref = cell
        If Not IsEmpty(plan_ref) Then
            ReDim Preserve tempARY(0 To i)
            tempARY(UBound(tempARY)) = plan_ref
            i = i + 1
        End If
    Next
    GenerateList = tempARY
End Function

