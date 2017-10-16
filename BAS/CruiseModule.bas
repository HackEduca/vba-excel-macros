Attribute VB_Name = "CruiseModule"
'
'
' A supplement module for snapshot search result tables
'
'
' By Cruising it means it will execute functions based on the rows from search results.
' It usually involves modifying data in SAP's H0/HI plans. So use at your own risk.

Sub main()
    Set mysap = getSession()
    For Each shit In ActiveWorkbook.Sheets
        If shit.Name <> "Result" Then Cruise_Result shit, mysap
    Next
    Set mysap = Nothing
End Sub

Sub Cruise_Result(ByRef shit As Variant, ByRef mysap As Variant, Optional j As Long = 2)
    
    shit.Activate
    mysap.StartTransaction ("ia06")
    
    Do While Not IsEmpty(shit.Cells(j, 4))
        
        k = j
        Do While shit.Cells(j, 4) = shit.Cells(k, 4)
            k = k + 1
        Loop
        
        k = k - 1

        ' SAP returns "System call failed" on object "RC271-PLNNR" when not loaded on time
        On Error Resume Next
        Do
            Err.Clear
            mysap.FindById("wnd[0]/usr/ctxtRC271-PLNNR").text = shit.Cells(j, 3)
            If Err.Number <> 0 Then Application.Wait [Now() + "0:00:02"]
        Loop Until Err.Number = 0
        On Error GoTo 0
        ' Seriously? Fuck SAP. So much for German engineering.
        
        'SAP cannot load plan on time while saving the plan
        Do While mysap.FindById("wnd[0]/usr").Children(1).Name = "RC271-PLNNR"
            mysap.FindById("wnd[0]").SendVKey 0
        Loop
        ' FFFFFFFFFFFFFFUUUUUUUUUUUUUUUUUU
        
        i = 0
        
        If InStr(1, mysap.Children(0).text, "Operation Overview") = 0 Then
            Do While 1
                If mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200").GetCell(i, 1).text = shit.Cells(k, 4) Then
                    mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200").GetAbsoluteRow(i).Selected = True
                    mysap.FindById("wnd[0]/mbar/menu[2]/menu[2]").Select
                    Exit Do
                End If
                i = i + 1
            Loop
        End If
        
        shit.Range(shit.Cells(j, 5), shit.Cells(k, 5)).Select
        SAP_ia06_SelectOps mysap, Selection                     ' Selecting OPs
        GotoLongText mysap, j                                   ' Go to Long Text Editor
        'GotoMaint mysap, j

        mysap.FindById("wnd[0]/tbar[0]/btn[11]").press
        
        mysap.ClearErrorList
        
        j = k + 1
    Loop
    
End Sub

Sub GotoMaint(ByRef mysap As Variant, ByRef j As Variant)
    lastop = "0000"
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press
    mysap.FindById("wnd[0]/tbar[1]/btn[26]").press
    Do While mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text <> lastop
        CopyMaint mysap, j, mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        lastop = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
    Loop
End Sub

Sub CopyMaint(ByRef mysap As Variant, ByVal j As Integer, CurrentOp As String)  'Variation of "Trans_package_copier"

    n = j
    Do While Cells(n, 5).text <> CurrentOp
        n = n + 1
    Loop
    mystr = ""

    ia = 0
    ja = 1
    
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    Do While mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & ia & "]").text <> ""
        If ia = 18 Then
            check = mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(ia - 1) & "]").text
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            If ja > 19 Then mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            If mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(ia - 1) & "]").text = check Then Exit Do
            ia = 0
        End If
    
        mystr = mystr & mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & ia & "]").text & ": "
        mystr = mystr & mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KTEX1[2," & ia & "]").text & " / "
        ia = ia + 1
        ja = ja + 1
    Loop
    Cells(n, 9) = mystr
End Sub
Public Sub Delete_Long_Text(ByRef mysap As Variant, str1 As String)
    'crude
    Dim i, page As Integer
    Dim ShrtTxt As String
    
    i = 1
    page = 0
    
    mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
    Do While 1 = 1
        If (mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i & "]").text = "") And (mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text = "") Then
            i = i - 1
            Exit Do
        Else
            If i <> 30 Then
                i = i + 1
            Else
                page = page + 1
                i = 1
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            End If
        End If
    Loop
    Do While (page >= 0 And i > 0)
        If mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text = str1 Then
            mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text = ""
        End If
        If i = 1 Then
            i = 29
            page = page - 1
            mysap.FindById("wnd[0]/tbar[0]/btn[81]").press
        Else
            i = i - 1
        End If
    Loop
End Sub
Sub GotoLongText(ByRef mysap As Variant, ByRef j As Variant)
    mysap.FindById("wnd[0]/tbar[1]/btn[16]").press
    Offset = 0
    Do While True
        'Measuring_Tool_Odyssey mysap, (j + Offset)                             ' Adding Measuring Tools to all eligible OPs
        Delete_Long_Text mysap, ActiveSheet.Cells((j + Offset), 1).text         ' Delete the row matches with the search result in SAP
        'highlight_if_keyword_exists mysap, "104", (j + Offset), , 255          ' Highlight the search hit if the keyword is found in the long text editor (Seriously?)
            'remove_raise_tsr mysap                                             ' outdated and hence broken, need fix (item)
            'find_the_shit mysap, item                                          ' outdated and hence broken, need fix (item)
            'breach_and_clear mysap                                             ' outdated and hence broken, need fix (item)
        
        Do
        Offset = Offset + 1
        Loop Until Cells(j + Offset - 1, 5) <> Cells(j + Offset, 5)
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press

        If mysap.Children.Count > 1 Then
            mysap.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
        Else
            Exit Do
        End If
    Loop
End Sub

Function highlight_if_keyword_exists(ByRef mysap As Variant, ByVal keyword As String, ByVal MyRow As Integer, Optional ByVal MyCol As Integer = 6, Optional ByVal colour As Long = 65535)
    'Stupid method, use as last resort
    mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
    i = 1
    Do While True
        If i > 30 Then
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            i = 2
        End If
        Set symbol = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i & "]")
        Set Content = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]")
        If InStr(1, Content.text, keyword) Then
            Cells(MyRow, MyCol).Interior.color = colour
            Exit Do
        ElseIf symbol.text = "" And Content.text = "" Then
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Function breach_and_clear(ByRef mysap As Variant)
    mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
End Function
Function remove_raise_tsr(ByRef mysap As Variant)
    mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
    mysap.FindById("wnd[0]/mbar/menu[1]/menu[4]").Select
    mysap.FindById("wnd[1]/usr/txtRSTXT-TXFISTRING").text = " OIN found expired."
    mysap.FindById("wnd[1]/usr/txtRSTXT-TXRESTRING").text = "."
    mysap.FindById("wnd[1]/tbar[0]/btn[5]").press
    mysap.FindById("wnd[0]/mbar/menu[1]/menu[4]").Select
    mysap.FindById("wnd[1]/usr/txtRSTXT-TXFISTRING").text = "Note: Please raise TSR if the."
    mysap.FindById("wnd[1]/usr/txtRSTXT-TXRESTRING").text = "."
    mysap.FindById("wnd[1]/tbar[0]/btn[5]").press
End Function
Function find_the_shit(ByRef mysap As Variant, item As Object)
    If MsgBox("Is this the shit?", vbYesNo) = vbYes Then
        item.Interior.color = 60000
    End If
End Function
Sub local_cruise()
    Set mysap = getSession()
    Cruise_Result ActiveSheet, mysap, 2
    Set mysap = Nothing
End Sub

Sub dumb_cruise()
    local_cruise
    ActiveWorkbook.Save
    Shutdown
End Sub

Sub Measuring_Tool_Odyssey(ByRef mysap As Variant, ByRef j As Variant)
    Header = "/:"
    MainTxt = "INCLUDE 'HK_RECORD_MEASURE TOOL S/N(GEN)' OBJECT TEXT ID ST"
    mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
    If mysap.Children.Count > 1 Then
        mysap.FindById("wnd[1]").Close
    End If
    i = 1
    AbsRow = 1
    found = 0
    LineRow = 1
    Set page = getEditor(mysap)
    Do While (page.GetCell(i, 0).text <> "") Or (page.GetCell(i, 2).text <> "")
        txtline = UCase(page.GetCell(i, 2).text)
        nextline = UCase(page.GetCell(i + 1, 2).text)
        PosRecord = InStr(1, txtline, "RECORD")
        PosResult = InStr(1, txtline, "RESULT")
        PosMo = InStr(1, txtline, ":")
        
        LogicM1 = (txtline = MainTxt)
        LogicM2 = (InStr(1, txtline, "MEASURE") <> 0) And ((InStr(1, txtline, "TOOL") <> 0) Or (InStr(1, txtline, "EQUIPMENT") <> 0)) And (InStr(1, txtline, ":") <> 0)
        LogicR1 = (PosRecord <> 0 And PosRecord < PosMo And (InStr(1, nextline, "FOR ") = 0 And InStr(1, nextline, "USE ") = 0 And (InStr(1, nextline, "ASSY") = 0 And (InStr(1, nextline, "ASSEMBLY") = 0))))
        LogicR2 = (InStr(1, txtline, "HK_RECORDING_TABLE_FRS3253") <> 0)
        LogicR3 = (InStr(1, txtline, "HK_RECORD_FRS3002") <> 0)
        LogicR4 = (txtline = "RECORD")
        LogicR5 = (PosResult <> 0 And PosResult < PosMo)
        LogicR6 = (txtline = "HK_RECORD_OP704_HARDNESS_TEST")
        
        LogicM = (LogicM1 Or LogicM2)
        LogicR = (LogicR1 Or LogicR2 Or LogicR3 Or LogicR4 Or LogicR5 Or LogicR6)
        
        If LogicM Then
            found = 2
            Exit Do
        ElseIf LogicR Then
            found = 1
            Exit Do
        End If
        If LineRow = 1 And InStr(1, txtline, "__") <> 0 Then LineRow = AbsRow
        AbsRow = AbsRow + 1
        If i = 29 Then
            i = 1
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Set page = getEditor(mysap)
        Else
            i = i + 1
        End If
    Loop
    
    If found = 0 Then
        AbsRow = LineRow
        Cells(j, 6).Interior.color = 65535  'Cannot find valid row to add into
    ElseIf found = 2 Then
        Cells(j, 6).Interior.color = 255    'Measuring Tool S/N Recording line already exists
        Exit Sub
    ElseIf found = 1 Then
        Cells(j, 6).Interior.color = 60000  'Sucessfully Added
    End If
    mysap.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
    mysap.FindById("wnd[1]/usr/txtRSTXT-TXLINENR").text = AbsRow - 1    'Before the "Record:" Hit
    mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
    Set page = getEditor(mysap)
    If page.GetCell(2, 0).text = "" Then
        mysap.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
        mysap.FindById("wnd[1]/usr/txtRSTXT-TXLINENR").text = 0
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
    End If
    mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
    Set page = getEditor(mysap)
    page.GetCell(2, 0).text = Header
    page.GetCell(2, 2).text = MainTxt
    
    Set page = Nothing
End Sub
