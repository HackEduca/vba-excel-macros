Attribute VB_Name = "SAP"
'When Britain first at Heaven's command
'Arose from out the azure main
'Arose, arose, arose from out the azure main!
'This was the charter, the charter of the Land,
'And guardian angels sang this strain:
'
'Rule Britannia! Britannia rule the waves!
'Britons never will be slaves!
'
'Rule Britannia! Britannia rule the waves!
'Britons never will be slaves!

' getSession() - Get SAPSession object.
' Optional variable "wnd": Selecting which SAPSession (0-5) to grab. Default is the first opened session (0). Will create new ones if the specified session is not opened.
' Optional variable "transaction": Check if the transaction code of the session grabbed matches the variable. This is used in conjuction with "NewSession".
' Optional variable "NewSession": This Boolean determines what will happen when the transcation of the grabbed SAPSession doesn't match the "transaction" variable.
'   When NewSession = False, Error message will pop up and the code would stop itself.
'   When NewSession = True, the code will send a command to change the transaction of the grabbed session to that specified by the "transaction" variable.

Public Function getSession(Optional wnd As Long = -1, Optional transaction As String = "", Optional NewSession As Boolean = False) As Object
    On Error Resume Next
    Do While 1
        Set SapGuiAuto = GetObject("SAPGUI")    'Get Built-in Object "SAPGUI"
        If Err.Number <> 0 Then 'If failed to get SAPGUI object, start up SAPlogon.exe, and get SAPGUI again
            Shell "C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe", vbMinimizedNoFocus
            Err.Number = 0
        Else
            Exit Do
        End If
    Loop
    On Error GoTo 0
    If Not IsObject(SAPApplication) Then
        Set SAPApplication = SapGuiAuto.GetScriptingEngine
        SAPApplication.AllowSystemMessages = False
    End If
    If SAPApplication.Connections.Count = 0 Then
        SAPApplication.OpenConnection ("1.1     PRD     [NEW] International Template")  '10.111.4.53
        NeedLogin = True
    End If
    If Not IsObject(SAPConnection) Then Set SAPConnection = SAPApplication.Children(0)  'Create connection
    If Not IsObject(mysap) Then
        If Not (wnd >= 0 And wnd <= 5) Then
            If Not NeedLogin Then
                If MsgBox("Default Window? [wnd = 0]", vbYesNo) = vbYes Then
                    wnd = 0
                Else
                    Do While Not (wnd >= 0 And wnd <= 5)
                        wnd = InputBox("Input the window you want to control. [0-5]")
                    Loop
                End If
            Else
                wnd = 0
            End If
        End If
        If SAPConnection.Children.Length < (wnd + 1) Then
            Set mysap = SAPConnection.Children(0)
            For i = SAPConnection.Children.Length To wnd
                mysap.createsession
                Do
                    Application.Wait [Now() + "0:00:01"]    'Fucking SAP can't even tell if it is busy, GuiSession.busy only considers in-transaction actions.
                Loop Until SAPConnection.Children.Length = (i + 1)
            Next
        End If
        Set mysap = SAPConnection.Children(CInt(wnd))       'THIS IS THE SHIT
    End If
    Set wnd0 = getWnd(mysap)
    wnd0.resizeWorkingPane 133, 34, False
    wnd0.height = 920
    wnd0.Width = 1000
    
    If NeedLogin Or IsTransaction(mysap, "S000") Then
        logged = AutoLogin(mysap, "4617871", "*KimJongIl420*")      'No, not your account
    Else
        logged = True
    End If
    If logged Then
        If NewSession Then
            mysap.SendCommand ("/n" & transaction)
        ElseIf transaction <> "" And UCase(transaction) <> mysap.Info.transaction Then
            MsgBox ("Incorrect SAP Transaction. Current transaction is " & mysap.Info.transaction & ".")
            GoTo Reject
        End If
        Set getSession = mysap      '..-. --- .-.. .-.. --- .-- / - .... . / .-. .- -... -... .. - / .... --- .-.. .
    Else
Reject:
        MsgBox "Failed to initialize. Terminating...", vbCritical, "Error"
    End If
    Set wnd0 = Nothing
End Function

' getXXX() - A cleaner way to pointing objects within SAPSessions. MUST reset the object everytime the UI of SAPSession is updated/changed.

Function getWnd(ByRef mysap As Variant, Optional i As Long = 0) As Object
    Set getWnd = mysap.Children(CInt(i))
End Function
Function getEditor(ByRef mysap As Variant) As Object
    Set getEditor = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA")
End Function
Function getIA06(ByRef mysap As Variant) As Object
    Set getIA06 = mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400")
End Function
Function getZL07(ByRef mysap As Variant) As Object
    Set getZL07 = mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010")
End Function

Function AutoLogin(ByRef mysap As Variant, Optional ID As String = "", Optional PW As String = "") As Boolean
    If ID = "" Then ID = InputBox("Please enter your staff ID.")
    If PW = "" Then PW = InputBox("Please enter the password.")
    mysap.FindById("wnd[0]/usr/txtRSYST-BNAME").text = ID
    mysap.FindById("wnd[0]/usr/pwdRSYST-BCODE").text = PW
    mysap.FindById("wnd[0]").SendVKey 0
    If mysap.Children.Count > 1 Then    'Terminate previous uncleared login and other pop-ups, if exists
        On Error Resume Next
        mysap.FindById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
        If Err.Number <> 0 Then
            mysap.FindById("wnd[1]").Close
            Err.Number = 0
        End If
        On Error GoTo 0
    ElseIf IsTransaction(mysap, "S000") Then    'Check if sucessfully logged on
        AutoLogin = False
        Exit Function
    End If
    AutoLogin = True
End Function

'DEB
Sub SAP_Extract_Ops_From_FRS()
    On Error Resume Next
    
    Dim FRS, SPMO As String
    Dim k, j, i As Integer
    
    FRS = InputBox("Please enter the Repair you wish to extract.")

    FRS = UCase(FRS)
    Cells(1, 1).Value = FRS
'    SPMO = InputBox("Please enter the Order you wish to search in.")
    
    Set mysap = getSession(, "sq01", True)
    
    '"/nsq01":DB
    
    mysap.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select
    mysap.FindById("wnd[1]/usr/radRAD1").Select
    mysap.FindById("wnd[1]/tbar[0]/btn[2]").press
    mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
    mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = 4
    mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "4"
    mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
    mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DB"
    mysap.FindById("wnd[0]").SendVKey 8
    mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
    mysap.FindById("wnd[0]/usr/txtSP$00001-LOW").text = "*" & FRS & "*"
    mysap.FindById("wnd[0]").SendVKey 8
    mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
    
    Set sapPage = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100")
    If sapPage.GetCell(0, 0).text = "" Then
        MsgBox ("No Package has been found. The program will now end.")
        Exit Sub
    End If
    
    k = 0
    j = 1
    MaxR = sapPage.VisibleRowCount
    
    Do While sapPage.Rows(k).Count <> 0
                
        If k > MaxR Then
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Set sapPage = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100")
            k = 0
        End If
        
        If Left(sapPage.GetCell(k, 0).text, 1) = "H" Then
            Cells(j, 3) = sapPage.GetCell(k, 0).text
            Cells(j, 4) = sapPage.GetCell(k, 1).text
            Cells(j, 5) = sapPage.GetCell(k, 2).text
            j = j + 1
        End If
        k = k + 1
    Loop
    
    '"/nsq01":DB
    
    '"/nip03"
    j = 1
    
    Do While IsEmpty(Cells(j, 3)) = False
        If j <> 1 Then
            If Cells(j, 3).Value <> Cells(j - 1, 3).Value Then
                mysap.SendCommand ("/nip03")
                mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Cells(j, 3) & "/1"
                mysap.FindById("wnd[0]").SendVKey 0
                Cells(j, 6) = "Active Group Counter:"
                Cells(j, 7).Formula = "=""" & mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text & """"
            Else
                Cells(j, 6) = "Active Group Counter:"
                Cells(j, 7).Formula = Cells(j - 1, 7).Formula
            End If
        Else
            mysap.SendCommand ("/nip03")
            mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Cells(j, 3) & "/1"
            mysap.FindById("wnd[0]").SendVKey 0
            Cells(j, 6) = "Active Group Counter:"
            Cells(j, 7).Value = "=""" & mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text & """"
        End If
        j = j + 1
    Loop
    
    '"/nip03"
    
    '"/nsq01":DC
    
    j = 1
    i = 0
    Do While IsEmpty(Cells(j, 3)) = False
        
        mysap.SendCommand ("/nsq01")
        mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DC"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtPACKAGE-LOW").text = Cells(j, 4).Value
        mysap.FindById("wnd[0]/usr/ctxtTASKLIST").text = Cells(j, 3).Value
        mysap.FindById("wnd[0]/usr/btn%_PACKAGE_%_APP_%-VALU_PUSH").press
        mysap.FindById("wnd[1]/tbar[0]/btn[16]").press
        mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 4).Value
        
        Do While Cells(j, 3).Value = Cells(j + 1, 3).Value
            j = j + 1
            mysap.FindById("wnd[1]/tbar[0]/btn[13]").press
            mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 4).Value
        Loop
        
        mysap.FindById("wnd[1]").SendVKey 8
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        i = i + 1
        Cells(i, 9).Value = Cells(j, 3).text
        i = i + 1
        Cells(i, 9).Value = "Group Counter"
        Cells(i, 10).Value = "Operation"
        Cells(i, 11).Value = "Workcentre"
        Cells(i, 12).Value = "Short Text"
        Cells(i, 13).Value = "Packages"
        i = i + 1
        k = 0
        
        Do While Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text, 1) <> "_"
            If k = 30 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                k = 0
            End If
            If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text = Cells(j, 7).text Then
                Cells(i, 9).Formula = "=""" & mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text & """"
                Cells(i, 10).Formula = "=""" & mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-VORNR[3," & CStr(k) & "]").text & """"
                Cells(i, 12).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-LTXA1[5," & CStr(k) & "]").text
                Cells(i, 13).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PACKAGES[6," & CStr(k) & "]").text
                i = i + 1
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    
    '"/nsq01":DC
    
    '"/nsq01":DE
    j = 1
    i = 0
    Do While IsEmpty(Cells(j, 3)) = False
        
        mysap.SendCommand ("/nsq01")
        mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DE"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtPACKAGE-LOW").text = Cells(j, 4).Value
        mysap.FindById("wnd[0]/usr/ctxtTASKLIST").text = Cells(j, 3).Value
        mysap.FindById("wnd[0]/usr/btn%_PACKAGE_%_APP_%-VALU_PUSH").press
        mysap.FindById("wnd[1]/tbar[0]/btn[16]").press
        mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 4).Value
        
        Do While Cells(j, 3).Value = Cells(j + 1, 3).Value
            j = j + 1
            mysap.FindById("wnd[1]/tbar[0]/btn[13]").press
            mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 4).Value
        Loop
        
        mysap.FindById("wnd[1]").SendVKey 8
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        i = i + 3
        k = 0
        
        Do While Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[1," & CStr(k) & "]").text, 1) <> "_"
            If k = 30 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                k = 0
            End If
            If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[1," & CStr(k) & "]").text = Cells(j, 7).text Then
                Cells(i, 11).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-ARBPL[4," & CStr(k) & "]").text
                i = i + 1
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    
    
    '"/nsq01":DE
    
    Cells.Columns.AutoFit
    If IsEmpty(Cells(1, 3)) Then
        MsgBox ("No Package has been found. The program will now end.")
    Else
        MsgBox "Done."
    End If
End Sub

Sub SAP_ia06_longtext_copier(ByRef mysap As Variant, Optional SuppressNoti As Boolean = False)

    Dim i, j, PageRow, r, c, AbsRow, OpCount, RowPerOp, LoopCount As Integer

    r = 1
    c = 1
    PageRow = 0
    AbsRow = 0
    RowPerOp = 5
    
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    Set page = getIA06(mysap)
    
    Do While page.GetCell(PageRow, 2).text <> "" And page.GetCell(PageRow, 5).text <> ""

        If page.GetAbsoluteRow(AbsRow).Selected = True Then
            Cells(r, c + 0).Formula = "= """ & page.GetCell(PageRow, 0).text & """"
            Cells(r, c + 1).Value = page.GetCell(PageRow, 2).text
            c = c + RowPerOp
        End If
        If PageRow = 22 Then
            PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Set page = getIA06(mysap)
        Else
            PageRow = PageRow + 1
        End If
        AbsRow = AbsRow + 1
    Loop
    
    j = 1
    OpCount = (c - 1) / RowPerOp
    c = 1
    LoopCount = OpCount
    
    mysap.FindById("wnd[0]/tbar[1]/btn[16]").press
    
    Do While LoopCount > 0
        i = 1
        r = 1
        
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
        
        Set page = getEditor(mysap)
        
        Do While 1 = 1
            LineHdr = page.GetCell(i, 0).text
            LineTxt = page.GetCell(i, 2).text
            If (LineHdr = "") And (LineTxt = "") Then
                i = i - 1
                Exit Do
            Else
                Cells(r, c + 2).Value = LineHdr
                If Len(LineTxt) <> 72 Then
                    Cells(r, c + 3).Value = LineTxt
                Else
                    a = 36
                    Do While a < 72
                        If Mid(LineTxt, a, 1) = " " Then Exit Do
                        a = a + 1
                    Loop
                    If a <> 72 Then
                        Cells(r, c + 3).Value = Left(LineTxt, a)
                        r = r + 1
                        Cells(r, c + 3).Value = Right(LineTxt, 72 - a)
                    Else
                        Cells(r, c + 3).Value = LineTxt
                    End If
                    If IsEmpty(Cells(r, c + 3)) And IsEmpty(Cells(r, c + 4)) Then
                        Range(Cells(r, c + 3), Cells(r, c + 4)).Delete xlShiftUp
                        r = r - 1
                    End If
                End If

                If i <> 30 Then
                    i = i + 1
                Else
                    i = 2
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    Set page = getEditor(mysap)
                End If

                r = r + 1
            End If
        Loop
        
        If LoopCount <> 1 Then
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
            mysap.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
        Else
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
        
        j = j + 1
        c = c + RowPerOp
        LoopCount = LoopCount - 1
    Loop
   
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_FHM").press
    
    c = 1
    LoopCount = OpCount
    
    
    Do While LoopCount > 0
        If mysap.Children.Count > 1 Then mysap.FindById("wnd[1]").Close

        Set page = mysap.FindById("wnd[0]/usr/tblSAPLCFDITCTRL_0102")
        r = page.Children.Count / 20
        If r <> 0 Then
            For PRT = 1 To r
                Cells(PRT, c + 4).Value = page.GetCell(PRT - 1, 2).text
                Cells(PRT, c + 4).Value = Cells(PRT, c + 4).Value & " +" & page.GetCell(PRT - 1, 8).text
            Next
        End If
        
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
        c = c + RowPerOp
        LoopCount = LoopCount - 1
    Loop
    
    mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    
    Cells.Columns.AutoFit
    If Not SuppressNoti Then MsgBox "Successfully copied."
End Sub

Sub SAP_ia06_longtext_paster(ByRef mysap As Variant, Optional SupressNoti As Boolean = False)

    'On Error Resume Next

    Dim i, j, PageRow, page, r, c, AbsRow, OpCount, RowPerOp, LoopCount As Integer
    Dim ShortTxt As String
    
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    c = 1
    RowPerOp = 5
    OpCount = Application.WorksheetFunction.RoundUp((ActiveSheet.UsedRange.Columns.Count) / RowPerOp, 0)
    
    LoopCount = OpCount
    
    Do While LoopCount <> 0
        i = 0
        j = 0
        r = 1
        mysap.FindById("wnd[0]/tbar[0]/btn[83]").press
        mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
        
        Set mypage = getIA06(mysap)
        Do While 1 = 1
            If (mypage.GetCell(i, 2).text <> "" And mypage.GetCell(i, 5).text <> "") Then
                If i <> 22 Then
                    i = i + 1
                Else
                    i = 0
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                End If
            ElseIf mypage.GetCell(i, 0).text = "9999" Then
                i = 1
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set mypage = getIA06(mysap)
            Else
                mypage.GetCell(i, 0).text = Cells(r, c).Value
                mypage.GetCell(i, 2).text = Cells(r, (c + 1)).Value
                mypage.GetCell(i, 8).text = "MIN"
                mypage.GetCell(i, 9).text = "1"
                mypage.GetCell(i, 11).text = "H"
                mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
                Set mypage = getIA06(mysap)
                Exit Do
            End If
        Loop
        
        i = 0
        Do While mypage.GetCell(i, 2).text <> ""
            If mypage.GetCell(i, 0).text = Cells(r, c).Value Then
                mypage.GetCell(i, 6).SetFocus
                mysap.FindById("wnd[0]").SendVKey 2
                Exit Do
            Else
                If i <> 22 Then
                    i = i + 1
                Else
                    i = 0
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    Set mypage = getIA06(mysap)
                End If
            End If
        Loop
        
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select

        AbsRow = 1
        
        Do While (IsEmpty(Cells(AbsRow, c + 2)) = False Or IsEmpty(Cells(AbsRow, c + 3)) = False)
            If AbsRow <> 1 Then
                mysap.FindById("wnd[0]").SendVKey 0
            End If
            AbsRow = AbsRow + 1
        Loop
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        Set mypage = getEditor(mysap)
        
        PageRow = 1
        AbsRow = 1
        
        Do While (IsEmpty(Cells(AbsRow, c + 2)) = False Or IsEmpty(Cells(AbsRow, c + 3)) = False)
            If PageRow = 30 Then
                PageRow = 1
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set mypage = getEditor(mysap)
            End If
            mypage.GetCell(PageRow, 0).text = Cells(AbsRow, c + 2).Value
            mypage.GetCell(PageRow, 2).text = Cells(AbsRow, c + 3).Value
            PageRow = PageRow + 1
            AbsRow = AbsRow + 1
        Loop
        
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        
        Do While IsEmpty(Cells(r, c + 4)) = False
            j = 1
            mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_FHM").press
            mysap.FindById("wnd[1]/tbar[0]/btn[8]").press
            If InStr(1, Cells(r, c + 4).Value, "DAT") <> 0 Then
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-DOKNR").text = Mid(Cells(r, c + 4).Value, 1, (InStr(1, Cells(r, c + 4).Value, " ") - 1))
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-DOKAR").text = "DAT"
                mysap.FindById("wnd[1]/usr/txtPLFHD-DOKTL").text = "000"
                mysap.FindById("wnd[1]/usr/txtPLFHD-DOKVR").text = "01"
            ElseIf InStr(1, Cells(r, c + 4).Value, "REC") <> 0 Then
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-DOKNR").text = Mid(Cells(r, c + 4).Value, 1, (InStr(1, Cells(r, c + 4).Value, " ") - 1))
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-DOKAR").text = "REC"
                mysap.FindById("wnd[1]/usr/txtPLFHD-DOKTL").text = "000"
                mysap.FindById("wnd[1]/usr/txtPLFHD-DOKVR").text = "01"
            Else
                mysap.FindById("wnd[1]/tbar[0]/btn[6]").press
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-MATNR").text = Mid(Cells(r, c + 4).Value, 1, (InStr(1, Cells(r, c + 4).Value, " ") - 1))
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-FHWRK").text = "HK01"
            End If
            mysap.FindById("wnd[1]/usr/ctxtAFFHD-MGEINH").text = Mid(Cells(r, c + 4).Value, InStr(1, Cells(r, c + 4).Value, "+"))
            mysap.FindById("wnd[1]/usr/ctxtPLFHD-STEUF").text = "1"
            If IsEmpty(Cells(r + 1, c + 4)) Then
                mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
                Exit Do
            Else
                mysap.FindById("wnd[1]/tbar[0]/btn[5]").press
                r = r + 1
            End If
        Loop
        If j = 1 Then
            mysap.FindById("wnd[1]").Close
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
            mysap.FindById("wnd[1]/usr/btnSPOP-OPTION2").press
        End If
        
        c = c + RowPerOp
        LoopCount = LoopCount - 1
        
    Loop
    Set mypage = Nothing
    If Not SupressNoti Then MsgBox "Successfully pasted."
End Sub

Sub SAP_zl07TECO_longtext_copier(ByRef mysap As Variant, Optional SuppressNoti As Boolean = False)
    On Error Resume Next

    Dim i, j, PageRow, page, r, c, AbsRow, OpCount, RowPerOp, LoopCount As Integer
    Dim ShortTxt As String
    
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    r = 1
    c = 1
    PageRow = 0
    AbsRow = 0
    RowPerOp = 5
    
    Set sapPage = getZL07(mysap)
    'Do While (Left(sappage.GetCell(PageRow, 2).Text, 1) <> "_" And Left(sappage.GetCell(PageRow, 7).Text, 1) <> "_")
    Do While Err.Number = 0
        If sapPage.GetAbsoluteRow(AbsRow).Selected = True Then
            Cells(r, c + 0).Formula = "= """ & sapPage.GetCell(PageRow, 0).text & """"
            Cells(r, c + 1).Value = sapPage.GetCell(PageRow, 2).text
            c = c + RowPerOp
        End If
        If PageRow = 22 Then
            PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Set sapPage = getZL07(mysap)
        Else
            PageRow = PageRow + 1
        End If
        AbsRow = AbsRow + 1
    Loop

    Err.Number = 0
    
    j = 1
    OpCount = (c - 1) / RowPerOp
    c = 1
    LoopCount = OpCount

    Do While LoopCount > 0
        i = 1
        page = 0
        r = 1
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        Set sapPage = getZL07(mysap)
        PageRow = 0
        Do While sapPage.GetCell(PageRow, 0).text <> Cells(r, c).text
            If PageRow = 22 Then
                PageRow = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set sapPage = getZL07(mysap)
            Else
                PageRow = PageRow + 1
            End If
        Loop
        
        sapPage.GetCell(PageRow, 8).SetFocus
        sapPage.GetCell(PageRow, 8).press

        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
        Set sapPage = getEditor(mysap)
        Do While 1 = 1
            
            If (Left(sapPage.GetCell(i, 0).text, 1) = "_") And (Left(sapPage.GetCell(i, 2).text, 1) = "_") Then
            'If Err.Number <> 0 Then
                i = i - 1
                Err.Number = 0
                Exit Do
            Else
                Cells(r, c + 2).Value = sapPage.GetCell(i, 0).text
                If Len(sapPage.GetCell(i, 2).text) <> 72 Then
                    Cells(r, c + 3).Value = sapPage.GetCell(i, 2).text
                Else
                    a = 36
                    Do While a <= 72
                        If Mid(sapPage.GetCell(i, 2).text, a, 1) = " " Then
                            Exit Do
                        End If
                        a = a + 1
                    Loop
                    If i <> 72 Then
                        Cells(r, c + 3).Value = Left(sapPage.GetCell(i, 2).text, a)
                        r = r + 1
                        Cells(r, c + 3).Value = Right(sapPage.GetCell(i, 2).text, 72 - a)
                    Else
                        Cells(r, c + 3).Value = sapPage.GetCell(i, 2).text
                    End If
                    If IsEmpty(Cells(r, c + 3)) And IsEmpty(Cells(r, c + 4)) Then
                        Range(Cells(r, c + 3), Cells(r, c + 4)).Delete xlShiftUp
                        r = r - 1
                    End If
                End If
                
                If i <> 30 Then
                    i = i + 1
                Else
                    page = page + 1
                    i = 2
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    Set sapPage = getEditor(mysap)
                End If

                r = r + 1
            End If
        Loop

        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
        
        r = 1
        Set sapPage = mysap.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0102")
        'Do While Left(sappage.GetCell((r - 1), 2).Text, 1) <> "_"
        Do While Err.Number = 0
            Cells(r, c + 4).Value = sapPage.GetCell((r - 1), 2).text
            Cells(rT, c + 4).Value = Cells(r, c + 4).Value & " +" & sapPage.GetCell((r - 1), 8).text
            r = r + 1
        Loop
        
        Err.Number = 0
        
        j = j + 1
        c = c + RowPerOp
        LoopCount = LoopCount - 1
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    Loop
    
    Cells.Columns.AutoFit
    If Not SuppressNoti Then MsgBox "Successfully copied."
End Sub

Sub Trans_Component_Copier()
    On Error Resume Next
    Dim i, j, k As Integer
    Dim check As String
    
    Set mysap = getSession()
    
    mysap.FindById("wnd[0]").resizeWorkingPane 105, 31, False
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_MAT").press
    k = 1
    
    Do While 1
        Cells(1, k).Formula = "= """ & mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text & """"
        
        i = 0
        j = 1
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i) & "]").text <> ""
            If i = 20 Then
                check = mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i - 1) & "]").text
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                If mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i - 1) & "]").text = check Then
                    Exit Do
                End If
                i = 0
            End If
        
            Cells(j, k + 1).Formula = "= """ & mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i) & "]").text & """"
            Cells(j, k + 2).Formula = "= """ & mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/txtRIHSTPX-MENGE[1," & CStr(i) & "]").text & """"
            i = i + 1
            j = j + 1
        Loop
        
        check = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
        
        If mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text = check Then Exit Do
        k = k + 3
    Loop

End Sub

Sub Trans_Component_Paster()
    On Error Resume Next
    Dim i, j, k As Integer
    Dim check As String
    
    Set mysap = getSession()

    mysap.FindById("wnd[0]").resizeWorkingPane 105, 31, False
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_MAT").press
    k = 1

    Do While IsEmpty(Cells(1, k)) = False

        i = 0
        j = 1
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While IsEmpty(Cells(j, (k + 1))) = False
            If i = 20 Then
                mysap.FindById("wnd[0]").SendVKey 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                i = 0
            End If

            mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i) & "]").text = Cells(j, (k + 1)).Value
            mysap.FindById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/txtRIHSTPX-MENGE[1," & CStr(i) & "]").text = Cells(j, (k + 2)).Value
            i = i + 1
            j = j + 1
        Loop

        mysap.FindById("wnd[0]").SendVKey 0
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press

        k = k + 3

    Loop

End Sub

Sub Trans_Package_Copier()
    On Error Resume Next
    Dim i, j, k As Integer
    Dim check As String
    
    Set mysap = getSession()
    
    mysap.FindById("wnd[0]").resizeWorkingPane 105, 31, False
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press
    mysap.FindById("wnd[0]/tbar[1]/btn[26]").press
    k = 1
    
    Do While 1 = 1
        
        Cells(1, k).Formula = "= """ & mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text & """"
        
        i = 0
        j = 1
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(i) & "]").text <> ""
            If i = 19 Then
                check = mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(i - 1) & "]").text
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                If j > 20 Then
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                End If
                If mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(i - 1) & "]").text = check Then
                    Exit Do
                End If
                i = 0
            End If
        
            Cells(j, k + 1).Formula = "= """ & mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(i) & "]").text & """"
            i = i + 1
            j = j + 1
        Loop
        
        check = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
        
        If mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text = check Then
            Exit Do
        End If
        k = k + 2
        
    Loop

End Sub

Sub Trans_Package_Paster()
    On Error Resume Next
    Dim i, j, k As Integer
    Dim check As String
    
    Set mysap = getSession()

    mysap.FindById("wnd[0]").resizeWorkingPane 105, 31, False
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press
    mysap.FindById("wnd[0]/tbar[1]/btn[26]").press
    k = 1

    Do While IsEmpty(Cells(1, k)) = False

        i = 0
        j = 1
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While IsEmpty(Cells(j, (k + 1))) = False
            If i = 19 Then
                mysap.FindById("wnd[0]").SendVKey 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                i = 0
            End If

            mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(i) & "]").text = Cells(j, (k + 1)).Value
            i = i + 1
            j = j + 1
        Loop

        mysap.FindById("wnd[0]").SendVKey 0
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press

        k = k + 2

    Loop

End Sub

Sub Temp_Package_Stripper()

    On Error Resume Next
    Dim check As String
    
    Set mysap = getSession()
    
    Do While mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text <> check
        check = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/tbar[1]/btn[8]").press
        mysap.FindById("wnd[0]/tbar[1]/btn[14]").press
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
    Loop
    
End Sub

Sub Operation_Long_Text_Finder()
    On Error Resume Next

    Dim i, page As Integer
    Dim str1, str2, ShrtTxt As String
    
    str1 = InputBox("Text to be replaced.")
    str2 = InputBox("Text to replace the text.")
'    Str1 = Range("A1").Value
'    Str2 = Range("A2").Value
    
    Set mysap = getSession()

    mysap.FindById("wnd[0]/tbar[1]/btn[16]").press
    
    Do While 1 = 1
        i = 1
        page = 0
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select

        If mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text = ShortTxt Then
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
            mysap.FindById("wnd[0]/tbar[0]/btn[11]").press
            Exit Do
        End If

        Do While 1 = 1

            If (mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i & "]").text = "") And (mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text = "") Then
                i = i - 1
                Exit Do
            Else
                If i <> 26 Then
                    i = i + 1
                Else
                    page = page + 1
                    i = 1
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                End If
            End If
        Loop

        Do While (page >= 0 And i > 0)

            If InStr(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text, "OP") <> 0 Then
                Range("B1") = Range("B1") & mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text & ", "
            End If
            If i = 1 Then
                i = 26
                page = page - 1
                mysap.FindById("wnd[0]/tbar[0]/btn[81]").press
            Else
                i = i - 1
            End If
        Loop
        ShrtTxt = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        mysap.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
    Loop
End Sub

Public Sub Operation_Long_Text_Replacer_callable()
    Operation_Long_Text_Replacer
End Sub

Public Sub Operation_Long_Text_Replacer(Optional mysap, Optional str1 As String = "*no*input*", Optional str2 As String = "*no*input*")
    On Error Resume Next

    Dim i, j, page As Integer
    Dim ShrtTxt As String
    
    Do While str1 = "*no*input*"
        str1 = InputBox("Text to be replaced.")
    Loop
    Do While str2 = "*no*input*"
        str2 = InputBox("Text to replace the text.")
    Loop
    
    If Not IsObject(mysap) Then Set mysap = getSession()
    
    j = 0
    i = 0
    k = 0
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    Do While mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txtPLPOD-VORNR[0," & CStr(i) & "]").text <> "" And mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txtPLPOD-LTXA1[5," & i & "]").text <> ""
        If mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").GetAbsoluteRow(k).Selected = True Then
            j = j + 1
        End If
        If i = 22 Then
            i = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
        Else
            i = i + 1
        End If
        k = k + 1
    Loop
    
    
    Do While j <> 0
        i = 1
        page = 0
        
        mysap.FindById("wnd[0]/tbar[1]/btn[16]").press
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
        
            mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text = Replace(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text, str1, str2)
            If i = 1 Then
                i = 29
                page = page - 1
                mysap.FindById("wnd[0]/tbar[0]/btn[81]").press
            Else
                i = i - 1
            End If
        Loop
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        
        If j <> 1 Then
            mysap.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
        End If
        
        j = j - 1
    Loop
    
End Sub

Sub Trans_PRT_Stripper()
    On Error Resume Next
    
    Dim i, j, k As Integer
    Dim check As String
    
    Set mysap = getSession()
    
    mysap.FindById("wnd[0]").resizeWorkingPane 105, 31, False
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_FHM").press

    Do While mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text <> check
        mysap.FindById("wnd[1]").Close
        i = 0
        k = CInt(mysap.FindById("wnd[0]/usr/txtRC27X-ENTRIES").text)
        mysap.FindById("wnd[0]/tbar[1]/btn[33]").press
        Do While i < k
            j = 1
            If mysap.FindById("wnd[0]/usr/tblSAPLCFDITCTRL_0102/ctxtPLFHD-FHMAR[1," & CStr(i) & "]").text <> "D" Then
                mysap.FindById("wnd[0]/usr/tblSAPLCFDITCTRL_0102").GetAbsoluteRow(i).Selected = False
            End If
            Do While IsEmpty(Cells(j, 1)) = False
                If CStr(Cells(j, 1).Value) = Left(mysap.FindById("wnd[0]/usr/tblSAPLCFDITCTRL_0102/txtPLFHD-FHMNR[2," & CStr(i) & "]").text, Len(CStr(Cells(j, 1).Value))) Then
                    mysap.FindById("wnd[0]/usr/tblSAPLCFDITCTRL_0102").GetAbsoluteRow(i).Selected = False
                    Exit Do
                End If
                j = j + 1
            Loop
            i = i + 1
        Loop
        
        If k <> 0 Then
            mysap.FindById("wnd[0]/tbar[1]/btn[14]").press
            mysap.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        End If
        
        check = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
    Loop

End Sub

Sub SAP_zl07_longtext_copier(ByRef mysap As Variant, Optional SuppressNoti As Boolean = False)
    On Error Resume Next

    Dim i, j, PageRow, page, r, c, AbsRow, OpCount, RowPerOp, LoopCount As Integer
    Dim ShortTxt As String
    
    r = 1
    c = 1
    PageRow = 0
    AbsRow = 0
    RowPerOp = 5
    
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press

    Do While (mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-ARBPL[2," & PageRow & "]").text <> "" And mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7," & PageRow & "]").text <> "")

        If mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").GetAbsoluteRow(AbsRow).Selected = True Then
            Cells(r, c + 0).Formula = "= """ & mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0," & PageRow & "]").text & """"
            Cells(r, c + 1).Value = mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-ARBPL[2," & PageRow & "]").text
            c = c + RowPerOp
        End If
        If PageRow = 22 Then
            PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
        Else
            PageRow = PageRow + 1
        End If
        AbsRow = AbsRow + 1
    Loop

    j = 1
    OpCount = (c - 1) / RowPerOp
    c = 1
    LoopCount = OpCount

    Do While LoopCount > 0
        i = 1
        page = 0
        r = 1
        
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        PageRow = 0
        Do While mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0," & PageRow & "]").text <> Cells(r, c).text
            If PageRow = 22 Then
                PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Else
                PageRow = PageRow + 1
            End If
        Loop
        
        mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8," & PageRow & "]").SetFocus
        mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8," & PageRow & "]").press

        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select

        Do While 1 = 1

            If (mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i & "]").text = "") And (mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text = "") Then
                i = i - 1
                Exit Do
            Else
                Cells(r, c + 2).Value = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i & "]").text
                If Len(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text) <> 72 Then
                    Cells(r, c + 3).Value = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text
                Else
                    a = 36
                    Do While a <= 72
                        If Mid(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text, a, 1) = " " Then
                            Exit Do
                        End If
                        a = a + 1
                    Loop
                    If a <> 72 Then
                        Cells(r, c + 3).Value = Left(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text, a)
                        r = r + 1
                        Cells(r, c + 3).Value = Right(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text, 72 - a)
                    Else
                        Cells(r, c + 3).Value = mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").text
                    End If
                    If IsEmpty(Cells(r, c + 3)) And IsEmpty(Cells(r, c + 4)) Then
                        Range(Cells(r, c + 3), Cells(r, c + 4)).Delete xlShiftUp
                        r = r - 1
                    End If
                End If
                If i <> 30 Then
                    i = i + 1
                Else
                    page = page + 1
                    i = 2
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                End If

                r = r + 1
            End If
        Loop

        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
        mysap.FindById("wnd[1]").Close
        r = 1
        Set sapPage = mysap.FindById("wnd[0]/usr/tblSAPLCOFUTCTRL_0102")
        
        Err.Number = 0
        Do While Err.Number = 0 'Left(sappage.GetCell((r - 1), 2).Text, 1) <> "_"
            Cells(r, c + 4).Value = sapPage.GetCell((r - 1), 2).text
            Cells(r, c + 4).Value = Cells(r, c + 4).Value & " +" & sapPage.GetCell((r - 1), 7).text
            r = r + 1
            Debug.Print sapPage.GetCell((r - 1), 2).text
        Loop
        Err.Number = 0
        
        j = j + 1
        c = c + RowPerOp
        LoopCount = LoopCount - 1
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    Loop
    
    Cells.Columns.AutoFit
    If Not SuppressNoti Then MsgBox "Successfully copied."
End Sub

Sub SAP_ia06_SelectOps(ByRef mysap As Variant, oplist)
    
    Dim PageRow, AbsRow As Integer

    PageRow = 0

    mysap.FindById("wnd[0]/tbar[1]/btn[8]").press
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    op = unique_list(oplist)
    
    Set ia06 = getIA06(mysap)
    
    For Each i In op
        Do While 1
            If PageRow = 23 Then
                PageRow = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set ia06 = getIA06(mysap)
            End If
            
            If ia06.GetCell(PageRow, 0).text = i Then
                ia06.Rows(PageRow).Selected = True
                Exit Do
            End If
            PageRow = PageRow + 1
        Loop
    Next
End Sub

' DEB
Sub SAP_Extract_Ops_From_H0Plan()
    On Error Resume Next
    
    Dim Plan, Package As String
    Dim k, j, i As Integer
    
    Set mysap = getSession()

    i = 1
    
    Plan = InputBox("Please enter the Plan you wish to extract.")
    
    Plan = UCase(Plan)
    Cells(i, 1).Value = Plan
    
MorePackage:
    Package = InputBox("Please enter the Package you wish to extract.")
    If i <> 1 Then Cells(i, 1).Value = Cells(i - 1, 1).Value
    Cells(i, 2).Value = Package
    
    If MsgBox("Do you want to add more packages?", vbYesNo) = vbYes Then
        i = i + 1
        GoTo MorePackage
    End If
    
    '"/nip03"
    j = 1
    
    Do While IsEmpty(Cells(j, 2)) = False
        If j <> 1 Then
            If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
                mysap.SendCommand ("/nip03")
                mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Plan & "/1"
                mysap.FindById("wnd[0]").SendVKey 0
                Cells(j, 3).Value = "Active Group Counter:"
                Cells(j, 4).Value = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text
            Else
                Cells(j, 3).Value = "Active Group Counter:"
                Cells(j, 4).Value = Cells(j - 1, 4).Value
            End If
        Else
            mysap.SendCommand ("/nip03")
            mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Plan & "/1"
            mysap.FindById("wnd[0]").SendVKey 0
            Cells(j, 3).Value = "Active Group Counter:"
            Cells(j, 4).Value = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text
        End If
        j = j + 1
    Loop
    
    '"/nip03"
    
    '"/nsq01":DC
    
    j = 1
    i = 0
    Do While IsEmpty(Cells(j, 3)) = False
    
        mysap.SendCommand ("/nsq01")
        mysap.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select
        mysap.FindById("wnd[1]/usr/radRAD1").Select
        mysap.FindById("wnd[1]/tbar[0]/btn[2]").press
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
        mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = 4
        mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "4"
        mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
        mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DC"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtPACKAGE-LOW").text = Cells(j, 2).Value
        mysap.FindById("wnd[0]/usr/ctxtTASKLIST").text = Cells(j, 1).Value
        mysap.FindById("wnd[0]/usr/btn%_PACKAGE_%_APP_%-VALU_PUSH").press
        mysap.FindById("wnd[1]/tbar[0]/btn[16]").press
        mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 2).Value
        
        Do While Cells(j, 3).Value = Cells(j + 1, 3).Value
            j = j + 1
            mysap.FindById("wnd[1]/tbar[0]/btn[13]").press
            mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 2).Value
        Loop
        
        mysap.FindById("wnd[1]").SendVKey 8
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        i = i + 1
        Cells(i, 6).Value = Cells(j, 1).text
        i = i + 1
        Cells(i, 6).Value = "Group Counter"
        Cells(i, 7).Value = "Operation"
        Cells(i, 8).Value = "Workcentre"
        Cells(i, 9).Value = "Short Text"
        Cells(i, 10).Value = "Packages"
        i = i + 1
        k = 0
        
        Do While Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text, 1) <> "_"
            If k = 30 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                k = 0
            End If
            If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text = Cells(j, 4).text Then
                Cells(i, 6).Formula = "=""" & mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text & """"
                Cells(i, 7).Formula = "=""" & mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-VORNR[3," & CStr(k) & "]").text & """"
                Cells(i, 9).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-LTXA1[5," & CStr(k) & "]").text
                Cells(i, 10).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PACKAGES[6," & CStr(k) & "]").text
                i = i + 1
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    
    '"/nsq01":DC
    
    '"/nsq01":DE
    j = 1
    i = 0
    Do While IsEmpty(Cells(j, 3)) = False
        
        mysap.SendCommand ("/nsq01")
        mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DE"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtPACKAGE-LOW").text = Cells(j, 2).Value
        mysap.FindById("wnd[0]/usr/ctxtTASKLIST").text = Cells(j, 1).Value
        mysap.FindById("wnd[0]/usr/btn%_PACKAGE_%_APP_%-VALU_PUSH").press
        mysap.FindById("wnd[1]/tbar[0]/btn[16]").press
        mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 2).Value
        
        Do While Cells(j, 3).Value = Cells(j + 1, 3).Value
            j = j + 1
            mysap.FindById("wnd[1]/tbar[0]/btn[13]").press
            mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Cells(j, 2).Value
        Loop
        
        mysap.FindById("wnd[1]").SendVKey 8
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        i = i + 1
        i = i + 1
        i = i + 1
        k = 0
        
        Do While Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[1," & CStr(k) & "]").text, 1) <> "_"
            If k = 30 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                k = 0
            End If
            If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[1," & CStr(k) & "]").text = Cells(j, 4).text Then
                Cells(i, 8).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-ARBPL[4," & CStr(k) & "]").text
                i = i + 1
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    
    '"/nsq01":DE
    Cells.Columns.AutoFit
    If IsEmpty(Cells(1, 3)) Then
        MsgBox ("No Package has been found. The program will now end.")
    Else
        MsgBox "Done."
    End If
End Sub

Sub SAP_BulkChecking_From_FRS()
    On Error Resume Next
    
    Dim item As Range
    Dim FRS, SPMO As String
    Dim k, j, i As Integer
    
    Set mysap = getSession(, "sq01", True)
    
    '"/nsq01":DB
    
    mysap.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select
    mysap.FindById("wnd[1]/usr/radRAD1").Select
    mysap.FindById("wnd[1]/tbar[0]/btn[2]").press
    mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
    mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = 4
    mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "4"
    mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
    mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DB"
    mysap.FindById("wnd[0]").SendVKey 8
    
    For Each item In Selection
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtSP$00001-LOW").text = "*" & item.text & "*"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        Cells(item.row, item.Column + 1) = "Package Found:"
        
        If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0,0]").text = "" Then
            Cells(item.row, item.Column + 2) = "NONE"
        Else
            Cells(item.row, item.Column + 2) = "YES"
            
            k = 0
            j = 0
            Do While mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0," & CStr(k) & "]").text <> "____________"
                        
                If k = 30 Then
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    k = 0
                End If
                
                If Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0," & CStr(k) & "]").text, 2) = "H0" Then
                    Cells(item.row, item.Column + 2 + (j * 3)).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-STRAT[0," & CStr(k) & "]").text
                    Cells(item.row, item.Column + 3 + (j * 3)).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-KZYK1[1," & CStr(k) & "]").text
                    Cells(item.row, item.Column + 4 + (j * 3)).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDB============TVIEW100/txt%%G00-T351X-KTEX1[2," & CStr(k) & "]").text
                    j = j + 1
                End If
                k = k + 1
            Loop
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
    Next
    
    '"/nsq01":DB

    'new code added on 1 Feb 2016

'    mysap.sendcommand("/nzip11")
'    mysap.findById("wnd[0]").sendVKey 4
'    mysap.findById("wnd[1]/tbar[0]/btn[17]").press
'    mysap.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").Text = "H01*"
'    mysap.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").Text = "*723231*"
'    mysap.findById("wnd[1]").sendVKey 0
'
'    i = 3
'    Do While mysap.findById("wnd[1]/usr/lbl[1," & i & "]").Text Is Not Null
'        MsgBox mysap.findById("wnd[1]/usr/lbl[1," & i & "]").Text
'        Cells(1, i + 1) = mysap.findById("wnd[1]/usr/lbl[1," & i & "]").Text
'        i = i + 1
'    Loop
    
    'new code added on 1 Feb 2016
    
    Cells.Columns.AutoFit
    
    MsgBox "Done."
End Sub

'DEB
Sub SAP_ia06_UnSelectOps()  'not updated, won't accept array other than "selection" range.
    
    Dim item As Range
    Dim PageRow, AbsRow As Integer
        
    Set mysap = getSession()

    PageRow = 0
    AbsRow = 0
    
    If MsgBox("Do you want to select all the operations first?", vbYesNo) = vbYes Then mysap.FindById("wnd[0]/tbar[1]/btn[34]").press

    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    op = unique_list(Selection)
    
    For Each i In op
        Do While 1 = 1
            If PageRow = 23 Then
                PageRow = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            End If
            
            If mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").GetCell(PageRow, 0).text = i Then
                mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").GetAbsoluteRow(AbsRow).Selected = False
                Exit Do
            End If
            PageRow = PageRow + 1
            AbsRow = AbsRow + 1
        Loop
    Next
End Sub

Sub SAP_zl07_and_zl07TECO_SelectOps(ByRef mysap As Variant, oplist)
    'On Error Resume Next
    
    Dim item As Range
    Dim PageRow As Integer
    Dim op() As String

    PageRow = 0
    
    mysap.FindById("wnd[0]/mbar/menu[1]/menu[2]/menu[2]").Select
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    op = unique_list(oplist)
    
    For Each i In op
        Do While 1 = 1
            If PageRow = 23 Then
                PageRow = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            End If
            
            If mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").GetCell(PageRow, 0).text = i Then
                mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010").Rows(PageRow).Selected = True
                Exit Do
            End If
            PageRow = PageRow + 1
        Loop
    Next
End Sub

Sub SAP_zl07_PendingTV()

    Dim Op1, Op2, TSR, TV As String

    TSR = InputBox("TSR NUMBER and ISSUE (e.g. TSR103691 Issue 3)")
    TV = InputBox("TV in pending (e.g. TV157945)")
    Op1 = InputBox("1st operation (Pending TV), in 4 digits.")
    Op2 = InputBox("2nd operation (Awaiting Route Card), in 4 digits.")
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ' Don't ever do the followings unless you want it to be inefficient and a laughing stock. Down right ugly
    ' Better: Use array to store the strings and paste them
    ' Best: One line killer
    
    If MsgBox("Transformed Format?", vbYesNo) = vbYes Then
    
        Cells(1, 1).Value = "=""" & Op1 & """"
        Cells(1, 2).Value = "HKKSQES"
        Cells(1, 3).Value = "*"
        Cells(1, 4).Value = "Pending " & TV
        Cells(2, 3).Value = "*"
        Cells(3, 3).Value = "*"
        Cells(3, 4).Value = TSR
        Cells(4, 3).Value = "*"
        Cells(5, 3).Value = "*"
        Cells(5, 4).Value = "TS Clearance:___________________________________________________"
        
        Cells(1, 6).Value = "=""" & Op2 & """"
        Cells(1, 7).Value = "HKYPEPCS"
        Cells(1, 8).Value = "*"
        Cells(1, 9).Value = "Awaiting Route Card"
        Cells(2, 8).Value = "*"
        Cells(2, 9).Value = "."
        Cells(3, 8).Value = "*"
        Cells(3, 9).Value = "."
        
    Else
    
        Cells(1, 1).Value = "=""" & Op1 & """"
        Cells(1, 2).Value = "HKKSQES"
        Cells(1, 3).Value = "*"
        Cells(1, 4).Value = "Pending " & TV
        Cells(2, 3).Value = "*"
        Cells(3, 3).Value = "*"
        Cells(3, 4).Value = "TS Clearance:___________________________________________________"
        Cells(4, 3).Value = "*"
        Cells(5, 3).Value = "*"
        Cells(5, 4).Value = "Ref. " & TSR
        
        Cells(1, 6).Value = "=""" & Op2 & """"
        Cells(1, 7).Value = "HKYPEPCS"
        Cells(1, 8).Value = "*"
        Cells(1, 9).Value = "Awaiting Route Card"
        Cells(2, 8).Value = "*"
        Cells(2, 9).Value = "."
        Cells(3, 8).Value = "*"
        Cells(3, 9).Value = "."
        
    End If
    
    Set mysap = getSession()
    SAP_zl07_longtext_paster mysap
    
    If MsgBox("Do you want to print the order now?", vbYesNo) = vbYes Then
    
        mysap.FindById("wnd[0]/mbar/menu[0]/menu[6]/menu[0]").Select
        mysap.FindById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS/chkWWORKPAPER-TDIMMED[6,1]").Selected = True
        mysap.FindById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS/chkWWORKPAPER-TDDELETE[7,1]").Selected = True
        mysap.FindById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS/ctxtWWORKPAPER-TDDEST[2,1]").text = "H224"
        mysap.FindById("wnd[1]/tbar[0]/btn[8]").press
    End If
End Sub

Sub SAP_ia06_PRTAdder()
    On Error Resume Next
    
    Dim Doc, OpNo As String
    Dim r As Integer
    r = 1
    Doc = InputBox("Please Input the PRT attachment you want to add.")
    Sheets.Add After:=Sheets(Sheets.Count)
    Cells(r, 1).Value = Doc
    If MsgBox("Do you want to add more?", vbYesNo) = vbYes Then
        r = r + 1
        Doc = InputBox("Please Input the PRT attachment you want to add.")
        Cells(r, 1).Value = Doc
    End If
    
    Set mysap = getSession()

    OpNo = "0000"
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_FHM").press
    Do While OpNo <> mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        r = 1
        mysap.FindById("wnd[1]").Close
        
        Do While IsEmpty(Cells(r, 1)) = False
            If Len(Cells(r, 1).Value) = 11 Then
                mysap.FindById("wnd[0]/mbar/menu[1]/menu[0]/menu[2]").Select
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-DOKNR").text = Cells(r, 1).Value
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-DOKAR").text = "DAT"
                mysap.FindById("wnd[1]/usr/txtPLFHD-DOKTL").text = "000"
                mysap.FindById("wnd[1]/usr/txtPLFHD-DOKVR").text = "01"
            Else
                mysap.FindById("wnd[0]/mbar/menu[1]/menu[0]/menu[0]").Select
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-MATNR").text = Cells(r, 1).Value
                mysap.FindById("wnd[1]/usr/ctxtPLFHD-FHWRK").text = "HK01"
            End If
            mysap.FindById("wnd[1]/usr/ctxtPLFHD-MGEINH").text = "NPT"
            mysap.FindById("wnd[1]/usr/ctxtPLFHD-STEUF").text = "1"
            mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
            r = r + 1
        Loop
        
        OpNo = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
        
    Loop
    mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    MsgBox "Done."
End Sub

Sub SAP_ia06_RefAdder()
    On Error Resume Next

    Dim ref, Prefx As String
    Dim i, j, page, LoopCount, PageRow, AbsRow As Integer

    ref = InputBox("Please Input the Reference you want to add.")
    If MsgBox("Command Line?", vbYesNo) = vbYes Then
        Prefx = "/:"
    Else
        Prefx = "*"
    End If

    Set mysap = getSession()
    
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press

    LoopCount = 0
    PageRow = 0
    AbsRow = 0

    Do While (mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/ctxtPLPOD-ARBPL[2," & PageRow & "]").text <> "" And mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txtPLPOD-LTXA1[5," & PageRow & "]").text <> "")

        If mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").GetAbsoluteRow(AbsRow).Selected = True Then
            LoopCount = LoopCount + 1
        End If
        If PageRow = 22 Then
            PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
        Else
            PageRow = PageRow + 1
        End If
        AbsRow = AbsRow + 1
    Loop

    Do While LoopCount <> 0
        i = 0
        j = 0
        mysap.FindById("wnd[0]/tbar[1]/btn[16]").press
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select

        Do While 1 = 1
            If i = 29 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                i = 0
            End If

            If mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & i + 1 & "]").text = "" And mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i + 1 & "]").text = "" Then
                Exit Do
            Else
            i = i + 1
            j = j + 1
            End If
        Loop

        mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & i & "]").SetFocus
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        
        page = j / 28
        page = Int(page)
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While page <> 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            page = page - 1
            j = j - 29
        Loop
        
        mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & j + 1 & "]").text = Prefx
        mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & j + 1 & "]").text = ref
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        mysap.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
        LoopCount = LoopCount - 1
    Loop

    MsgBox "Successfully added."

End Sub

'DEB
Sub SAP_DetailOps()
    On Error Resume Next
    
    Dim Plan, Package, WarningTxt As String
    Dim k, j, i As Integer

    i = 1
    
    Plan = InputBox("Please enter the Plan you wish to extract.")
    
    Plan = UCase(Plan)
    If Not IsPlan(Plan) Then
        escmsg
        Exit Sub
    End If
    
    Cells(i, 1).Value = "Details of"
    Cells(i, 2).Value = Plan
    Package = "*"
    
    Set mysap = getSession()
    
    '"/nip03"

    mysap.SendCommand ("/nip03")
    mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Plan & "/1"
    mysap.FindById("wnd[0]").SendVKey 0
    Cells(1, 3).Value = "Active Group Counter:"
    Cells(1, 4).Value = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text

    '"/nip03"
    
    '"/nsq01":DC
    
    j = 1
    i = 0
    Do While IsEmpty(Cells(j, 3)) = False
    
        mysap.SendCommand ("/nsq01")
        mysap.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select
        mysap.FindById("wnd[1]/usr/radRAD1").Select
        mysap.FindById("wnd[1]/tbar[0]/btn[2]").press
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
        mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = 4
        mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "4"
        mysap.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
        mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DC"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtPACKAGE-LOW").text = Package
        mysap.FindById("wnd[0]/usr/ctxtTASKLIST").text = Plan
        mysap.FindById("wnd[0]/usr/btn%_PACKAGE_%_APP_%-VALU_PUSH").press
        mysap.FindById("wnd[1]/tbar[0]/btn[16]").press
        mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Package
        
'        Do While Cells(j, 3).Value = Cells(j + 1, 3).Value
'            j = j + 1
'            mysap.FindById("wnd[1]/tbar[0]/btn[13]").press
'            mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").Text = Cells(j, 2).Value
'        Loop
        
        mysap.FindById("wnd[1]").SendVKey 8
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        i = i + 1
        Cells(i, 6).Value = Plan
        i = i + 1
        Cells(i, 6).Value = "GrpCtr"
        Cells(i, 7).Value = "Operation"
        Cells(i, 8).Value = "Workcentre"
        Cells(i, 9).Value = "Short Text"
        Cells(i, 10).Value = "Packages"
        i = i + 1
        k = 0
        
        Do While Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text, 1) <> "_"
            If k = 30 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                k = 0
            End If
            If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text = Cells(j, 4).text Then
                Cells(i, 6).Formula = "=""" & mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[2," & CStr(k) & "]").text & """"
                Cells(i, 7).Formula = "=""" & mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-VORNR[3," & CStr(k) & "]").text & """"
                Cells(i, 9).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-LTXA1[5," & CStr(k) & "]").text
                Cells(i, 10).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDC============TVIEW100/txt%%G00-ZRRSTRUCTA-PACKAGES[6," & CStr(k) & "]").text
                i = i + 1
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    
    '"/nsq01":DC
    
    '"/nsq01":DE
    j = 1
    i = 0
    Do While IsEmpty(Cells(j, 3)) = False
        
        mysap.SendCommand ("/nsq01")
        mysap.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "DE"
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/usr/ctxtPLANT-LOW").text = "hk01"
        mysap.FindById("wnd[0]/usr/txtPACKAGE-LOW").text = Package
        mysap.FindById("wnd[0]/usr/ctxtTASKLIST").text = Plan
        mysap.FindById("wnd[0]/usr/btn%_PACKAGE_%_APP_%-VALU_PUSH").press
        mysap.FindById("wnd[1]/tbar[0]/btn[16]").press
        mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").text = Package
        
'        Do While Cells(j, 3).Value = Cells(j + 1, 3).Value
'            j = j + 1
'            mysap.FindById("wnd[1]/tbar[0]/btn[13]").press
'            mysap.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,0]").Text = Cells(j, 2).Value
'        Loop
        
        mysap.FindById("wnd[1]").SendVKey 8
        mysap.FindById("wnd[0]").SendVKey 8
        mysap.FindById("wnd[0]/tbar[1]/btn[6]").press
        
        i = i + 3
        k = 0
        
        Do While Left(mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[1," & CStr(k) & "]").text, 1) <> "_"
            If k = 30 Then
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                k = 0
            End If
            If mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-PLNAL[1," & CStr(k) & "]").text = Cells(j, 4).text Then
                Cells(i, 8).Value = mysap.FindById("wnd[0]/usr/tblAQTGSUPPLY_CHAINDE============TVIEW100/txt%%G00-ZRRSTRUCTA-ARBPL[4," & CStr(k) & "]").text
                i = i + 1
            End If
            k = k + 1
        Loop
        j = j + 1
    Loop
    
    '"/nsq01":DE
    
    '"/nzip11"
    
    mysap.SendCommand ("/nzip11")
    mysap.FindById("wnd[0]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = Plan
    mysap.FindById("wnd[0]").SendVKey 0
    Cells(1, 2).Value = mysap.FindById("wnd[0]/usr/txtZV_T351-KTEXT").text
    mysap.FindById("wnd[0]/shellcont/shell").selectItem "02", "Column1"
    mysap.FindById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem "02", "Column1"
    mysap.FindById("wnd[0]/shellcont/shell").doubleClickItem "02", "Column1"
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    i = 2
    j = 0
    Do While Left(mysap.FindById("wnd[0]/usr/tblSAPLZ_IP11IP12_TMTCTRL_ZV_T351P/txtZV_T351P-ZAEHL[0," & j & "]").text, 1) <> "_"
    
        If j = 23 Then
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            j = 0
        End If
        
        Cells(i, 1).Value = mysap.FindById("wnd[0]/usr/tblSAPLZ_IP11IP12_TMTCTRL_ZV_T351P/txtZV_T351P-KZYK1[3," & j & "]").text
        Cells(i, 2).Value = mysap.FindById("wnd[0]/usr/tblSAPLZ_IP11IP12_TMTCTRL_ZV_T351P/txtZV_T351P-KTEX1[1," & j & "]").text
        
        j = j + 1
        i = i + 1
    Loop
    
    '"/nzip11"
    
    i = 2
    Do While IsEmpty(Cells(i, 1)) = False
        Cells(2, i + 9).Value = Cells(i, 1).Value
        i = i + 1
    Loop
    
    i = i - 1
    
    WarningTxt = ""
    
    Do While i > 1
        Package = Cells(2, i + 9).Value
        j = 3
        Do While IsEmpty(Cells(j, 10)) = False
            
            If InStr(1, Cells(j, 10).text, Package) <> 0 Then
                Cells(j, i + 9).Value = "X"
            End If
            j = j + 1
        Loop
        i = i - 1
    Loop

    j = 3
    Do While IsEmpty(Cells(j, 10)) = False
    
        If Len(Cells(j, 10).text) >= 80 Then
            WarningTxt = WarningTxt & Cells(j, 7).text & ", "
        End If
        j = j + 1
    Loop
    
    WarningTxt = Left(WarningTxt, Len(WarningTxt) - 2)
    
    Cells.Columns.AutoFit
    If IsEmpty(Cells(1, 3)) Then
        MsgBox ("No Package has been found. The program will now end.")
    Else
        MsgBox "Done."
        If WarningTxt <> "" Then
            MsgBox "WARNING: There are operation(s) that exceeded max package limit of SQ01!" & vbCrLf & "Please check operation: " & WarningTxt
        End If
    End If
    
End Sub

Sub SAP_ia06_GoToPlan()
    Dim sp As String
    If Selection.Rows.Count <> 1 Or Not IsPlan(Selection(1).text) Then
        escmsg
        Exit Sub
    Else
        Set mysap = getSession()
        If Selection.Columns.Count > 1 Then
            sp = Selection(2).text
        Else
            sp = ""
        End If
        GotoPlan mysap, Selection(1).text, sp
    End If
End Sub

Sub GotoPlan(ByRef mysap As Variant, Plan As String, Optional SpecificName As String = "")
    On Error GoTo Escape
    mysap.SendCommand ("/nia06")
    mysap.FindById("wnd[0]/usr/ctxtRC271-PLNNR").text = Replace(Plan, "/1", "")
    mysap.FindById("wnd[0]").SendVKey 0
    
    If mysap.Children.Count > 1 Then GoTo Escape
    
    i = 0
    Do While 1 = 1
        shttxt = UCase(mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200/txtPLKOD-KTEXT[1," & i & "]").text)
        
        If SpecificName = "" Then
            Logic1 = InStr(1, shttxt, "INVALID") = 0
            Logic2 = InStr(1, shttxt, "DUMMY") = 0
            Logic3 = InStr(1, shttxt, "VOID") = 0
            Logic4 = InStr(1, shttxt, "DO NOT USE") = 0
            Logic = Logic1 And Logic2 And Logic3 And Logic4
        Else
            Logic = (UCase(SpecificName) = shttxt)
        End If
        
        If Logic Then
            mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200").GetAbsoluteRow(i).Selected = True
            mysap.FindById("wnd[0]/mbar/menu[2]/menu[2]").Select
            Exit Sub
        End If
        i = i + 1
    Loop
    
Escape:
    Exit Sub
    
End Sub
Sub sap_element_analyzer()

    Sheets.Add After:=Sheets(Sheets.Count)
    Dim a As Object
    Set a = ActiveSheet
    Set mysap = getSession()
    Set wnd = mysap.FindById("wnd[0]")
    
    a.Cells(1, 1) = "Window Name"
    a.Cells(1, 2) = "Window Text"
    a.Cells(1, 3) = "Window Type"
    a.Cells(1, 4) = "Item Name"
    a.Cells(1, 5) = "Item Text"
    a.Cells(1, 6) = "Item Type"
    a.Cells(1, 7) = "Object Name"
    a.Cells(1, 8) = "Object Text"
    a.Cells(1, 9) = "Object Type"
    
    i = 2
    a.Cells(i, 1) = wnd.Name
    a.Cells(i, 2) = wnd.text
    a.Cells(i, 3) = wnd.Type
    For Each item In wnd.Children
        a.Cells(i, 4) = item.Name
        a.Cells(i, 5) = item.text
        a.Cells(i, 6) = item.Type
        For Each obj In item.Children
            a.Cells(i, 7) = obj.Name
            a.Cells(i, 8) = obj.text
            a.Cells(i, 9) = obj.Type
            i = i + 1
        Next obj
    Next item
    
    a.Columns.AutoFit
    a.Range("1:1").Font.Bold = True
    With a.Range("A1:C1").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With a.Range("D1:F1").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    With a.Range("G1:I1").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    MsgBox ("Done!")
End Sub

Sub sap_element_analyzer2()
    
    Set mysap0 = getSession(0)
    Set mysap1 = getSession(1)
    
    MsgBox mysap0.Children(0).text
    MsgBox mysap1.Children(0).text
    
End Sub

Sub SAP_zl07_AddRef()
    
    Dim Ops() As String
    ReDim Ops(0)
    
    ref = InputBox("Please enter the reference line you want to add.")
    
    PageRow = 0
    AbsRow = 0

    Set mysap = getSession(, "iw32", False)
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    Set sapPage = getZL07(mysap)
    Do While (sapPage.GetCell(PageRow, 2).text <> "" And sapPage.GetCell(PageRow, 7).text <> "")

        If sapPage.GetAbsoluteRow(AbsRow).Selected = True Then
            ReDim Preserve Ops(UBound(Ops) + 1)
            Ops(UBound(Ops) - 1) = sapPage.GetCell(PageRow, 0).text
        End If
        If PageRow = 22 Then
            PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Set sapPage = getZL07(mysap)
        Else
            PageRow = PageRow + 1
        End If
        AbsRow = AbsRow + 1
    Loop

    ReDim Preserve Ops(UBound(Ops) - 1)

    For Each op In Ops
        i = 1
        r = 1
        'Find the matching OP
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        Set sapPage = getZL07(mysap)
        PageRow = 0
        Do While sapPage.GetCell(PageRow, 0).text <> op
            If PageRow = 22 Then
                PageRow = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set sapPage = getZL07(mysap)
            Else
                PageRow = PageRow + 1
            End If
        Loop
        
        sapPage.GetCell(PageRow, 8).SetFocus
        sapPage.GetCell(PageRow, 8).press
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
        Set sapPage = getEditor(mysap)
        'Locating the line

        
        Do While 1 = 1
            ''''''''''''''''''''''''''''''''''' LOGIC IS A BITCH '''''''''''''''''''''''''''''''''''
            txtline = UCase(sapPage.GetCell(i, 2).text)
            StartsWithB2 = (Left(txtline, 4) = "<B2>")
            StartsWithFRS = (Left(txtline, 3) = "FRS")
            StartsWithSB = (Left(txtline, 2) = "SB")
            StartsWithTV = (Left(txtline, 2) = "TV")
            StartsWithRef = (Left(txtline, 3) = "REF")
            StartsWithRefer = (Left(txtline, 5) = "REFER")
            ContainsRef = (InStr(1, txtline, "REF") <> 0)
            ContainsRefer = (InStr(1, txtline, "REFER") <> 0)
            StartsWithSUBTASK = (Left(txtline, 7) = "SUBTASK")
            ContainsDot = (InStr(1, txtline, ".") <> 0)
            ContainsComma = (InStr(1, txtline, ",") <> 0)
            ContainsSUBTASK = (InStr(1, txtline, "SUBTASK") <> 0)
            ContainsSAP = (InStr(1, txtline, "SAP") <> 0)
            ContainsPARA = (InStr(1, txtline, "PARA") <> 0)
            
            Logic1 = (StartsWithRef And Not StartsWithRefer And Not ContainsSAP)
            Logic2 = (StartsWithB2 And ContainsRef And Not ContainsRefer)
            
            T_Logic1 = (StartsWithFRS And Not ContainsDot And ContainsSUBTASK)
            T_Logic2 = (StartsWithSB And ((Not ContainsDot And Not ContainsPARA) Or (ContainsDot And ContainsPARA)))
            T_Logic3 = (StartsWithB2 And Not ContainsRef)
            T_Logic4 = (StartsWithSUBTASK And Not ContainsDot And Not ContainsComma)
            T_Logic5 = (StartsWithTV And Not ContainsComma)
            
            Jackpot = Logic1 Or Logic2
            Jackpot_T = T_Logic1 Or T_Logic2 Or T_Logic3 Or T_Logic4 Or T_Logic5
            ''''''''''''''''''''''''''''''''''' LOGIC IS A BITCH '''''''''''''''''''''''''''''''''''
            
            If Jackpot Or Jackpot_T Then
                If Jackpot Then
                    transformed = False
                Else
                    transformed = True
                End If
                i = i - 1
                r = r - 1
                Exit Do
            Else
                If i <> 30 Then
                    i = i + 1
                Else
                    i = 2
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    Set sapPage = getEditor(mysap)
                End If
            End If
            r = r + 1
        Loop
    
        'Adding the line
        mysap.FindById("wnd[0]/mbar/menu[1]/menu[7]").Select
        mysap.FindById("wnd[1]/usr/txtRSTXT-TXLINENR").text = r
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
        mysap.FindById("wnd[0]").SendVKey 0
        Set sapPage = getEditor(mysap)
        If transformed Then
            sapPage.GetCell(2, 2).text = ref
        Else
            sapPage.GetCell(2, 2).text = "Ref. " & ref
        End If
        
        'Previous Page
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    Next
    
    MsgBox ("Done.")
End Sub

Sub test()
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    Set SAPConnection = SAPApplication.Children(0)
    Set mysap = SAPConnection.Children(0)
    EOS = False
    i = 1
    
    With mysap
        .FindById("wnd[0]/tbar[0]/btn[80]").press
        Do While EOS = 0
            For Each item In .FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/").FindAllByName("AFVGD-VORNR", GuiTextField)
                If IsNumeric(item.text) Then
                    Cells(i, 1) = item.text
                    If item.text = "9999" Then
                        EOS = True
                        Exit For
                    Else
                        i = i + 1
                    End If
                End If
            Next
            .FindById("wnd[0]/tbar[0]/btn[82]").press
        Loop
    End With
End Sub

Sub SAP_zl07_longtext_paster(ByRef mysap As Variant)

    On Error Resume Next
    
    Dim i, PageRow, page, r, c, AbsRow, ExistingRow, OpCount, RowPerOp, LoopCount As Integer
    Dim ShortTxt As String
    Dim OpExists As Boolean
    
    c = 1
    RowPerOp = 5
    OpCount = Application.WorksheetFunction.RoundUp((ActiveSheet.UsedRange.Columns.Count) / RowPerOp, 0)
    
    LoopCount = OpCount
    
    Do While LoopCount <> 0
        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        Set sapTbl = getZL07(mysap)
        i = 0
        r = 1
        Do While 1 = 1
            If i = 23 Then
                i = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set sapTbl = getZL07(mysap)
            Else
                If (sapTbl.GetCell(i, 0).text <> "" And sapTbl.GetCell(i, 7).text <> "") Then
                    '' Added July 2016: Check if Op exists
                    If sapTbl.GetCell(i, 0).text = Cells(r, c).Value Then
                        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
                        Set sapTbl = getZL07(mysap)
                        OpExists = True
                        Exit Do
                    End If
                    '' Added July 2016: Check if Op exists
                    i = i + 1
                Else
                    sapTbl.GetCell(i, 0).text = Cells(r, c).Value
                    sapTbl.GetCell(i, 2).text = Cells(r, (c + 1)).Value
                    sapTbl.GetCell(i, 11).text = "1"
                    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
                    Set sapTbl = getZL07(mysap)
                    OpExists = False
                    Exit Do
                End If
            End If
        Loop
        
        i = 0
        Do While sapTbl.GetCell(i, 0).text <> ""
            If sapTbl.GetCell(i, 0).text = Cells(r, c).text Then
                sapTbl.GetCell(i, 8).SetFocus
                sapTbl.GetCell(i, 8).press
                Exit Do
            Else
                If i <> 22 Then
                    i = i + 1
                Else
                    i = 0
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    Set sapTbl = getZL07(mysap)
                End If
            End If
        Loop
        
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
        
        PageRow = 1
        AbsRow = 1
        
        ''' Amended July 2016 OpExists, go match the rows
        If OpExists Then
            Set sapTbl = getEditor(mysap)
            Do While (sapTbl.GetCell(2, 0).text <> "") Or (sapTbl.GetCell(2, 2).text <> "")
                For j = 2 To 30 Step 1
                    sapTbl.GetCell(j, 0).text = ""
                    sapTbl.GetCell(j, 2).text = ""
                Next
                mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
                Set sapTbl = getEditor(mysap)
            Loop
        End If
        ''' Amended July 2016 OpExists, go match the rows

        Do While (IsEmpty(Cells(AbsRow, c + 2)) = False Or IsEmpty(Cells(AbsRow, c + 3)) = False)
            If AbsRow <> 1 Then
                mysap.FindById("wnd[0]").SendVKey 0
            End If
            AbsRow = AbsRow + 1
        Loop

        mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        Set sapTbl = getEditor(mysap)
        
        PageRow = 1
        AbsRow = 1

        Do While (IsEmpty(Cells(AbsRow, c + 2)) = False Or IsEmpty(Cells(AbsRow, c + 3)) = False)
            If PageRow = 30 Then
                PageRow = 1
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set sapTbl = getEditor(mysap)
            End If
            sapTbl.GetCell(PageRow, 0).text = Cells(AbsRow, c + 2).Value
            sapTbl.GetCell(PageRow, 2).text = Cells(AbsRow, c + 3).Value
            PageRow = PageRow + 1
            AbsRow = AbsRow + 1
        Loop
        
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        
        Do While IsEmpty(Cells(r, c + 4)) = False
            mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/btnBTN_FHUE").press
            ''' Amended July 2016 OpExists, delete all PRTs to start anew
            If OpExists Then
                mysap.FindById("wnd[1]").Close
                mysap.FindById("wnd[0]/mbar/menu[1]/menu[2]/menu[0]").Select
                mysap.FindById("wnd[0]/tbar[1]/btn[14]").press
                mysap.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                mysap.FindById("wnd[0]/mbar/menu[1]/menu[0]/menu[0]").Select
                OpExists = False
            End If
            ''' Amended July 2016 OpExists, delete all PRTs to start anew
            mysap.FindById("wnd[1]/tbar[0]/btn[8]").press
            If InStr(1, Cells(r, c + 4).Value, "DAT") <> 0 Then
                mysap.FindById("wnd[1]/usr/ctxtAFFHD-DOKNR").text = Mid(Cells(r, c + 4).Value, 1, (InStr(1, Cells(r, c + 4).Value, " ") - 1))
                mysap.FindById("wnd[1]/usr/ctxtAFFHD-DOKAR").text = "DAT"
                mysap.FindById("wnd[1]/usr/txtAFFHD-DOKTL").text = "000"
                mysap.FindById("wnd[1]/usr/txtAFFHD-DOKVR").text = "01"
            ElseIf InStr(1, Cells(r, c + 4).Value, "REC") <> 0 Then
                mysap.FindById("wnd[1]/usr/ctxtAFFHD-DOKNR").text = Mid(Cells(r, c + 4).Value, 1, (InStr(1, Cells(r, c + 4).Value, " ") - 1))
                mysap.FindById("wnd[1]/usr/ctxtAFFHD-DOKAR").text = "REC"
                mysap.FindById("wnd[1]/usr/txtAFFHD-DOKTL").text = "000"
                mysap.FindById("wnd[1]/usr/txtAFFHD-DOKVR").text = "01"
            Else
                mysap.FindById("wnd[1]/tbar[0]/btn[6]").press
                mysap.FindById("wnd[1]/usr/ctxtAFFHD-MATNR").text = Mid(Cells(r, c + 4).Value, 1, (InStr(1, Cells(r, c + 4).Value, " ") - 1))
                mysap.FindById("wnd[1]/usr/ctxtAFFHD-FHWRK").text = "HK01"
            End If
            mysap.FindById("wnd[1]/usr/ctxtAFFHD-MGEINH").text = Mid(Cells(r, c + 4).Value, InStr(1, Cells(r, c + 4).Value, "+") + 1)
            
            mysap.FindById("wnd[1]/usr/ctxtAFFHD-STEUF").text = "1"
            If IsEmpty(Cells(r + 1, c + 4)) Then
                mysap.FindById("wnd[1]/tbar[0]/btn[29]").press
                Exit Do
            Else
                mysap.FindById("wnd[1]/tbar[0]/btn[20]").press
                r = r + 1
            End If
        Loop
        
        If IsEmpty(Cells(1, c + 4)) = False Then
            mysap.FindById("wnd[1]").Close
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
        
        c = c + RowPerOp
        LoopCount = LoopCount - 1
        
    Loop
    MsgBox "Successfully pasted."
End Sub

Sub oneoff_ip03()

    Plan = InputBox("please enter the plan (with ""/1"") you want to serach.")
    
    Set mysap = getSession(, "ip03")
    
    mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Plan
    mysap.FindById("wnd[0]").SendVKey 0
    mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15").Select
    Cells(1, 1) = "Call Number"
    Cells(1, 2) = "PlanDate"
    Cells(1, 3) = "Call date"
    Cells(1, 4) = "Completion date"
    Cells(1, 5) = "Due Package"
    Cells(1, 6) = "Scheduling Type / Status"
    Cells(1, 7) = "Released by"
    i = 0
    j = 2
    call_num = -1
    Do While Left(mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-ABNUM[0," & CStr(i) & "]").text, 1) <> "_"
        For i = 0 To 13 Step 1
            If call_num < CInt(Replace(mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-ABNUM[0," & CStr(i) & "]").text, ".", "")) Then
                Cells(j, 1) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-ABNUM[0," & CStr(i) & "]").text
                Cells(j, 2) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-MANDA[1," & CStr(i) & "]").text
                Cells(j, 3) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-ABRUD[2," & CStr(i) & "]").text
                Cells(j, 4) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-LRMDT[3," & CStr(i) & "]").text
                Cells(j, 5) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtTERMFELD3[4," & CStr(i) & "]").text
                Cells(j, 6) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtTERMFELD2[5," & CStr(i) & "]").text
                Cells(j, 7) = mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtAMHIS-ABRNA[6," & CStr(i) & "]").text
                call_num = CInt(Replace(mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-ABNUM[0," & CStr(i) & "]").text, 3))
                j = j + 1
            Else
                end_page = True
            End If
        Next
        If end_page Then
            Exit Do
        Else
            i = 0
            mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY2:SAPLIWP3:8025/subSUBSCREEN_SCHED_CALLS_ITEM:SAPLIWP3:0124/tblSAPLIWP3TCTRL_0124/txtRMIPM-ABNUM[0," & CStr(i) & "]").SetFocus
            mysap.FindById("wnd[0]").SendVKey 82
        End If
    Loop
    ActiveSheet.Columns.AutoFit
    MsgBox ("Done.")
End Sub

Function IsTransaction(ByVal mysap As Object, ByVal transaction As String) As Boolean
    If UCase(mysap.Info.transaction) <> UCase(transaction) Then
        IsTransaction = False
    Else
        IsTransaction = True
    End If
End Function

Sub err_IncorrectTranscation(Optional ByVal Trans As String = "")
    If Trans <> "" Then
        MsgBox ("Incorrect SAP Transaction. The current Transaction is " & Trans & ".")
    Else
        MsgBox ("Incorrect SAP Transaction.")
    End If
End Sub

Sub SAP_longtext_paster()
    Set mysap = getSession()
    If IsTransaction(mysap, "IW32") Then
        SAP_zl07_longtext_paster mysap
    ElseIf IsTransaction(mysap, "IA06") Then
        SAP_ia06_longtext_paster mysap
    Else
        err_IncorrectTranscation mysap.Info.transaction
    End If
End Sub

Sub SAP_longtext_copier()
    Set mysap = getSession()
    If IsTransaction(mysap, "IW32") Then
        If getZL07(mysap).GetCell(0, 0).Changeable = True Then
            SAP_zl07_longtext_copier mysap
        Else
            SAP_zl07TECO_longtext_copier mysap
        End If
    ElseIf IsTransaction(mysap, "IA06") Then
        SAP_ia06_longtext_copier mysap
    Else
        err_IncorrectTranscation mysap.Info.transaction
    End If
End Sub

'DEB
Sub SAP_selectops_callable()
    SAP_selectops
End Sub

Sub SAP_selectops(Optional sp_array = "")
    Dim oplist() As Variant
    If Not IsArray(sp_array) Then   'If no specified array feeded in, the function will choose "Selection" as the array
        If IsArray(Selection) Then  ' Fucking excel returns single selected cell as non-range non-array item
            oplist = Selection.Value
        Else
            ReDim oplist(0 To 0)
            oplist(0) = Selection.Value
        End If
    Else
        oplist = sp_array
    End If
    For Each Ops In oplist
        If Not IsNumeric(Ops) And Len(Ops) <> 4 Then
            escmsg
            Exit Sub
        End If
    Next
    
    Set mysap = getSession()
    
    If IsTransaction(mysap, "IW32") Or IsTransaction(mysap, "IW33") Then
        SAP_zl07_and_zl07TECO_SelectOps mysap, oplist
    ElseIf IsTransaction(mysap, "IA06") Then
        SAP_ia06_SelectOps mysap, oplist
    Else
        err_IncorrectTranscation mysap.Info.transaction
    End If
End Sub

'TV163423 Issue 2 PARA.3 (TSR105811 Issue 1)
Sub zl07_replace()

    ' Field usage of cList
    
    Dim strFind As clist
    Set strFind = New clist
    
    Dim strReplace As clist
    Set strReplace = New clist
    
    strFind.Add "TSR106269 Issue 6"
    strReplace.Add "TSR106269 Issue 8"
    strFind.Add "TV169987"
    strReplace.Add "TV171907"
    strFind.Add "PARA.16"
    strReplace.Add "PARA.14"
    
    Set mysap = getSession()
    SAP_zl07_longtext_replacer mysap, strFind, strReplace
        
End Sub
Sub SAP_zl07_longtext_replacer(ByRef mysap As Variant, ByRef strFind As Variant, ByRef strReplace As Variant)
    Dim PageRow, AbsRow As Integer
    Dim oplist As clist
    Set oplist = New clist

    PageRow = 0
    AbsRow = 0

    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    Set sapTbl = getZL07(mysap)
    Do While sapTbl.GetCell(PageRow, 2).text <> "" And sapTbl.GetCell(PageRow, 7).text <> ""
        If sapTbl.GetAbsoluteRow(AbsRow).Selected = True Then
            oplist.Add sapTbl.GetCell(PageRow, 0).text
        End If
        If PageRow = 22 Then
            PageRow = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
            Set sapTbl = getZL07(mysap)
        Else
            PageRow = PageRow + 1
        End If
        AbsRow = AbsRow + 1
    Loop

    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    PageRow = 0
    For Each op In oplist.Value
        Set sapTbl = getZL07(mysap)
        Do While sapTbl.GetCell(PageRow, 0).text <> op
            If PageRow = 22 Then
                PageRow = 0
                mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                Set sapTbl = getZL07(mysap)
            Else
                PageRow = PageRow + 1
            End If
        Loop
        
        sapTbl.GetCell(PageRow, 8).SetFocus
        sapTbl.GetCell(PageRow, 8).press
        
        mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
        
        For i = LBound(strFind.Value) To UBound(strFind.Value)
            mysap.FindById("wnd[0]/mbar/menu[1]/menu[4]").Select
            mysap.FindById("wnd[1]/usr/chkRSTXT-TXFIWORD").Selected = False
            mysap.FindById("wnd[1]/usr/chkRSTXT-TXFIUPCASE").Selected = False
            mysap.FindById("wnd[1]/usr/chkRSTXT-TXFIBACKWD").Selected = False
            mysap.FindById("wnd[1]/usr/radRSTXT-TXFITEXT").Select
            mysap.FindById("wnd[1]/usr/txtRSTXT-TXFISTRING").text = strFind.Value(i)
            mysap.FindById("wnd[1]/usr/txtRSTXT-TXRESTRING").text = strReplace.Value(i)
            mysap.FindById("wnd[1]/tbar[0]/btn[5]").press
        Next
        
        mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    Next
End Sub

Sub logoff(ByRef mysap As Variant)
    mysap.parent.CloseConnection
End Sub

Sub sap_ca72_yes_or_no()
    On Error Resume Next
    Set mysap = getSession()
    For Each item In Selection
        mysap.SendCommand ("/nca72")
        mysap.FindById("wnd[0]/tbar[1]/btn[16]").press
        mysap.FindById("wnd[1]/usr/cmbLISTBOX").key = "Document"
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
        mysap.FindById("wnd[0]/usr/ctxtRCA80-DOKNR").text = item
        mysap.FindById("wnd[0]/usr/ctxtRCA80-DOKAR").text = "DAT"
        mysap.FindById("wnd[0]/usr/txtRCA80-DOKTL").text = "000"
        mysap.FindById("wnd[0]/usr/txtRCA80-DOKVR").text = "01"
        mysap.FindById("wnd[0]/usr/subPLAN:SAPMC27V:1100/ctxtRCA80-PLANWERK").text = "HK01"
        mysap.FindById("wnd[0]/usr/subPLAN:SAPMC27V:1100/ctxtRCA80-PLNTY_V").text = "A"
        mysap.FindById("wnd[0]/tbar[1]/btn[8]").press
        Debug.Print mysap.FindById("wnd[0]/usr/ctxtRCA80-DOKNR").text
        If Err.Number <> 0 Then
            item.Offset(0, 5) = "NO"
        Else
            item.Offset(0, 5) = "YES"
        End If
        Err.Number = 0
    Next
    
End Sub
Sub sap_zl07_remove_top_PMO()
' Prob rekt, dont use this unless you wanna make havoc
' Go see zl07print_WASH_NDT for robust solution
    Set mysap = getSession(, "zl07", False)
    For i = 1 To InputBox("How many times?")
        mysap.FindById("wnd[0]/usr/chk[0,7]").Selected = True
        mysap.FindById("wnd[0]/mbar/menu[4]/menu[5]").Select
        mysap.FindById("wnd[1]/usr/btnBUTTON_1").press
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
    Next
End Sub

Sub zl07test()
    Set mysap = getSession(, "zl07", False)
    OpCount = (mysap.FindById("wnd[0]/usr").Children.Count - 54) / 30
    MsgBox OpCount
End Sub

Sub zl07test2()
    Set mysap = getSession(, "zl07", False)
    i = 0
    For Each item In mysap.FindById("wnd[0]/usr").Children
        If item.Type = "GuiCheckBox" Then
            a = item.ID
        End If
    Next
    mysap.FindById(a).Selected = True
End Sub

Sub tickpackages(ByVal mysap As Object, Package As String)
    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press
    mysap.FindById("wnd[0]/tbar[1]/btn[26]").press
    Do
        lastop = mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
        mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0,0]").text = Package
        mysap.FindById("wnd[0]").SendVKey 0
        mysap.FindById("wnd[0]/tbar[1]/btn[19]").press
    Loop While lastop <> mysap.FindById("wnd[0]/usr/txtPLPOD-VORNR").text
    mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
    mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
End Sub

