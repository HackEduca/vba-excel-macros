Attribute VB_Name = "Operation_Odyssey"
' Late 2015 on Andrew K.'s demand
' Unfinished

Sub SAP_Extract()
'    On Error Resume Next

    Dim op, SUBTASK As String
    Dim i, j, k, l, pos As Integer
    Dim Repeated As Boolean
    
    ATA = InputBox("Please enter the TASK you wish to extract.")
    SPMO = InputBox("Please enter the Order you wish to search in.")
    
    If Len(SPMO) <> 8 Or IsNumeric(SPMO) = False Then
        MsgBox ("Invalid SPMO Selection. Please try again.")
        Exit Sub
    End If

    MsgBox "Starting SAP GUI Scripts. Please make sure your SAP Client is ready."
    
    Set mysap = getSession()
    
    mysap.FindById("wnd[0]").resizeWorkingPane 1, 1, True
    mysap.FindById("wnd[0]").resizeWorkingPane 133, 34, False
    mysap.SendCommand ("/niw32")
    mysap.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = SPMO
    mysap.FindById("wnd[0]").SendVKey 0
    mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE").Select
    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
    
    Cells(1, 1).Value = "Input SPMO:"
    Cells(2, 1).Value = "ATA to search for:"
    Cells(1, 2).Value = SPMO
    Cells(2, 2).Value = ATA
    Cells(1, 4).Value = "Operation No."
    Cells(1, 5).Value = "SUBTASK No."

    k = 2
    i = 0
    
    Do While 1 = 1
    
        If i = 22 Then
            i = 0
            mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
        End If
        
        If (Left(mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/ctxtAFVGD-ARBPL[2," & i & "]").text, 1) = "_" And Left(mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7," & i & "]").text, 1) = "_") Then
            Exit Do
        End If
        
        op = mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-VORNR[0," & i & "]").text
        
        If CInt(op) > 499 And CInt(op) < 9900 Then
            mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8," & i & "]").SetFocus
            mysap.FindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLTICON-LTOPR[8," & i & "]").press
            mysap.FindById("wnd[0]/mbar/menu[2]/menu[3]").Select
            j = 1
            
            Do While (Left(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0," & j & "]").text, 1) <> "_") And (Left(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & j & "]").text, 1) <> "_")
                pos = 1

                Do While InStr(pos, mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & j & "]").text, ATA) <> 0
                    pos = InStr(pos, mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & j & "]").text, ATA)
                    SUBTASK = Mid(mysap.FindById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2," & j & "]").text, pos, 16)
                    l = k - 1
                    Repeated = False
                    Do While Cells(l, 4).text = op
                        If Cells(l, 5).text = SUBTASK Then
                            Repeated = True
                        End If
                        l = l - 1
                    Loop
                    If Repeated = False Then
                        Cells(k, 4) = op
                        Cells(k, 5) = SUBTASK
                        k = k + 1
                    End If
                    pos = pos + 1
                Loop
                If j <> 30 Then
                    j = j + 1
                Else
                    j = 2
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                End If
            Loop
            mysap.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If
        i = i + 1
    Loop

    Cells.Columns.AutoFit
    MsgBox "Search complete. Please double-check for any unwanted SUBTASK extracted."
End Sub
