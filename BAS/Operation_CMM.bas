Attribute VB_Name = "Operation_CMM"
' Early 2016?
' lol this shit looks stupid

Sub SAP_ia06_CMMfindPackage()
    
    On Error Resume Next
    
    Set mysap = getSession()
    
    j = 8
    
    Do While IsEmpty(Cells(j, 2).Value) = False Or IsEmpty(Cells(j + 1, 2).Value) = False
            
        If Left(Cells(j, 2).Value, 1) = "H" And IsEmpty(Cells(j, 11).Value) Then
        
            mysap.SendCommand ("/nia06")     ' Finding packages using IA06 WHAT THE FUCK MATE
            
            mysap.FindById("wnd[0]/usr/ctxtRC271-PLNNR").text = Cells(j, 2).Value
            mysap.FindById("wnd[0]").SendVKey 0
            
            i = 0
            
            Do While 1 = 1

                If IsNumeric(mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200/txtPLKOD-PLNAL[0," & i & "]").text) = False Then
                    Exit Do
                ElseIf mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200/txtPLKOD-PLNAL[0," & i & "]").text = Cells(j, 3).text Then
                    mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3200/txtPLKOD-KTEXT[1," & i & "]").SetFocus
                    mysap.FindById("wnd[0]").SendVKey 2
                    Exit Do
                End If
                i = i + 1
            Loop
                    mysap.FindById("wnd[0]/tbar[0]/btn[80]").press
        
            r = 0
            AbsR = 0
            
            Do While mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txtPLPOD-VORNR[0," & r & "]").text <> ""
                If r = 23 Then
                    mysap.FindById("wnd[0]/tbar[0]/btn[82]").press
                    r = 1
                End If
                
                If mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400/txtPLPOD-VORNR[0," & r & "]").text = Cells(j, 4).text Then
                    mysap.FindById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").GetAbsoluteRow(AbsR).Selected = True
                    mysap.FindById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press
                    mysap.FindById("wnd[0]/tbar[1]/btn[26]").press
                    
                    k = 1
                    l = 0
                    Do While mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & k & "]").text <> ""
                        Cells(j, 12 + l).text = mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & k & "]").text
                        Cells(j, 13 + l).text = mysap.FindById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KTEX1[2," & k & "]").text
                        k = k + 1
                        l = l + 2
                    Loop
                End If
                r = r + 1
                AbsR = AbsR + 1
            Loop
            
            If IsEmpty(Cells(j, 12)) Then
                Cells(j, 12 + l).text = "Operation not found"
            End If
        End If
        
        j = j + 1
    Loop
    
End Sub

Sub SAP_CheckIfInvalid()
    Set mysap = getSession()
    
    j = 8
        
    Do While IsEmpty(Cells(j, 2).Value) = False Or IsEmpty(Cells(j + 1, 2).Value) = False
        If Left(Cells(j, 2).Value, 1) = "H" Then
            mysap.SendCommand ("/nip03")
            mysap.FindById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Cells(j, 2).Value & "/1"
            mysap.FindById("wnd[0]").SendVKey 0
            
            If mysap.FindById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text <> Cells(j, 3).text Then
                Cells(j, 11).Value = "INVALID"
            End If
        End If
        
        j = j + 1
    Loop
End Sub
