Attribute VB_Name = "ztext_tasklistTVcounter"
' Just for reference
' Don't be stupid enough to run this

Sub ztext_tasklistTVcounter()
    On Error Resume Next
    
    Set mysap = get_mymysap()
    
    Range("C3:C13") = Empty
    Dim i As Integer
    i = 3
    mysap.SendCommand ("/nztext_tasklists")
    
    Do While i <= 13
    
        With mysap
            .FindById("wnd[0]/usr/radRB_OPERA").Select
            .FindById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "HK01"
            .FindById("wnd[0]/usr/ctxtS_PLNTY-LOW").text = "A"
            .FindById("wnd[0]/usr/ctxtS_PLNNR-LOW").text = Range("A" & CStr(i)).Value
            .FindById("wnd[0]/usr/ctxtS_PLNNR-HIGH").text = Range("B" & CStr(i)).Value
            .FindById("wnd[0]/usr/txtP_STRNG1").text = "*"
            .FindById("wnd[0]").SendVKey 8
            .FindById("wnd[0]").SendVKey 71
            .FindById("wnd[1]/usr/chkSCAN_STRING-START").Selected = False
            .FindById("wnd[1]/usr/chkSCAN_STRING-RANGE").Selected = False
            .FindById("wnd[1]/usr/txtRSYSF-STRING").text = Range("B1").Value
            .FindById("wnd[1]").SendVKey 0

        
            Range("C" & CStr(i)).Value = .FindById("wnd[2]/usr/lbl[16,0]").text
            If IsEmpty(Range("C" & CStr(i))) = True Then
                Range("C" & CStr(i)).Value = 0
            End If
            .FindById("wnd[2]").Close
            .FindById("wnd[1]/tbar[0]/btn[12]").press
            .FindById("wnd[0]/tbar[0]/btn[3]").press
        End With
        i = i + 1
    Loop
    
End Sub

