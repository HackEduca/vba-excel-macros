Attribute VB_Name = "CountTV"
'One-off code for counting TV in ia17.
'Outdated and slow.
'One of the first subs ever made!

Sub CountTV()
    On Error Resume Next
    Set mysap = getSession(, "ia17", False)
    
    With mysap
        .FindById("wnd[0]").Maximize
        For Each item In Selection
            .FindById("wnd[0]").SendVKey 71
            .FindById("wnd[1]/usr/chkSCAN_STRING-START").Selected = False
            .FindById("wnd[1]/usr/chkSCAN_STRING-RANGE").Selected = False
            '.FindById("wnd[1]/usr/txtRSYSF-STRING").Text = Cells(item.row, item.Column).Text
            .FindById("wnd[1]/usr/txtRSYSF-STRING").text = item
            .FindById("wnd[1]").SendVKey 0

            'Cells(item.row, item.Column + 1) = .FindById("wnd[2]/usr/lbl[16,0]").Text
            item.Offset(0, 1) = .FindById("wnd[2]/usr/lbl[16,0]").text
            If IsEmpty(item.Offset(0, 1)) = True Then
                item.Offset(0, 1) = 0
            End If
            .FindById("wnd[2]").Close
            .FindById("wnd[1]/tbar[0]/btn[12]").press
        Next
    End With
End Sub

Sub Runtime()
    CountTV
    ActiveWorkbook.Save
    Shutdown
End Sub
