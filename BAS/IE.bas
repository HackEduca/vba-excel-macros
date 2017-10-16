Attribute VB_Name = "IE"
' Collection of IE related scripts.
' Loosely connected. most of them are extracted from the excel file to add XWB EIPC

Sub open_RRCare()
    Dim IE As InternetExplorer
    Set IE = getIE()
    IE.Visible = True
    IE.Navigate2 "https://customers.rolls-royce.com/secure/rollsroycecare"
    Do
    Loop While IE.Busy 'Or IE.readyState <> READYSTATE_COMPLETE
    set input =Ie.Document.getelementsbytagname
'    IE.Document.getElementById("login") = strLogin
'    IE.Document.getElementById("password") = strPassword
'    IE.Document.getElementsByName("tandcaccept")(0).Click
    IE.quit
    Set IE = Nothing
End Sub

Function getIE() As Object
    Dim shellWin As ShellWindows
    Set shellWin = New ShellWindows
    
    If shellWin.Count > 0 Then
        Set getIE = shellWin.item(0)
    Else
        Set getIE = CreateObject("InternetExplorer.Application")
        getIE.Visible = True
    End If
    Set shellWin = Nothing
End Function

Sub clearIE()
    Dim objIE As InternetExplorer
    Set objIE = getIE()
    objIE.quit
    Set objIE = Nothing
End Sub
Sub XWB_Extract()
'
'
'   Link w/ #REF# tag = http://10.83.19.7:10001/ietp-s1000d/viewDataModuleWindow.do?resourceId=DMC_#REF#&mode=html&target=
'
'
    Dim sheet As Worksheet
    Dim objIE As InternetExplorer
    Set objIE = getIE()
    Datarange = Sheets("InspectionGate").UsedRange.Value
    objIE.Visible = True
    For i = 2 To UBound(Datarange)
        objIE.Navigate2 Datarange(i, 7)
        
        Set sheet = getsheet(Mid(Datarange(i, 4), 12, 11))
        sheet.Columns(1).NumberFormat = "@"
        Do
        Loop While objIE.Busy Or objIE.readyState <> READYSTATE_COMPLETE
        
        Set objFrame = objIE.Document.frames("newframe")
        Set objTable = objFrame.Document.getElementsByTagName("table")
        MaxR = objTable.Length - 1
        l = 1
        Dim Editrange() As Variant
        ReDim Editrange(0 To MaxR)
        For j = 0 To MaxR
'            strText = Trim(objTable(j).innerText)
            Editrange(j) = Trim(objTable(j).innerText)
'            If Len(strText) = 3 And IsNumeric(strText) Then
'                sheet.Cells(l, 1) = strText
'                sheet.Cells(l, 2) = objTable(j + 1).innerText
'                l = l + 1
'            End If
        Next
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(MaxR + 1, 1)) = Editrange
        
        Set objTable = Nothing
        Set objFrame = Nothing
        Set sheet = Nothing
    Next
    objIE.quit
    Set objIE = Nothing
End Sub

Sub XWB_test()
'
'
'   Link w/ #REF# tag = http://10.83.19.7:10001/ietp-s1000d/viewDataModuleWindow.do?resourceId=DMC_#REF#&mode=html&target=
'
'
    Dim sheet As Worksheet
    Dim objIE As InternetExplorer
    Set objIE = getIE()
    Datarange = Sheets("InspectionGate").UsedRange.Value
    objIE.Visible = True
    For i = 2 To 20 'UBound(Datarange)
        objIE.Navigate2 Datarange(i, 7)
        
        Set sheet = getsheet(Mid(Datarange(i, 4), 12, 11))
        sheet.Columns(1).NumberFormat = "@"
        Do
        Loop While objIE.Busy Or objIE.readyState <> READYSTATE_COMPLETE
        
        Set objFrame = objIE.Document.frames("newframe")
        Set objTable = objFrame.Document.getElementsByTagName("td")
        MaxR = objTable.Length - 1
        l = 1
        Dim Editrange() As Variant
        For j = 0 To MaxR
            strText = Trim(objTable(j).innerText)
            If strText = "1.1" Then
                Content = True
            ElseIf Content And (InStr(1, strText, "Fig 1 ") <> 0 Or Left(strText, 21) = "Close-up requirements") Then
                Content = False
                Exit For
            End If
            If Content Then
                sheet.Cells(l, 1) = strText
                l = l + 1
            End If
        Next
        sheet.Columns(1).ColumnWidth = 120
        sheet.Rows.AutoFit
        Set objTable = Nothing
        Set objFrame = Nothing
        Set sheet = Nothing
    Next
    objIE.quit
    Set objIE = Nothing
End Sub

