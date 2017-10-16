Attribute VB_Name = "Testing_Ground"
' Trying out all things in the name of experiments
' Use at your own risk

Sub TypeTest()
    MsgBox VarType(Cells(1, 1))
    
End Sub

Sub FSOTesting()

    Dim fso, folder, item As Object
    Dim path As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    path = "Z:\TechData\G\PCdrawing\RB211_524GHT\"
    
    Set folder = fso.GetFolder(path)
    Set item = folder.Files
    
    For Each i In item
        If Right(i.Name, 11) = "20SEP14.pdf" Then
            MsgBox i.Name
        End If
    Next
    

End Sub

Sub sorttest()

    Dim RMax, CMax, NewR, r, c, Count As Long
    
    RMax = ActiveSheet.UsedRange.Rows.Count
    CMax = ActiveSheet.UsedRange.Columns.Count
    
    NewR = RMax * CMax
    
    r = RMax
    
    Do While r >= 1
        c = CMax
        Do While c >= 1
            Cells(NewR, 1).Value = Cells(r, c).Value
            NewR = NewR - 1
            If NewR <> 1 Then
                Cells(r, c) = ""
            End If
            c = c - 1
        Loop
        r = r - 1
    Loop
    
    NewR = (RMax * CMax)

    r = 1
    Count = 1

    Do While Count <= NewR
        If IsEmpty(Cells(r, 1)) Or Len(Cells(r, 1)) <= 2 Then
'            Cells(R, 1).Delete Shift:=xlUp
            Cells(r, 1).EntireRow.Delete
        Else
            r = r + 1
        End If
        Count = Count + 1
    Loop
    
    
End Sub

Sub remember_no_dups()  ' Deleting older TV Issues entries copied from TVF2013.
    Dim Count As Integer
    Count = 1
    Do While IsEmpty(Cells(Count, 2)) = False
        Count = Count + 1
    Loop
    
    r = 3
    Do While r < Count
        If Cells(r, 3) = Cells(r - 1, 3) And Cells(r, 12) = Cells(r - 1, 12) Then
            If CDate(Cells(r, 6)) > CDate(Cells(r - 1, 6)) Then
                Rows(r - 1).Delete
            ElseIf CDate(Cells(r, 6)) < CDate(Cells(r - 1, 6)) Then
                Rows(r).Delete
            ElseIf CDate(Cells(r, 6)) = CDate(Cells(r - 1, 6)) Then
                Cells(r, 6).Interior.color = 65535
                Cells(r - 1, 6).Interior.color = 65535
                r = r + 1
            End If
        Else
            r = r + 1
        End If
    Loop
End Sub

Sub beeptest()
    Beep
    Application.Wait [Now() + "0:00:00.2"]
    Beep
    Application.Wait [Now() + "0:00:00.2"]
    Beep
End Sub

Sub ChangeFileNames()
    '11 Dec 15, Working
    Dim fso, folder, item As Object
    Dim path As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = "C:\Users\orix.auyeung\Desktop\New folder\HellOnEarth\6508\"
    Set folder = fso.GetFolder(path)
    Set item = folder.Files
    
    For Each i In item
        If Left(i.Name, 9) = "6508_Part" Then
            i.Name = "GHT_FRS6508_FIG71-11-03-990-" & CInt(Left(Right(i.Name, 6), 2)) + 324 & "-20SEP14.pdf"
        End If
    Next
    
End Sub

Public Sub Playful()
'
' Macro1 Macro
'
    Dim r, g, b As Integer
    r = 255
    g = 0
    b = 0

    For i = 0 To 3 Step 1
        Do While r <> 0
            Cells.Interior.color = RGB(r, g, b)
            'Application.Wait [Now() + "0:00:00.01"]
            change_color r, g
        Loop
        
        Do While g <> 0
            Cells.Interior.color = RGB(r, g, b)
            'Application.Wait [Now() + "0:00:00.01"]
            change_color g, b
        Loop
        
        Do While b <> 0
            Cells.Interior.color = RGB(r, g, b)
            'Application.Wait [Now() + "0:00:00.01"]
            change_color b, r
        Loop
    Next
End Sub

Public Sub change_color(ByRef c1 As Variant, ByRef c2 As Variant)
    c1 = c1 - 1
    c2 = c2 + 1
End Sub

Sub returncolor()
    MsgBox Selection.Interior.color
End Sub

Sub classTest()
    Dim op As cSAPOperation
    Set op = New cSAPOperation
    For Each item In op.vtestlist
        MsgBox item
    Next
    For i = 1 To 10
        op.Add (CStr(i))
    Next
    For Each item In op.vtestlist
        MsgBox item
    Next
End Sub

Sub split_into_cell()
    i = 1
    Do While Not IsEmpty(Cells(i, 1))
        coll = Split(Cells(i, 1))
        For Each item In coll
            Cells(i, 2 + item) = item
        Next
    Loop
End Sub

Sub lolWUT()
    Set maint = Columns("C:C").Find(what:="MntPack.", LookIn:=xlValues, LookAt:=xlWhole)
    MsgBox (maint Is Nothing)
End Sub

Sub clisttest()
    Dim lol As clist
    Set lol = New clist
    For i = 0 To 10
        lol.Add "lol"
    Next
    lol.DeleteByValue "lol"
    For Each item In lol.Value
        MsgBox item
    Next
End Sub
Sub oneoff()
    For Each cell In ActiveSheet.UsedRange.Columns(1).Cells
        pos = 1
        If Left(cell, 3) = "SAP" Then
            pos = 3
        ElseIf Left(cell, 3) = "PDS" Then
            pos = 4
        End If
        cell.Offset(0, 1) = Left(cell, pos) & "-" & Mid(cell, pos + 1)
    Next
End Sub

Sub oneoff2()
    i = 2
    Do While Not IsEmpty(Cells(i, 1))
        mystr = Cells(i, 1).text
        If mystr <> "" Then
            If IsLCase(Right(mystr, 1)) Then mystr = Left(mystr, Len(mystr) - 1)
            If InStr(1, Cells(i - 1, 1).text, mystr) And Cells(i - 1, 1).Interior.color = 255 Then Cells(i, 1).Interior.color = 65535
        End If
        i = i + 1
    Loop
End Sub
Sub oneoff3()
    mystr = Cells(1, 1).text
    MsgBox
End Sub
Sub Macro1()
    Range("V23").NumberFormat = "dd mmm yyyy"
    ActiveWorkbook.Save
    ActiveWindow.SelectedSheets.PrintOut copies:=1, collate:=True, IgnorePrintAreas:=False
End Sub

Sub replace_file_name()
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = get_folder()
    filelist = get_file(path)
    For Each file In filelist
        file = path & "\" & file
        fso.MoveFile file, Replace(file, " Issue ", "_")
    Next
End Sub

Sub ran()
    Dim invset(1 To 1000000) As Single
    For i = 1 To 1000000
        invset(i) = Rnd
    Next
    Debug.Print "Max: " & max(invset)
    Debug.Print "Min: " & min(invset)
    Debug.Print "Avg: " & avg(invset)
End Sub

Sub wall_of_text()
    For i = 1 To 500
        ActiveSheet.Cells(i, 1) = ArrayPrint(list(RandStringGenerate(200), 5), ", ")
    Next
End Sub

Sub stat()
    mylist = list(RandStringGenerate(10000))
    u = 0
    l = 0
    n = 0
    For Each char In mylist
        If IsUCase(char) Then
            u = u + 1
        ElseIf IsLCase(char) Then
            l = l + 1
        Else
            n = n + 1
        End If
    Next
    Debug.Print u & " " & l & " " & n
End Sub
Sub enc_trial()
    Dim elist() As Integer
    Dim m() As Integer
    Dim k() As Integer
    ReDim m(1 To 50)
    ReDim k(1 To 50)
    mstr = ""
    kstr = ""
    estr = ""
    For i = LBound(m) To UBound(m)
        m(i) = CInt(Rnd() * 255)
        k(i) = CInt(Rnd() * 255)
    Next
    elist = enc(m, k)
    Debug.Print ArrayPrint(IntHex(m), "")
    Debug.Print ArrayPrint(IntHex(k), "")
    Debug.Print ArrayPrint(IntHex(elist), "")
End Sub

Sub testcell()
    For i = 1 To 100
        Cells(i, 1) = BruteForce(i - 1, 3)
    Next
End Sub

Function risktest()
    principal = 5000
    gain = 125
    Prob = 0.33333
    ratio = 3
    Wave = 176
    For i = 1 To Wave
        If Prob >= Rnd() Then
            principal = principal + gain
        Else
            principal = principal - (gain / ratio)
        End If
    Next
    risktest = principal
End Function

Sub finstab()
    total = 0
    trial = 100000
    For i = 1 To trial
        total = total + risktest()
    Next
    Debug.Print total / trial
End Sub

Sub zeno_paradox()
    Dim init, dist As Double
    init = 1000000
    dist = 0
    For i = 1 To 10000
        dist = dist + init
        init = init / 2
        Debug.Print i & ": " & dist
    Next
End Sub

Sub STD_Read()
    Dim myfile As String
    myfile = Application.GetOpenFilename
    If myfile = "False" Then Exit Sub
    Open myfile For Input As #1
    Line Input #1, text
    Debug.Print Dequote(text)
    Close #1
End Sub

Sub STD_Write()
    Dim myfile As String
    myfile = Application.GetOpenFilename
    If myfile = "False" Then Exit Sub
    Open myfile For Output As #1
    Write #1, RandStringGenerate(100)
    Close #1
End Sub

Sub emptytest()
    Debug.Print Application.CountA(Columns(5))
End Sub

Sub array2d_test()
    Dim ary() As Long
    aryLength = 900
    aryWidth = 900
    ReDim ary(aryLength - 1, aryWidth - 1)
    For i = 1 To aryLength
        For j = 1 To aryWidth
            ary(i - 1, j - 1) = (i * j)
        Next
    Next
    ActiveSheet.Cells(1, 1).Resize(aryLength, aryWidth) = ary
End Sub

Sub haha()
    For Each item In Selection
        item.Value = RandStringGenerate(1000)
    Next
End Sub
