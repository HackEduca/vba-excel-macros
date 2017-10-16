Attribute VB_Name = "Global_Callables"
'READABILITY COUNTS.'
'SIMPLICITY COUNTS.'
'COMMENTING SAVES TIME AND LIVES.'

Public Const XWB_Link As String = "http://10.83.19.7:10001/ietp-s1000d/viewDataModuleWindow.do?resourceId=DMC_#REF#&mode=html&target="

Public Enum Order
    AscOrder
    DescOrder
End Enum

Public Sub Reboot()
    Shell ("shutdown -r")
End Sub

Public Sub Shutdown()
    Shell ("shutdown -s")
End Sub

Function fibon(itermax As Integer)
    Dim a, b, c As Double
    a = 0
    b = 1

    For i = 1 To itermax
        a = a + b
        c = a
        a = b
        b = c
    Next
    fibon = a
End Function
Sub morse_callable_marco()
    For Each item In Selection
        item.Cells(1, 2) = morse(item)
    Next
End Sub

Function morse(ByVal m As String) As String
    t = list(LCase(m))
    p = Split("a*_b*_c*_d*_e*_f*_g*_h*_i*_j*_k*_l*_m*_n*_o*_p*_q*_r*_s*_t*_u*_v*_w*_x*_y*_z*_0*_1*_2*_3*_4*_5*_6*_7*_8*_9*_ *_.*_,*_?*_'*_!*_/*_(*_)*_&*_:*_;*_=*_+*_-*__*_""*_$*_@", "*_")
    c = Split(".-*_-...*_-.-.*_-..*_.*_..-.*_--.*_....*_..*_.---*_-.-*_.-..*_--*_-.*_---*_.--.*_--.-*_.-.*_...*_-*_..-*_...-*_.--*_-..-*_-.--*_--..*_-----*_.----*_..---*_...--*_....-*_.....*_-....*_--...*_---..*_----.*_/*_.-.-.-*_--..--*_..--..*_.----.*_-.-.--*_-..-.*_-.--.*_-.--.-*_.-....*_---...*_-.-.-.*_-...-*_.-.-.*_-....-*_..--.-*_.-..-.*_...-..-*_.--.-.", "*_")
    r = ""
    For Each char In t
        num = 0
        Do While char <> p(num)
            num = num + 1
            If num = 55 Then
                num = 36
                Exit Do
            End If
        Loop
        r = r & c(num) & " "
    Next
    morse = Trim(r)
End Function

' IsUCase, IsLCase, IsAlpha, IsAlphaNumeric, IsSpace - Boolean checks on string
'
' Can you imagine how come VB doesn't come with these stupidly simple functions?!
' I mean, FUCK Microsoft, I can do it better.

Function IsUCase(ByVal str As String) As Boolean
    Dim chars() As String
    chars = list(str)
    IsUCase = True
    For Each char In chars
        If (Asc(char) < 65 Or Asc(char) > 90) Then
            IsUCase = False
            Exit Function
        End If
    Next
End Function

Function IsLCase(ByVal str As String) As Boolean
    Dim chars() As String
    chars = list(str)
    IsLCase = True
    For Each char In chars
        If (Asc(char) < 97 Or Asc(char) > 122) Then
            IsLCase = False
            Exit Function
        End If
    Next
End Function

Function IsAlphaNumeric(ByVal str As String) As Boolean
    Dim chars() As String
    chars = list(str)
    IsAlphaNumeric = True
    For Each char In chars
        If Not IsNumeric(char) And Not IsAlpha(char) Then
            IsAlphaNumeric = False
            Exit Function
        End If
    Next
End Function

Function IsAlpha(ByVal str As String) As Boolean
    Dim chars() As String
    chars = list(str)
    IsAlpha = True
    For Each char In chars
        If (Asc(char) < 97 Or Asc(char) > 122) And (Asc(char) < 65 Or Asc(char) > 90) Then
            IsAlpha = False
            Exit Function
        End If
    Next
End Function

Function IsSpace(ByVal str As String) As Boolean
    Dim chars() As String
    chars = list(str)
    IsSpace = True
    For Each char In chars
        If char <> " " Then
            IsSpace = False
            Exit Function
        End If
    Next
End Function

Function RoundUp(Number) As Integer
    inta = CInt(Split(CStr(Number), ".")(0))
    If Number <> inta Then Number = inta + 1
    RoundUp = Number
End Function
Function RoundDown(Number) As Integer
    RoundDown = CInt(Split(CStr(Number), ".")(0))
End Function

' list() - Converting string into array of strings of specified length

Function list(InputString As String, Optional Length As Integer = 1) As String()
    l = RoundUp(Len(InputString) / Length) - 1
    Dim segment() As String
    If l = -1 Then
        ReDim segment(0)
        segment(0) = ""
        list = segment
    Else
        ReDim segment(0 To l)
        For i = 0 To l
            segment(i) = Mid(InputString, (i * Length + 1), Length)
        Next
        list = segment
    End If
End Function

Function compare(ByVal str1 As String, ByVal str2 As String) As Single
    ' Compaing two strings, returns similiarity as float.
    If str1 <> str2 Then
        Dim list1() As String
        Dim list2() As String
        Dim Count, countT As Integer
        
        Count = 0
        countT = 0
        
        If Len(str1) <= Len(str2) Then
            list1 = compare_str(str1)
            list2 = compare_str(str2)
        Else
            list1 = compare_str(str2)
            list2 = compare_str(str1)
        End If
        
        For Each x In list1
            For Each y In list2
                If x = y Then
                    Count = Count + 1
                End If
            Next
        Next
        
        For Each x In list2
            For Each y In list2
                If x = y Then
                    countT = countT + 1
                End If
            Next
        Next
        
        compare = Count / countT
    Else
        compare = 1
    End If
End Function

Function compare_str(ByVal str As String) As String()
    ' Sub-Function for compare()
    ' Variation of list(), but each len(item) = 2
    ' OBSOLETE as list() now offers optional length input as of 2016/2017. Just use list(___, 2) instead.
    Dim list1() As String
    Dim com1() As String
    ReDim list1(1 To Len(str))
    ReDim com1(1 To UBound(list1) - 1)
    list1 = list(str)
    For i = 1 To (UBound(com1))
        com1(i) = list1(i) & list1(i + 1)
    Next
    compare_str = com1
End Function

Function get_folder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Function
        Else
            get_folder = .SelectedItems(1)
            Exit Function
        End If
    End With
End Function

Function get_file(ByVal path As String, Optional ByVal filter As String = "") As String()
    Dim list() As String
    Dim ext As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.GetFolder(path)
        ReDim list(0)
        i = 0
        For Each item In .Files
            ext = Mid(item.Name, InStrRev(item.Name, "."))
            If filter = "" Or LCase(filter) = LCase(ext) Then
                list(i) = item.Name
                i = i + 1
                ReDim Preserve list(0 To i)
            End If
        Next item
    End With
    ReDim Preserve list(i - 1)
    get_file = list
    Set fso = Nothing
End Function

Sub testfso()
    Dim a() As String
    a = get_file(get_folder())
    For Each item In a
        MsgBox item
    Next item
End Sub

Sub escmsg()
    MsgBox "Invalid input. Escaping...", vbExclamation
End Sub

Function IsPlan(ByVal in_str As String) As Boolean
    in_str = Trim(in_str)
    If Left(in_str, 1) = "H" And Len(in_str) >= 4 And Len(in_str) <= 8 Then
        IsPlan = True
    Else
        IsPlan = False
    End If
End Function

Function unique_list(ByRef arr As Variant) As String()
    Dim re() As String
    ReDim re(0)
    pre = ""
    For Each item In arr
        If item <> pre Then
            pre = item
            re(UBound(re)) = item
            ReDim Preserve re(UBound(re) + 1)
        End If
    Next
    If UBound(re) <> 0 Then
        ReDim Preserve re(UBound(re) - 1)
    End If
    unique_list = re
End Function

Function filename_normalize(ByRef filename As String) As String
    Dim crit() As String
    crit() = Split("\ / * ? "" < > | : [ ]")
    For Each char In crit
        filename = Replace(filename, char, "_")
    Next
    filename_normalize = filename
End Function

Sub getfiles(ByVal DPath As String, list() As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fo = fso.GetFolder(DPath)
    If fo.Files.Count > 0 Then
        For Each file In fo.Files
            If InStr(1, file, ".xls") <> 0 Then
                list(UBound(list)) = file
                ReDim Preserve list(0 To UBound(list) + 1)
            End If
        Next
    End If
    If fo.SubFolders.Count > 0 Then
        For Each SubFo In fo.SubFolders
            getfiles SubFo, list
        Next
    End If
End Sub

' I thought I don't need explainations on this shit?
Function max(l As Variant) As Single
    rec = l(LBound(l))
    For i = LBound(l) To UBound(l)
        If l(i) > rec Then
            rec = l(i)
        End If
    Next
    max = rec
End Function

' #BroFistTwiceADay
Function min(l As Variant) As Single
    rec = l(LBound(l))
    For i = LBound(l) To UBound(l)
        If l(i) < rec Then
            rec = l(i)
        End If
    Next
    min = rec
End Function

' Exploiting VBA's weird properties
Function HasValue(anArray() As Variant) As Boolean
    HasValue = Not Not anArray
End Function
Sub arrtest()
    Dim lol() As Variant
    Sort lol
End Sub
Function Sort(arr As Variant, Optional Order As Order = Order.AscOrder) As Variant()
    If VarType(arr) > vbArray Then
        TempArr = arr
        i = LBound(TempArr)
        Do While i =
       
    Else
        Sort = arr
    End If
End Function
Function avg(ary As Variant) As Single
    Count = UBound(ary) - LBound(ary)
    total = 0
    For Each item In ary
        total = total + item
    Next
    avg = total / Count
End Function
Function StrInt(strg() As String)
    Dim tmp() As Integer
    ReDim tmp(LBound(strg) To UBound(strg))
    For i = LBound(strg) To UBound(strg)
        tmp(i) = Asc(strg(i))
    Next
    StrInt = tmp
End Function
Function IntStr(strg() As Integer) As String()
    Dim tmp() As String
    ReDim tmp(LBound(strg) To UBound(strg))
    For i = LBound(strg) To UBound(strg)
        tmp(i) = Chr(strg(i))
    Next
    IntStr = tmp
End Function
Function IntHex(strg() As Integer) As String()
    Dim tmp() As String
    ReDim tmp(LBound(strg) To UBound(strg))
    For i = LBound(strg) To UBound(strg)
        tmp(i) = hex(strg(i))
        If Len(tmp(i)) = 1 Then
            tmp(i) = "0" & tmp(i)
        End If
    Next
    IntHex = tmp
End Function
Function HexInt(strg() As String)
    Dim tmp() As String
    ReDim tmp(LBound(strg) To UBound(strg))
    For i = LBound(strg) To UBound(strg)
        tmp(i) = dec(strg(i))
    Next
    HexInt = tmp
End Function
Function StrHex(strg() As String) As String()
    StrHex = IntHex(StrInt(strg()))
End Function
Function HexStr(strg() As String) As String()
    HexStr = IntStr(HexInt(strg()))
End Function
Function dec(hex As String)
    strg = list(hex)
    For i = 0 To Size(strg) - 1
        If IsNumeric(strg(i)) Then
            strg(i) = CInt(strg(i))
        Else
            strg(i) = Asc(strg(i)) - 55
        End If
        strg(i) = strg(i) * (16 ^ (UBound(strg) - i))
    Next
    total = 0
    For Each num In strg
        total = total + num
    Next
    dec = total
End Function
Function enc(message() As Integer, key() As Integer) As Integer()
    Dim temp() As Integer
    ReDim temp(LBound(message) To UBound(message))
    For i = LBound(message) To UBound(message)
        temp(i) = message(i) Xor key(i)
    Next
    enc = temp
End Function
Function ArrayPrint(ByVal SourceArray As Variant, Optional Delimiter As String = "") As String
    If IsArray(SourceArray) Then
        tmp = ""
        For Each item In SourceArray
            tmp = tmp & ArrayPrint(item) & Delimiter
        Next
        ArrayPrint = Left(tmp, Len(tmp) - Len(Delimiter))
    Else
        ArrayPrint = CStr(SourceArray)
    End If
End Function
Function Size(Variable) As Integer
    If Not IsArray(Variable) Then
        Size = 1
    Else
        Size = UBound(Variable) - LBound(Variable) + 1
    End If
End Function
Function ascii_33_to_126()
    Dim tmp() As String
    ReDim tmp(0 To 93)
    For i = 33 To 126
        tmp(i - 33) = Chr(i)
    Next
    ascii_33_to_126 = tmp
End Function
Sub Cracktest(Optional sample As String)
    Dic = ascii_33_to_126() 'list("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
    Dim trial As Long
    If sample = "" Then
        Length = 3
        sample = RandStringGenerate(Length)
    Else
        Length = Len(sample)
    End If
    Debug.Print "Length = " & Length & ", Sample = " & sample & ", Mean Time = " & (Size(Dic) ^ Length) / 2
    tstart = Timer()
    trial = 0
    Do
        try = BruteForce(trial, Length, Dic)
        trial = trial + 1
    Loop While sample <> try
    totaltime = Timer() - tstart
    Debug.Print "Cracked. Trial time: " & trial & ", Time used = " & totaltime & "s, Avg. time per trial = " & totaltime / trial & "s"
End Sub
Function BruteForce(ByVal Number, Length, Dic) As String
    Codex = Number
    base = Size(Dic)
    Dim txt() As String
    ReDim txt(0 To Length - 1)
    If Codex >= base ^ Length Then Codex = (base ^ Length) - 1
    For i = UBound(txt) To 0 Step -1
        char = Codex Mod base
        txt(i) = Dic(char)
        Codex = (Codex - char) / base
    Next
    BruteForce = ArrayPrint(txt)
End Function
Function RandStringGenerate(ByVal Length As Long)
    Dim password() As String
    ReDim password(1 To Length)
    For i = LBound(password) To UBound(password)
        password(i) = get_random_char()
    Next
    st = ""
    For Each char In password
        st = st & char
    Next
    RandStringGenerate = st
End Function
Function get_random_char() As String
    Do
        rn = CInt(Rnd() * 126)
    Loop While rn < 33
    lol = Chr(rn)
    get_random_char = lol
End Function

Function Dequote(text) As String
    Dequote = Mid(Trim(text), 2, Len(Trim(text)) - 2)
End Function
