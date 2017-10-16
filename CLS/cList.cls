VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'08 Sep 2016: Last piece (Probably). Imitation of the list() function & list class in Python2/3.
'Kono tatakai ga owattara, ora ha mura ni kaette kanojo to kekkon surunda...

'09 Mar 2017: There are a few functions (read: zl07_replace, & zl07_longtext_replacer) rely on this class, DO NOT REMOVE

Option Explicit

Private arr() As Variant

Private Sub Class_Initialize()
    ReDim arr(0 To 0)
    arr(0) = ""
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Get Value(Optional ByRef index As Variant = -1) As Variant
Attribute Value.VB_UserMemId = 0
    If index = -1 Then
        Value = arr
    Else
        Value = arr(index)
    End If
End Property

Public Sub Add(newvalue As Variant)
    If UBound(arr) = 0 And arr(0) = "" Then
        arr(0) = newvalue
    Else
        ReDim Preserve arr(0 To UBound(arr) + 1)
        arr(UBound(arr)) = newvalue
    End If
End Sub

Public Sub DeleteByValue(Value As Variant)
    Dim i As Long
    i = 0
    Do While i <= UBound(arr)
        If arr(i) = Value Then
            DeleteByIndex (i)
        Else
            i = i + 1
        End If
    Loop
End Sub

Public Sub DeleteByIndex(index As Long)
    Do While index < UBound(arr)
        arr(index) = arr(index + 1)
        index = index + 1
    Loop
    If UBound(arr) = 0 Then
        arr(0) = ""
    Else
        ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
    End If
End Sub

Public Property Get Count() As Long
    Count = UBound(arr) + 1
End Property

Public Property Get Length() As Long
    Length = UBound(arr) + 1
End Property

Public Property Get FindIndex(Value As Variant, Optional Direction As Integer = 0)
    
    
End Property
