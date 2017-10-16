VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 09 Mar 2017: cPlan is a test class with no actual applications deployed.
' Can be removed without problems.

Option Explicit

Private PlanCode As String
Private PlanName As String
Private PlanCounter As Integer

Private Sub Class_Initialize()
    PlanCode = "H0000"
    PlanName = "DUMMY NAME"
    PlanCounter = 0
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Get Code() As String
    Code = PlanCode
End Property

Public Property Let Code(ByVal newvalue As String)
    If IsPlan(newvalue) Then
        Code = newvalue
    End If
End Property

Public Property Get Name() As String
    Name = PlanName
End Property

Public Property Let Name(ByVal newvalue As String)
    Name = newvalue
End Property

Public Property Get Counter() As Integer
    Counter = PlanCounter
End Property

Public Property Let Counter(ByVal newvalue As Variant)
    If IsNumeric(CStr(newvalue)) Then
        Counter = CInt(newvalue)
    End If
End Property
