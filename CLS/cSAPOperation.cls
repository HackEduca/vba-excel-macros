VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSAPOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 09 Mar 2017: Test class, no deployment except for tests.

Option Explicit

Private vName As String
Private vStrat As String
Private vPlan As String
Private vOp As String
Private vTest() As String

Private Sub Class_Initialize()
    vName = "Empty Operation"
    vStrat = "H00000"
    vPlan = "00"
    vOp = "0000"
    ReDim vTest(0 To 0)
    vTest(0) = "null"
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Get Operation() As String
    Operation = vOp
End Property

Public Property Let Operation(ByVal newvalue As String)
    Operation = newvalue
End Property

Public Sub Add(ByVal shit As String)
    If UBound(vTest) = 0 And vTest(0) = "null" Then
        vTest(0) = shit
    Else
        ReDim Preserve vTest(0 To UBound(vTest) + 1)
        vTest(UBound(vTest)) = shit
    End If
End Sub

Public Property Get vtestlist() As String()
    vtestlist = vTest
End Property

