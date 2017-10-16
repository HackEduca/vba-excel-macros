Attribute VB_Name = "CountingModule"
'
'
' A supplement module for snapshot search result tables
'
'
' Counting the hit counts, oepration counts and plan counts for the "Result" Summary page
' recount() can be ran independently to recalculate the hit counts after the search table is generated.

Sub main()
    plan_count = 0
    op_count = 0
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name <> "Result" Then
            count_inv sheet, plan_count, op_count, 0
        End If
    Next
    MsgBox "Total plans: " & plan_count & vbCrLf & "Total ops: " & op_count
End Sub

Sub count_inv(sheet As Variant, ByRef plan_count As Variant, ByRef op_count As Variant, ByRef hit_count As Variant, Optional StartRow As Long = 2)
    plan_name = ""
    op_name = ""
    With sheet
        Do While Not IsEmpty(.Cells(StartRow, 2))
            hit_count = hit_count + 1
            If plan_name <> .Cells(StartRow, 3) Then
                plan_count = plan_count + 1
                plan_name = .Cells(StartRow, 3)
                op_count = op_count + 1
                op_name = .Cells(StartRow, 5)
            ElseIf op_name <> .Cells(StartRow, 5) Then
                op_count = op_count + 1
                op_name = .Cells(StartRow, 5)
            End If
            StartRow = StartRow + 1
        Loop
    End With
End Sub

Sub recount()
    Dim r As Object
    Set r = ActiveWorkbook.Sheets("result")
    i = 3
    For Each sheet In ActiveWorkbook.Sheets
        With r
            If sheet.Name <> .Name Then
                Do While sheet.Name <> .Cells(i, 1)
                    i = i + 1
                    .Cells(i, 2) = 0
                    .Cells(i, 3) = 0
                    .Cells(i, 4) = 0
                Loop
                plan_count = 0
                op_count = 0
                hit_count = 0
                count_inv sheet, plan_count, op_count, hit_count
                .Cells(i, 2) = hit_count
                .Cells(i, 3) = op_count
                .Cells(i, 4) = plan_count
                i = i + 1
            End If
        End With
    Next
End Sub
