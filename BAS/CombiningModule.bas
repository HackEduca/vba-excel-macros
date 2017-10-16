Attribute VB_Name = "CombiningModule"
'
'
' A supplement module for snapshot search result tables
'
'
' Combining all sheets (except "Result") into a single sheet.

Sub main()
    Set combine = ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    'combine.Name = "ALL"
    MaxL = 1
    For Each page In ActiveWorkbook.Sheets
        If page.Name <> "Result" Or page.Name <> "ALL" Then
            MaxR = page.UsedRange.Rows.Count
            If MaxR >= 2 Then
                MsgBox page.Cells(2, 1).text
                combine.Range(Cells(MaxL, 1), cells(MaxL)
                MaxL = MaxL + MaxR - 1
            End If
        End If
    Next
    Set combine = Nothing
End Sub
