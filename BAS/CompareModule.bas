Attribute VB_Name = "CompareModule"
'
'
' A supplement module for snapshot search result tables
'
'
' Compare two results tables and highlight the matches.
' Use slave's rows to match with master's rows, if found, highlight the row in the master's table.

Sub main()
    For i = 1 To 2
        If i = 1 Then
            word = "master"
        Else
            word = "slave"
        End If
        MsgBox ("Please select the " & word & " file.")
        With Application.FileDialog(msoFileDialogOpen)
            .AllowMultiSelect = False
            .Show
            If i = 1 Then
                m = .SelectedItems(1)
            Else
                s = .SelectedItems(1)
            End If
        End With
    Next
    Set master = Workbooks.Open(m)
    Set slave = Workbooks.Open(s)
    colour = InputBox("Highlighting colour code")
    If Not IsNumeric(colour) Then
        colour = 65535
    Else
        colour = CInt(colour)
    End If
    
    ' SO UGLY
    ' NESTED STRUCTURES A
                        'H
                         'H
                          'H
                           'H
                            'H
                             'H
                              'H
                               'H
                                'H
                                 'H
                                  'H
                                   'H
                                    'H
                                     'H
                                      'H
                                       'H

    For Each m_shet In master.Worksheets
        With m_shet
            If .Name <> "Result" Then
                For Each s_shet In slave.Worksheets
                    If s_shet.Name = .Name Then
                        j = 2
                        Do While Not IsEmpty(.Cells(j, 1))
                            If .Cells(j, 3) = .Cells(j - 1, 3) And .Cells(j, 4) = .Cells(j - 1, 4) And .Cells(j, 5) = .Cells(j - 1, 5) Then
                                If .Cells(j - 1, 6).Interior.color = colour Then .Cells(j, 6).Interior.color = colour
                            Else
                                k = 2
                                triggered = False
                                Do While Not IsEmpty(s_shet.Cells(k, 2))
                                    If Not triggered And .Cells(j, 3) = s_shet.Cells(k, 3) Then
                                        triggered = True
                                        'Debug.Print .Cells(j, 3) & "...triggered"
                                    End If
                                    If triggered Then
                                        If .Cells(j, 4) = s_shet.Cells(k, 4) And .Cells(j, 5) = s_shet.Cells(k, 5) Then
                                            .Cells(j, 6).Interior.color = colour
                                            'Debug.Print "Record found, escaping"
                                            Exit Do
                                        ElseIf (.Cells(j, 3) <> s_shet.Cells(k, 3)) Then
                                            'Debug.Print "Reached the end of plan, escaping"
                                            Exit Do
                                        End If
                                    End If
                                    k = k + 1
                                Loop
                            End If
                            j = j + 1
                        Loop
                    End If
                Next
            End If
        End With
        Debug.Print m_shet.Name & " OK"
    Next
    
    slave.Close
    
End Sub
