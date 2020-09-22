Attribute VB_Name = "m_database"
Option Explicit

' ************************************************************

' m_database.bas - Longbow Database Access Module

' ************************************************************
' (c)Injected Software 2001, Coded by Dale Reidy

' Version 1.0

' Usage: Allows access to Longbow Database Files
' External Functions Required: None
' Visual Basic Version: 6

' For support using this module contact   dj_fuzzy_beast@hotmail.com

' Needs...

' GetCharReps [SUB] [main.bas/Longbow Server]

' ************************************************************

Dim database_data$() ' Redim for num of sockets

Public Function IssueDatabaseCommand(databasefilename As String, param1 As String, param2 As String) As String
    ' *** This function will probably be replaced with easier to implement functions like the many shown below ***
    
    
    '  replace NAME (where NAME = 'Keith') and (where PASSWORD = 'Letmein') with 'Keefy'
    

End Function

Public Function DeleteEntry(databasefilename As String, fieldname As String, fielddata As String) As Integer
    
    'On Error Resume Next
    
    DeleteEntry = -1
    Dim a%, b%, t%, u$, k$(30), fd1%, f2%, tcc$
    fd1% = FreeFile
    Open databasefilename For Input As #fd1
    f2% = FreeFile
    Open "c:\" & Trim$(Str$(fd1)) & ".tmp" For Output As #f2
    
    Input #fd1, a
    Print #f2, a
    
    For t = 1 To a
        Input #fd1, tcc$
        Print #f2, tcc$;
        If t < a Then Print #f2, ","; Else Print #f2, ""
        If tcc$ = fieldname$ Then b = t
    Next t
    
    Do Until EOF(fd1)
    
        Erase k$()
        
        For t = 1 To a
            Input #fd1, k$(t)
        Next t
        
        If k$(b) = fielddata$ Then
            Erase k$()
        Else
            For t = 1 To a
                Print #f2, k$(t);
                If t < a Then Print #f2, ",";
                If t = a Then Print #f2, ""
            Next t
        End If
    Loop
    
    Close fd1, f2

    SetAttr databasefilename, vbNormal
    Kill databasefilename
    FileCopy "c:\" & Trim$(Str$(fd1)) & ".tmp", databasefilename
    Kill "c:\" & Trim$(Str$(fd1)) & ".tmp"
    DeleteEntry = 1
End Function


Public Function GetNumOfDatabaseFields(databasefilename As String) As Integer
    On Error Resume Next
    Dim fd1%, temp$
    fd1 = FreeFile
    Open databasefilename For Input As #fd1
        Input #fd1, temp$
    Close fd1
    GetNumOfDatabaseFields = Val(temp$)
End Function

Public Function GetFieldNames(databasefilename As String) As String
    On Error Resume Next
    Dim fd1%, temp$
    fd1 = FreeFile
    Open databasefilename For Input As #fd1
        Input #fd1, temp$
        Line Input #fd1, temp$
    Close fd1
    GetFieldNames = temp$
End Function

Public Function GetFieldName(databasefilename As String, fieldnumber) As String
    On Error Resume Next
    Dim fd1%, itemp%, temp1$, temp2$, t%
    fd1 = FreeFile
    Open databasefilename For Input As #fd1
        Input #fd1, itemp%
        If fieldnumber > itemp Or fieldnumber < 1 Then
            Close fd1
            Exit Function
        End If
        For t = 1 To fieldnumber
            Input #fd1, temp2$
        Next t
    Close fd1
        GetFieldName = temp2$
End Function

Public Function GetEntry(databasefilename As String, entrynumber As Integer, fieldname As String) As String
    
    ' Get data from a specified field in an entry
    
    On Error Resume Next
    Dim fd1%, temp1$, temp2$, itemp1%, itemp2%, t%
    fd1 = FreeFile
    Open databasefilename For Input As #fd1
        Input #fd1, itemp1
        For t = 1 To itemp1%
            Input #fd1, temp1$
            If temp1$ = fieldname$ Then itemp2 = t
        Next t
        For t = 1 To entrynumber - 1
            Line Input #fd1, temp1$
        Next t
        For t = 1 To itemp2
            Input #fd1, temp1$
        Next t
    Close fd1
        GetEntry = temp1$
End Function

Public Function EntryExist(databasefilename As String, fieldname As String, fielddata As String, comparisontype As String) As Integer
    
    ' Checks to see if an entry exists in the database
    
    On Error Resume Next
    Dim fd1%, itemp1%, itemp2%, temp1$, temp2$, t%, u%, do_exist%
    fd1 = FreeFile
    Open databasefilename For Input As #fd1
    ' Get the number of the fieldname
        Input #fd1, itemp1
        For t = 1 To itemp1%
            Input #fd1, temp1$
            If temp1$ = fieldname$ Then itemp2 = t
        Next t
    ' Now scan through all the entries, look at the specific field and check the data
        Do Until EOF(fd1)
            For t = 1 To itemp1%
                Input #fd1, temp2$
                If t = itemp2 Then
                    Select Case comparisontype
                        Case ">"
                            If fielddata$ < temp2$ Then do_exist = 1
                        Case "<"
                            If fielddata$ > temp2$ Then do_exist = 1
                        Case "="
                            If fielddata$ = temp2$ Then do_exist = 1
                        Case "<>"
                            If fielddata$ <> temp2$ Then do_exist = 1
                    End Select
                End If
            Next t
        Loop
    Close fd1
    EntryExist = do_exist
End Function

Public Function ReplaceEntry(databasefilename As String, findfieldname As String, findfielddata As String, findfieldcomparison As String, changefieldname As String, changefielddata As String) As Integer
    
    ' Replaces Specified Field In An Entry Or Multiple Entries In The Database
    
    On Error Resume Next
    Dim fi%, a$, b$, c$, d$, e%, f%, g%, numfields%, t%, u%, f2%, k$(30)
    
    fi = FreeFile
    Open databasefilename For Input As #fi
    f2 = FreeFile
    Open "c:\" & Trim$(Str$(fi)) & ".tmp" For Output As #f2
    Input #fi, numfields
    Print #f2, numfields
    For t = 1 To numfields
        Input #fi, a$
        Print #f2, a$;
        If t < numfields Then Print #f2, ","; Else Print #f2, ""
        
        If a$ = findfieldname Then f = t
        If a$ = changefieldname Then g = t
    Next t
    
    Do Until EOF(fi)
        Erase k$()
        For t = 1 To numfields
            Input #fi, k$(t)
        Next t
        
        Select Case findfieldcomparison
            Case ">"
                If k$(f) < findfielddata$ Then k$(g) = changefielddata$: ReplaceEntry = 1
            Case "<"
                If k$(f) > findfielddata$ Then k$(g) = changefielddata$: ReplaceEntry = 1
            Case "="
                If k$(f) = findfielddata$ Then k$(g) = changefielddata$: ReplaceEntry = 1
            Case "<>"
                If k$(f) <> findfielddata$ Then k$(g) = changefielddata$: ReplaceEntry = 1
        End Select
        
        For t = 1 To numfields
            Print #f2, k$(t);
            If t < numfields Then Print #f2, ","; Else Print #f2, ""
        Next t
    Loop
    
    Close fi, f2
    
    SetAttr databasefilename, vbNormal
    Kill databasefilename
    FileCopy "c:\" & Trim$(Str$(fi)) & ".tmp", databasefilename
    Kill "c:\" & Trim$(Str$(fi)) & ".tmp"
End Function

Public Function GetEntryNum(databasefilename As String, startent As Integer, fieldname As String, fielddata As String, fieldcomparison As String) As Long

    ' Get entry number of next entry in the database whose field data complies with the parameters passed
    
    On Error Resume Next
    
    Dim fi%, a$, b$, c%, d%, e%, t%, u%, numfields%, foundone%, cent%
    
    fi = FreeFile
    
    Open databasefilename For Input As #fi
    
    Input #fi, numfields
    
    For t = 1 To numfields
        Input #fi, a$
        If a$ = fieldname$ Then e = t
    Next t
    
    startent = startent - 1
        
    For t = 1 To startent
        Line Input #fi, a$
    Next t
    
    Do Until EOF(fi)
        cent = cent + 1
        For t = 1 To numfields
            Input #fi, a$
                If t = e Then
                    foundone = -1
                    Select Case fieldcomparison
                        Case "<"
                            If a$ > fielddata$ Then foundone = startent + cent
                        Case ">"
                            If a$ < fielddata$ Then foundone = startent + cent
                        Case "="
                            If a$ = fielddata$ Then foundone = startent + cent
                        Case "<>"
                            If a$ <> fielddata$ Then foundone = startent + cent
                    End Select
                    If foundone <> -1 Then Close fi: GetEntryNum = foundone: Exit Function
                End If
        Next t
    Loop
    
    Close fi
    GetEntryNum = -1
End Function

Public Function CreateDatabase(databasefilename As String, numfields As String, fieldnames As String) As Integer
    On Error Resume Next
    
    If Exists(databasefilename) = 1 Then CreateDatabase = -1: Exit Function
    
    Dim fi%
    
    fi = FreeFile
    
    Open databasefilename For Output As #fi
        Print #fi, numfields
        Print #fi, fieldnames
    Close fi
    
    CreateDatabase = 1
End Function

Public Function AddEntry(databasefilename As String, fielddata As String) As Long
    ' Adds entry to the database, will return -1 if it failed or the entry number in the database
    
    ' Return values
    ' -1, Not added
    ' -2, Added, but not found in database, possible corrupt data
    ' x, Added and found @ this entry number
    
    AddEntry = -1
    
    On Error Resume Next
    
   Dim fi%, a%, b$, cent%, dbnumfields%
   
   fi = FreeFile
   ' Append the entry to the database
   
    'Debug.Print databasefilename
    
    Open databasefilename For Input As #fi
        Input #1, dbnumfields%
    Close fi
   
    'Debug.Print dbnumfields - 1
    'Debug.Print GetCharReps(fielddata$, ",")
    
    If GetCharReps(fielddata$, ",") <> dbnumfields - 1 Then Exit Function
    
    Open databasefilename For Append As #fi
        Print #fi, fielddata$
    Close fi
    ' Find the entry
    Open databasefilename For Input As #fi
        Input #fi, b$
        Line Input #fi, b$
        Do Until EOF(fi)
            cent = cent + 1
            Line Input #fi, b$
            If fielddata$ = b$ Then AddEntry = cent: Close fi: Exit Function
        Loop
    Close fi
    AddEntry = -2
    
End Function
