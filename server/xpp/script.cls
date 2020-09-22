VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private code$(3000)
Private std$(1000), stn$(1000), ind(1000), inn$(1000)
Private gs(10) ' Gosub Stack
Private tots, toti As Long ' Total number of strings and integers
Private CURRENTSOCK As Integer ' Stores the current socket for subs/funcs which dont have it passed as one of their param

Private Const DB_EXTENDINFO = 0

Private file(10) As Integer ' File access numbers
Private filea(10) As Integer ' File access type

Public Sub Execute(sck As Integer, filename As String, postdata As String)
   'On Error Resume Next
   On Error GoTo Execute_Error
   CURRENTSOCK = sck
   tots = 1: toti = 1
   Dim DESTX%, DESTY%, SRCX%, SRCY%, RECTWIDTH%, RECTHEIGHT%
   Dim rtt$
   Dim p$(100)
   Dim w$(20)
   Dim cf As Integer
   Dim cv As Integer
   Dim dj As Integer
   Dim scr, iq As Boolean
   Dim lc, tx, a, b, SFA As Integer
   Dim ww As Integer
   Dim x As Integer
   Dim rt2 As Integer
   Dim rt3 As Integer
   Dim rt1 As Integer
   Dim h As Integer
   Dim pc As Long
   Dim ptemp$
   Dim cvalue As Integer
   Dim pre_op$
   Dim c_check$
   Dim ik As Integer
   Dim hl$
   Dim cc As Integer
   Dim c0x$
   Dim c2x$
   Dim ISTHISINT As Integer
   Dim StartTimer As Long
   Dim gg As Integer
   Dim xcz As Integer
   Dim IMGQUAL As Integer
   Dim tr1$, tr2$, tr3$
   Dim ts1%, ts2%, ts3%
   Dim INSIDEME%
   Dim ww1%, ww2%, ww3%, ww4%, ww5%, ww6%, ww7%, ww8% ' Database Command Variables
   
   StartTimer = Timer
   
   Erase code$(), std$(), stn$(), ind(), inn$(), gs()
   
   For t = 0 To 10
    If file(t) <> 0 Then Close file(t): filea(t) = 0
   Next t
   
   Erase file(), filea()
   
   ' Load The Script File
   
   dx = FreeFile
   lc = 0
   Open filename For Input As #dx
   Do Until EOF(dx)
      Line Input #1, code$(lc)
      code$(lc) = ReplaceStr(code$(lc), vbTab, "   ")
      lc = lc + 1
   Loop
   Close dx
   
    IMGQUAL = 80

   'Debug.Print postdata$
   
   postdata$ = ReplaceStr(Trim$(ReplaceStr(postdata$, vbCrLf, " ")), "&&", "&")
   
   If Right$(postdata$, 1) <> "&" Then postdata$ = postdata$ & "&"
   a = Len(postdata$)
   For b = 1 To a
      d$ = Mid$(postdata$, b, 1)
      If d$ = "&" And tmp$ <> "" Then
         c = InStr(tmp$, "=")
         If InStr(tmp$, "=") Then
            CreateString RidFormatting(Left$(tmp$, c - 1)), RidFormatting(Right$(tmp$, Len(tmp$) - c))
         End If
         tmp$ = ""
      End If
      If d$ <> "&" Then tmp$ = tmp$ & d$
   Next b
   
   ' Create Strings For Script To Access About Request
   CreateString "referer", sx(sck).Referer
   CreateString "requestip", frmmain.ws(sck).RemoteHostIP
   CreateString "requestport", frmmain.ws(sck).RemotePort
   CreateString "localport", frmmain.ws(sck).LocalPort
   CreateString "localip", frmmain.ws(sck).LocalIP
   
   
   
   ' Begin Parsing
   
   pc = 0
   scr = True
   
   SFA = 0
   If code$(0) = "//" & SERVER_SECURITY_TAG1$ And code$(1) = "//" & SERVER_SECURITY_TAG2$ Then SFA = 1
   
   If SFA = 1 Then CreateString "request", sx(sck).Header
   
   For tt = 0 To lc
      If LTrim$(LCase$(code$(tt))) = "proc main" Then pc = tt + 1: Exit For
   Next tt
   
   Do Until scr = False
      
      DoEvents
      
      If (Timer - StartTimer) > (Longbow.TimeOut / 10) Then scr = False
      
      cl$ = LTrim$(code$(pc))
      
      If Left$(cl$, 2) <> "//" And cl$ <> "" Then
        
         ' Only parse if the line is valid
         Erase w$()
         v$ = LTrim$(code$(pc)): c2 = Len(v$): xc = 0: dd = 0: iq = False
         
        For c1 = 1 To c2
            f$ = Mid$(v$, c1, 1)
            If f$ = "[" And iq = False Then INSIDEME = 1: f$ = ""
            If f$ = "]" And iq = False Then INSIDEME = 0: f$ = ""
            If f$ = "(" And iq = False And INSIDEME = 0 Then dd = dd + 1
            If f$ = ")" And iq = False And INSIDEME = 0 Then dd = dd - 1
            If f$ = " " And iq = False And dd = 0 Then xc = xc + 1
            If f$ = "'" Then iq = Not iq
            If iq = False And dd = 0 Then w$(xc) = w$(xc) & Trim$(f$)
            If iq = False And dd <> 0 Then w$(xc) = w$(xc) & f$
            If iq = True Then w$(xc) = w$(xc) & f$
         Next c1
          
         cmd$ = LCase$(w$(0)): xc = xc + 1
               
               If IsInteger(cmd$) = 1 Then
                  xx$ = ""
                  ww = 0
                  For t = 1 To xc
                     tturn = 1
                     If w$(t) = "+" Or w$(t) = "-" Or w$(t) = "*" Or w$(t) = "/" Then
                        xx$ = w$(t): tturn = 0
                     End If

                     If tturn = 1 Then
                        If xx$ = "+" Then ww = ww + Val(ParseParam(w$(t)))
                        If xx$ = "-" Then ww = ww - Val(ParseParam(w$(t)))
                        If xx$ = "/" Then ww = ww / Val(ParseParam(w$(t)))
                        If xx$ = "*" Then ww = ww * Val(ParseParam(w$(t)))
                        If xx$ = "^" Then ww = ww ^ Val(ParseParam(w$(t)))
                     End If
                     If xx$ = "" Then ww = ww + Val(ParseParam(w$(t)))
                  Next t

                  SetInteger cmd$, ww
               End If
               
               If cmd$ = "dec" Then
                  If IsInteger(w$(1)) = 1 Then
                     f = Val(w$(2))
                     If f = 0 Then f = 1
                     SetInteger w$(1), GetInteger(w$(1)) - f
                  End If
               End If
               
               If cmd$ = "inc" Then
                  If IsInteger(w$(1)) = 1 Then
                     f = Val(w$(2))
                     If f = 0 Then f = 1
                     SetInteger w$(1), GetInteger(w$(1)) + f
                  End If
               End If
               
               If cmd$ = "strtoint" Then
                  If IsString(w$(1)) = 1 Then
                     If IsInteger(w$(2)) = 1 Then
                        SetInteger w$(2), Val(GetString(w$(1)))
                     End If
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "inttostr" Then
                  If IsInteger(w$(1)) = 1 Then
                     If IsString(w$(2)) = 1 Then
                        SetString w$(2), Trim$(Str$(GetInteger(w$(1))))
                     End If
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "fwrite" Then
                  If file(Val(w$(1))) <> 0 Then
                     Print #file(Val(w$(1))), ParseParam(w$(2))
                  End If
                  GoTo WIPEOUT
               End If
               
'DeleteEntry(databasefilename, fieldname, fielddata As String) As Integer
'GetNumOfDatabaseFields(databasefilename As String) As Integer
'GetFieldNames(databasefilename As String) As String
'GetFieldName(databasefilename As String, fieldnumber) As String
'GetEntry(databasefilename As String, entrynumber As Integer, fieldname As String) As String
'EntryExist(databasefilename As String, fieldname As String, fielddata As String, comparisontype As String) As Integer
'ReplaceEntry(databasefilename As String, findfieldname As String, findfielddata As String, findfieldcomparison As String, changefieldname As String, changefielddata As String) As Integer
'GetEntryNum(databasefilename As String, startent As Integer, fieldname As String, fielddata As String, fieldcomparison As String) As Long
'CreateDatabase(databasefilename As String, numfields As String, fieldnames As String) As Integer
'AddEntry(databasefilename As String, fielddata As String) As Long
               
               If cmd$ = "database" Then
                    ' w$(4) = databasefilename, constant for all subfunctions
                    ' w$(3) = variable
                    w$(2) = ParseParam(w$(2)): ww2 = Val(w$(2))
                    w$(3) = ParseParam(w$(3)): ww3 = Val(w$(3))
                    w$(4) = ParseParam(w$(4)): ww4 = Val(w$(4))
                    w$(5) = ParseParam(w$(5)): ww5 = Val(w$(5))
                    w$(6) = ParseParam(w$(6)): ww6 = Val(w$(6))
                    w$(7) = ParseParam(w$(7)): ww7 = Val(w$(7))
                    w$(8) = ParseParam(w$(8)): ww8 = Val(w$(8))

                    w$(3) = GetDirectory(filename$) & w$(3)
                
                
                If DB_EXTENDINFO = 1 Then
                    WXB sck, "<html><body><pre>"
                    WXB sck, "-- Database Request Extended Information --<br>"
                    WXB sck, "1" & w$(1) & "<br>"
                    WXB sck, "2" & w$(2) & "<br>"
                    WXB sck, "3" & w$(3) & "<br>"
                    WXB sck, "4" & w$(4) & "<br>"
                    WXB sck, "5" & w$(5) & "<br>"
                    WXB sck, "6" & w$(6) & "<br>"
                    WXB sck, "7" & w$(7) & "<br>"
                    WXB sck, "8" & w$(8) & "<br>"
                    WXB sck, "    ---------------</pre>"
                End If
                    ' w$(1) = database command
                    ' w$(2) = database file
                    
                    
                    
                    If InStr(w$(2), "..") Then
                        WXB sck, "Script security prevented this script from being executed"
                        sx(sck).Reqok = True
                        GoTo WIPEOUT
                    End If
                    
                    ' Big wedge of database code :D

                  If DB_EXTENDINFO = 1 Then Debug.Print w$(1)
                    
                    Select Case w$(1)
                        Case "deleteentry"
                            
                            SetInteger w$(2), DeleteEntry(w$(3), w$(4), w$(5))
                        Case "getnumofdbfields"
                            
                            SetInteger w$(2), GetNumOfDatabaseFields(w$(3))
                        Case "getfieldnames"

                            SetString w$(2), GetFieldNames(w$(3))
                        Case "getfieldname"

                            SetString w$(2), GetFieldName(w$(3), ww4)
                        Case "getentry"

                            SetString w$(2), GetEntry(w$(3), ww4, w$(5))
                        Case "entryexist"

                            SetInteger w$(2), EntryExist(w$(3), w$(4), w$(5), w$(6))
                        Case "replaceentry"

                            SetInteger w$(2), ReplaceEntry(w$(3), w$(4), w$(5), w$(6), w$(7), w$(8))
                        Case "getentrynum"

                            SetInteger w$(2), GetEntryNum(w$(3), ww4, w$(5), w$(6), w$(7))
                        Case "createdatabase"

                            SetInteger w$(2), CreateDatabase(w$(3), w$(4), w$(5))
                        Case "addentry"
                            'Debug.Print "w$(3)=" & w$(3)
                            'Debug.Print "w$(4)=" & w$(4)
                            SetInteger w$(2), AddEntry(w$(3), w$(4))
                    End Select
               End If
               
               
               If cmd$ = "fmread" Then
                  ' Multi Param Read
                  If file(Val(w$(1))) <> 0 Then
                     Line Input #file(Val(w$(1))), xtm$
                     Erase p$()
                     df = 0: cv = 1
                     dj = Len(xtm$)
                     For h = 1 To dj
                        hi$ = Mid$(xtm$, h, 1)
                        If hi$ = "," Then
                           hi$ = "": cv = cv + 1
                        End If
                           p$(cv) = p$(cv) + hi$

                     Next h
                    ' Debug.Print cv
                     For h = 1 To cv
                        'If p$(h) <> "" Then Debug.Print p$(h)
                        If w$(h + 1) <> "" Then
                           SetString w$(h + 1), p$(h)
                        End If
                     Next h
                  End If
               End If
               
               If cmd$ = "fread" Then
                  If file(Val(w$(1))) <> 0 Then
                     Input #file(Val(w$(1))), g$
                     If IsString(w$(2)) = 1 Then SetString w$(2), g$
                     If IsInteger(w$(2)) = 1 Then SetInteger w$(2), Val(g$)
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "file" Then
                  Select Case w$(1)
                     
                     Case "openappend"
                        'Debug.Print "OPENING FOR APPEND"
                        w$(2) = ParseParam(w$(2))
                        xx = Val(w$(3))
                        If file(xx) = 0 Then
                            file(xx) = FreeFile
                            If InStr(w$(2), "..") Or InStr(w$(2), Longbow.SecurityFile) Then
                                WXB sck, "Script Security Prevented This Script From Being Run"
                                sx(sck).Reqok = True
                                GoTo WIPEOUT
                            End If
                            Open GetDirectory(filename$) & w$(2) For Append As file(xx)
                        End If
                     
                     Case "openin"
                        w$(2) = ParseParam(w$(2))
                        xx = Val(w$(3))
                        If file(xx) = 0 Then
                           file(xx) = FreeFile
                           If InStr(w$(2), "..") Or InStr(w$(2), Longbow.SecurityFile) Then
                              WXB sck, "Script Security Prevented This Script From Being Run"
                              sx(sck).Reqok = True
                              GoTo WIPEOUT
                           End If
                           If Exists(GetDirectory(filename$) & w$(2)) = 1 Then
                              Open GetDirectory(filename$) & w$(2) For Input As file(xx)
                           Else
                              file(xx) = 0: filea(xx) = 0
                           End If
                        End If
                     
                     Case "openout"
                        w$(2) = ParseParam(w$(2))
                        xx = Val(w$(3))
                        If file(xx) = 0 Then
                           file(xx) = FreeFile
                           If InStr(w$(2), "..") Or InStr(w$(2), Longbow.SecurityFile) Then
                              WXB sck, "Script Security Prevented This Script From Being Run"
                              sx(sck).Reqok = True
                              GoTo WIPEOUT
                           End If
                           Open GetDirectory(filename$) & w$(2) For Output As file(xx)
                        End If
                     
                     Case "close"
                        If file(Val(w$(2))) <> 0 Then
                           Close file(Val(w$(2)))
                           file(Val(w$(2))) = 0
                           filea(Val(w$(2))) = 0
                        End If
                     
                     Case "kill"
                        If Exists(GetDirectory(filename$) & w$(2)) = 1 Then
                           Kill GetDirectory(filename$) & w$(2)
                        End If
                     
                     Case "copy"
                        If Exists(GetDirectory(filename$) & w$(2)) = 1 Then
                           SecureCopy GetDirectory(filename$) & w$(2), GetDirectory(filename$) & w$(3)
                        End If
                  End Select
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "image" Then
                    If InStr(w$(2), "..") Then
                        WXB sck, "Script Security Error"
                        sx(sck).Reqok = True
                        GoTo WIPEOUT
                    End If

                    Select Case w$(1)
                        
                        Case "font"
                            rtt$ = ParseParam(w$(2))
                            fsu.pic(sck).Font = rtt$
                        Case "fontbold"
                            fsu.pic(sck).FontBold = True
                        Case "fontnobold"
                            fsu.pic(sck).FontBold = False
                        Case "fontitalic"
                            fsu.pic(sck).FontItalic = True
                        Case "fontnoitalic"
                            fsu.pic(sck).FontItalic = False
                        Case "fontunderline"
                            fsu.pic(sck).FontUnderline = True
                        Case "fontnounderline"
                            fsu.pic(sck).FontUnderline = False
                        Case "fontstrikethru"
                            fsu.pic(sck).FontStrikethru = True
                        Case "fontnostrikethru"
                            fsu.pic(sck).FontStrikethru = False
                        Case "fontsize"
                            rs1 = Val(ParseParam(w$(2)))
                            fsu.pic(sck).FontSize = rs1
                        Case "currentx"
                            rs1 = Val(ParseParam(w$(2)))
                            fsu.pic(sck).CurrentX = rs1
                        Case "currenty"
                            rs1 = Val(ParseParam(w$(2)))
                            fsu.pic(sck).CurrentY = rs1
                        Case "pset"
                            fsu.pic(sck).PSet (Val(ParseParam(w$(2))), Val(ParseParam(w$(3))))
                        Case "line"
                            fsu.pic(sck).Line (Val(ParseParam(w$(2))), Val(ParseParam(w$(3))))-(Val(ParseParam(w$(4))), Val(ParseParam(w$(5))))
                        Case "circle"
                            fsu.pic(sck).Circle (Val(ParseParam(w$(2))), Val(ParseParam(w$(2)))), Val(ParseParam(w$(2)))
                        Case "box"
                            fsu.pic(sck).Line (Val(ParseParam(w$(2))), Val(ParseParam(w$(3))))-(Val(ParseParam(w$(4))), Val(ParseParam(w$(5)))), , B
                        Case "rect"
                            fsu.pic(sck).Line (Val(ParseParam(w$(2))), Val(ParseParam(w$(3))))-(Val(ParseParam(w$(4))), Val(ParseParam(w$(5)))), , BF
                        Case "write"
                            fsu.Print ParseParam(w$(2))
                        Case "copyrect"
                            DESTX = Val(ParseParam(w$(2)))
                            DESTY = Val(ParseParam(w$(3)))
                            SRCX = Val(ParseParam(w$(4)))
                            SRCY = Val(ParseParam(w$(5)))
                            RECTWIDTH = Val(ParseParam(w$(6)))
                            RECTHEIGHT = Val(ParseParam(w$(7)))
                            m_imutils.BitBlt fsu.pic(sck).hdc, DESTX, DESTY, RECTWIDTH, RECTHEIGHT, fsu.pic(sck).hdc, SRCX, SRCY, SRCCOPY
                            fsu.pic(sck).Refresh
                        Case "width"
                            rt1 = Val(ParseParam(w$(2)))
                            If rt1 > 640 Then rt1 = 640
                            fsu.pic(sck).ScaleWidth = rt1
                        Case "height"
                            rt1 = Val(ParseParam(w$(2)))
                            If rt1 > 640 Then rt! = 480
                            fsu.pic(sck).ScaleHeight = rt1
                        Case "back"
                            rt1 = Val(ParseParam(w$(2)))
                            fsu.pic(sck).BackColor = rt1
                        Case "clear"
                            fsu.pic(sck).Cls
                        Case "fore"
                            rt1 = Val(ParseParam(w$(2)))
                            fsu.pic(sck).ForeColor = rt1
                        Case "penwidth"
                            rt1 = Val(ParseParam(w$(2)))
                            fsu.pic(sck).DrawWidth = rt1
                        Case "load"
                            fsu.pic(sck).Picture = LoadPicture(GetDirectory(filename) & ParseParam(w$(2)))
                        Case "quality"
                            IMGQUAL = Val(ParseParam(w$(2)))
                        Case "save"
                            SavePicture fsu.pic(sck).Picture, "c:\r" & Trim$(Str$(sck)) & ".bmp"
                            m_imutils.ConvertBMPtoJPG "c:\r" & Trim$(Str$(sck)) & ".bmp", GetDirectory(filename) & ParseParam(w$(2)), True, IMGQUAL, False
                        Case "send"
                            'Debug.Print IMGQUAL
                            SavePicture fsu.pic(sck).Picture, "c:\r" & Trim$(Str$(sck)) & ".bmp"
                            m_imutils.ConvertBMPtoJPG "c:\r" & Trim$(Str$(sck)) & ".bmp", "c:\r" & Trim$(Str$(sck)) & ".jpg", True, IMGQUAL, False
                            WX_FILE sck, "c:\r" & Trim$(Str$(sck)) & ".jpg"
                            Kill "c:\r" & Trim$(Str$(sck)) & ".jpg"
                    End Select
                End If
               
               
               If cmd$ = "debug" Then
                  Debug.Print w$(1)
               End If
               
               If IsString(cmd$) = 1 Then
                  we$ = ""
                  For t = 1 To xc
                     we$ = we$ & ParseParam(w$(t))
                  Next t
                  SetString cmd$, we$
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "proc" Then
                  For t = pc To lc
                     If LTrim$(Left$(code$(t), 7)) = "endproc" Then pc = t + 1: Exit For
                  Next t
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "call" Or cmd$ = "gosub" Then
                  For t = 0 To 10
                     If gs(t) = 0 Then
                        For gg = 0 To lc
                           'Debug.Print ">" & w$(1) & "<"
                           If LTrim$(code$(gg)) = "proc " & w$(1) Then
                              'Debug.Print "FOUND PROC"
                              gs(t) = pc: pc = gg
                              GoTo DONE_CALL_STACK
                           End If
                        Next gg
                     End If
                  Next t
DONE_CALL_STACK:
               GoTo WIPEOUT
               End If
               
               If cmd$ = "endproc" Then
                  For t = 10 To 0 Step -1
                     If gs(t) <> 0 Then
                        pc = gs(t)
                        gs(t) = 0
                        Exit For
                     End If
                  Next t
               GoTo WIPEOUT
               End If

               If cmd$ = "newint" Then
                  w$(2) = ParseParam(w$(2))
                  x = Val(w$(2))
                  CreateInteger w$(1), x
               End If
               
               If cmd$ = "redefint" Then
                  If IsInteger(w$(1)) = 0 Then
                     CreateInteger w$(1), Val(ParseParam(w$(2)))
                  Else
                     If w$(2) <> "" Then
                        SetInteger w$(1), Val(ParseParam(w$(2)))
                     End If
                  End If
                  GoTo WIPEOUT
               End If

               If cmd$ = "redefstr" Then
                  If IsString(w$(1)) = 0 Then
                     CreateString w$(1), ParseParam(w$(2))
                  Else
                     If w$(2) <> "" Then
                      SetString w$(1), ParseParam(w$(2))
                     End If
                  End If
               GoTo WIPEOUT
               End If
               
               If cmd$ = "strcopy" Then
                  ' strcopy dest src
                  SetString w$(1), ParseParam(w$(2))
               GoTo WIPEOUT
               End If
                  
               If cmd$ = "newstr" Then
                  w$(2) = ParseParam(w$(2))
                  CreateString w$(1), w$(2)
               GoTo WIPEOUT
               End If
                  
               If cmd$ = "includefile" Then
                  w$(1) = ParseParam(w$(1))
                  If Exists(GetDirectory(filename$) & w$(1)) Then
                     w$(1) = GetDirectory(filename$) & w$(1)
                  End If
                  
                  If Exists(w$(1)) = 0 Then
                     WXB sck, "SCRIPT ERROR: File Not Found On Line " & Trim$(Str$(pc))
                     sx(sck).Reqok = True
                     GoTo WIPEOUT
                  End If
                  g$ = GetDirectory(w$(1))
                  If Exists(g$ & Longbow.SecurityFile) = 0 Then
                     WXB sck, "SCRIPT ERROR: No Directory Access " & Trim$(Str$(pc))
                     sx(sck).Reqok = True
                     GoTo WIPEOUT
                  End If
                  DIR_READ = 1
                  Open g$ & Longbow.SecurityFile For Input As #4
                  Do Until EOF(4)
                     Line Input #4, f$
                     If LCase$(f$) = "read=no" Then DIR_READ = 0
                     If LCase$(f$) = "secure=yes" Then DIR_READ = 0
                  Loop
                  Close 4
                  If SFA = 1 Then DIR_READ = 1
                  If DIR_READ = 0 Then
                     WXB sck, "SCRIPT ERROR: Access Denied On Line " & Trim$(Str$(pc))
                     sx(sck).Reqok = True
                     GoTo WIPEOUT
                  End If
                  Open w$(1) For Binary As #33
                     ff = LOF(33)
                     jh$ = Space$(ff)
                     Get #33, , jh$
                  Close 33
                  WXB sck, jh$
                  jh$ = ""
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "goto" Then
                  For t = 0 To lc
                     If LTrim$(code$(t)) = w$(1) & ">" Then pc = t: Exit For
                  Next t
               GoTo WIPEOUT
               End If

               If cmd$ = "delint" Then
                  DeleteInteger w$(1)
                  GoTo WIPEOUT
               End If
               
               
               If cmd$ = "delstr" Then
                  DeleteString w$(1)
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "else" Then
                  nebo = 0
                  For t = pc To lc
                     
                     If Left$(LTrim$(code$(t)), 2) = "if" Then nebo = nebo + 1
                     If Left$(LTrim$(code$(t)), 5) = "endif" Then nebo = nebo - 1
                     'Debug.Print nebo
                     If Left$(LTrim$(code$(t)), 5) = "endif" And nebo = -1 Then pc = t: Exit For
                  Next t
               GoTo WIPEOUT
               End If
               
               If cmd$ = "end" Then
                    scr = False
                End If
               
               If cmd$ = "print" Then
                  ptemp$ = ""
                  For t = 1 To xc
                     ptemp$ = ptemp$ & ParseParam(w$(t))
                  Next t
                  WXB sck, ptemp$
                  ptemp$ = ""
               GoTo WIPEOUT
               End If

               If cmd$ = "writehttp" Then
                    WriteHTTP sck, Val(ParseParam(w$(1))), ParseParam(w$(2))
               End If
               
               If cmd$ = "if" Then
                    'Debug.Print "IF FUNCTION CALLED"
                  
                  cvalue = 0
                  pre_op$ = ""
                  total_if_params = xc
                    'Debug.Print xc
                  For if_loop = 1 To xc
                     c_check$ = w$(if_loop)
                     'Debug.Print c_check$
                     If c_check$ = "&&" Then pre_op$ = "&&"
                     If c_check$ = "||" Then pre_op$ = "||"
                     If Left$(c_check$, 1) = "(" Then
                        ' If Segment
                        '(name$(hello) == "PASS$")
                        c_check$ = Mid$(c_check$, 2, Len(c_check$) - 2)
                        'name$(hello) == "PASS$"
                        s1 = Len(c_check$)
                        Erase p$()
                        iq = False: cc = 0
                        For tt = 1 To s1
                           hl$ = Mid$(c_check$, tt, 1)
                           
                           If hl$ = "(" And iq = False Then ik = ik + 1
                           If hl$ = ")" And iq = False Then ik = ik - 1
                                                      
                           If hl$ = " " And iq = False And ik = 0 Then cc = cc + 1
                           
                           If hl$ = "'" Then iq = Not iq
                                                      
                           If iq = False And ik = 0 Then p$(cc) = p$(cc) & Trim$(hl$)
                           If iq = False And ik <> 0 Then p$(cc) = p$(cc) & hl$
                           
                           If iq = True Then p$(cc) = p$(cc) & hl$
                        Next tt
                        c0x$ = p$(0)
                        c2x$ = p$(2)
                        p$(0) = ParseParam(p$(0))
                        p$(2) = ParseParam(p$(2))
                        kk = 0
                        If (p$(0) <> "" And p$(2) <> "") And Val(p$(0)) = Val(Str$(Val(p$(0)))) Then
                            ISTHISINT = 1
                        Else
                            ISTHISINT = 0
                        End If
                        
                        If IsString(c0x$) = 1 Then ISTHISINT = 0
                        If IsString(c2x$) = 1 Then ISTHISINT = 0
                        If IsInteger(c0x$) = 1 Then ISTHISINT = 1
                        If IsInteger(c2x$) = 1 Then ISTHISINT = 1
                        
                        Select Case ISTHISINT
                                
                                Case 0
                                    Select Case p$(1)
                                       Case ">"
                                          If p$(0) > p$(2) Then kk = 1
                                       Case "<"
                                          If p$(0) < p$(2) Then kk = 1
                                       Case "=="
                                          If p$(0) = p$(2) Then kk = 1
                                       Case "!="
                                          If p$(0) <> p$(2) Then kk = 1
                                       Case ">="
                                          If p$(0) >= p$(2) Then kk = 1
                                       Case "<="
                                          If p$(0) <= p$(2) Then kk = 1
                                    End Select
                                
                                Case 1
                                        p1 = Val(p$(0))
                                        p2 = Val(p$(2))
                                    Select Case p$(1)
                                       Case ">"
                                          If p1 > p2 Then kk = 1
                                       Case "<"
                                          If p1 < p2 Then kk = 1
                                       Case "=="
                                          If p1 = p2 Then kk = 1
                                       Case "!="
                                          If p1 <> p2 Then kk = 1
                                       Case ">="
                                          If p1 >= p2 Then kk = 1
                                       Case "<="
                                          If p1 <= p2 Then kk = 1
                                    End Select
                        End Select

                       ' Debug.Print p$(0)
                       ' Debug.Print p$(2)
                       ' Debug.Print kk
                        
                        If kk = 0 Then
                           If pre_op$ = "||" Or pre_op$ = "OR" Then
                              If cvalue = 1 Then cvalue = 1
                              If cvalue = 0 Then cvalue = 0
                           End If
                           If pre_op$ = "&&" Or pre_op$ = "AND" Then
                              If cvalue = 0 Then cvalue = 0
                              If cvalue = 1 Then cvalue = 0
                           End If
                        End If
                        If kk = 1 Then
                           If pre_op$ = "||" Or pre_op$ = "OR" Then
                              If cvalue = 0 Then cvalue = 1
                              If cvalue = 1 Then cvalue = 1
                           End If
                           If pre_op$ = "&&" Or pre_op$ = "AND" Then
                              If cvalue = 1 Then cvalue = 1
                              If cvalue = 0 Then cvalue = 0
                           End If
                        End If
                        
                       ' Debug.Print kk
                        
                        If pre_op$ <> "&&" And pre_op$ <> "||" And pre_op$ <> "AND" And pre_op$ <> "OR" Then
                              cvalue = kk
                        End If
                              
                     End If
                  Next if_loop
                  
                  it = cvalue
                  If it = 0 Then
                     cando = 0
                     For t = pc + 1 To lc
                     ' THIS NEXT 2 LINES ARE A TEST SECTION
                        If cando < 1 And Left$(LTrim$(code$(t)), 4) = "else" Then pc = t: Exit For
                        If cando < 1 And Left$(LTrim$(code$(t)), 5) = "endif" Then pc = t:  Exit For
                        
                        If Left$(LTrim$(code$(t)), 2) = "if" Then cando = cando + 1
                        If Left$(LTrim$(code$(t)), 5) = "endif" Then cando = cando - 1
                        'Debug.Print Left$(LTrim$(code$(t)), 5); cando
                        
'                        If cando < 1 And Left$(LTrim$(code$(t)), 4) = "else" Then pc = t: Exit For
'                        If cando < 1 And Left$(LTrim$(code$(t)), 5) = "endif" Then pc = t:  Exit For
                     Next t
                  End If
                  'WXB sck, "I HAD BEEN CHECKED"
               GoTo WIPEOUT
               
               End If
    
        ' ** SECURE COMMANDS
        
        If SFA = 1 Then


        End If

WIPEOUT:
      End If
                  
SkipMe:
      
      pc = pc + 1
      If pc > lc Then scr = False
   Loop
      
      If sx(sck).Buffer = "" Then sx(sck).Buffer = vbCrLf & vbCrLf
      ScriptTidy
   
   Exit Sub
   
Execute_Error:
   Debug.Print "Script Error"
   WXB sck, "Error Occured In Script Processing On Line " & Trim$(Str$(pc)) & " ErrCode " & Err.Description
   ScriptTidy
   sx(sck).Reqok = True
End Sub

Public Function ParseParam(para As String) As String
   'Debug.Print "PARA=" & para$
   If para$ = "&" Or para$ = "" Or para$ = " &" Then Exit Function
   Dim d$(200)
   Dim iq As Boolean
   If para$ = "lbCrLf" Then ParseParam = vbCrLf: Exit Function
   
   ' Is String?
   para$ = Trim$(para$)
   
   If IsString(para$) = 1 Then
      ParseParam = GetString(para$)
      Exit Function
   End If
   
   ' Is A Number?
   If IsNumeric(para$) Then
      ParseParam = para$
      Exit Function
   End If
   
   ' Is Bracketed?
   '(3 * 10)
   If Left$(para$, 1) = "(" And Right$(para$, 1) = ")" Then
      ' This parameter is bracketed, we need to resolve whatevers inside it
      f1$ = Mid$(para$, 2, Len(para$) - 2)
      '3 * 10
      
      ' Now we parse the whole thing to get the sub params
      r = Len(f1$)
      jk = 0: wc = 0
      Erase d$()
      For g = 1 To r
         k$ = Mid$(f1$, g, 1)
         If k$ = "'" Then iq = Not iq
         If k$ = "(" And iq = False Then jk = jk + 1
         If k$ = ")" And iq = False Then jk = jk - 1
         If k$ = " " And iq = False And jk = 0 Then wc = wc + 1
            d$(wc) = d$(wc) & k$
      Next g
      
      For g = 0 To wc
         welly = 1
         d$(g) = Trim$(d$(g))
         If d$(g) = "+" Or d$(g) = "-" Or d$(g) = "/" Or d$(g) = "*" Then
            xc$ = d$(g): welly = 0
         End If
         If welly = 1 Then
            If xc$ = "+" Then tempo = tempo + Val(ParseParam(d$(g)))
            If xc$ = "-" Then tempo = tempo - Val(ParseParam(d$(g)))
            If xc$ = "/" Then tempo = tempo / Val(ParseParam(d$(g)))
            If xc$ = "*" Then tempo = tempo * Val(ParseParam(d$(g)))
         End If
         If xc$ = "" Then tempo = Val(ParseParam(d$(g)))
      Next g
      ParseParam = Trim$(Str$(tempo))
      Exit Function
   End If
   
   
   ' Is Integer?
   If IsInteger(para$) = 1 Then
      ParseParam = Trim$(Str$(GetInteger(para$)))
      Exit Function
   End If
      
   
   ' Is Server Const?
   If GetConst(para$) <> "" Then
      ParseParam = GetConst(para$)
      Exit Function
   End If
   
   ' Is Normal Text?
   dx = Len(para$)
   iq = False: jk = 0
   For t = 1 To dx
      If Mid$(para$, t, 1) = "(" And iq = False Then jk = jk + 1
      If Mid$(para$, t, 1) = ")" And iq = False Then jk = jk - 1
      If Mid$(para$, t, 1) = "'" Then iq = Not iq
      If Mid$(para$, t, 1) = "&" And iq = False And jk = 0 Then MULTI_PARAM = 1: GoTo MULTI
   Next t
   GoTo NOMULTI
MULTI:
   'Debug.Print "this para$=" & para$ & "="
   Erase d$()
   wc = 0
   iq = False: jk = 0
   If MULTI_PARAM = 1 Then
      'Debug.Print "MULTIPARAM"
      ' MultiParameter
      r = Len(para$)
      For g = 1 To r
         k$ = Mid$(para$, g, 1)
         If k$ = "'" Then iq = Not iq
         If k$ = "(" And iq = False Then jk = jk + 1
         If k$ = ")" And iq = False Then jk = jk - 1
         If k$ = " " And iq = False And jk = 0 Then wc = wc + 1
          d$(wc) = d$(wc) & k$
      Next g
   
      For g = 0 To wc
         
         d$(g) = Trim$(d$(g))
         'Debug.Print "tempo$=" & d$(g) & "="
         tempoo$ = tempoo$ & ParseParam(d$(g))
      Next g
      ParseParam = tempoo$
      Exit Function
   End If
   
   
NOMULTI:
   
   
   If Left$(para$, 1) = "'" Then
      ParseParam = ReplaceStr(Mid$(para$, 2, Len(para$) - 2), "\n", vbCrLf)
      Exit Function
   End If
   
   dd = Len(para$)
   For cc = 1 To dd
      f$ = Mid$(para$, cc, 1)
      If f$ = "'" Then Exit Function
      If f$ <> "(" Then GoTo BAILMEOUT
   Next cc
   
   Exit Function
   
BAILMEOUT:
   ' ucase(lcase(forname) & ' ' & lcase(surname))
   '      a                                    cc
   
   a = InStr(para$, "(")
   b = Len(para$)
   For c = b To 1 Step -1
      If Mid$(para$, c, 1) = ")" Then x = x + 1
      If Mid$(para$, c, 1) <> ")" Then GoTo ALLOUT
      
   Next c
   
ALLOUT:
   
   x = x - 1
   
   z$ = Mid$(para$, a + 1, Len(para$) - a - x)
   
   'Debug.Print "z$=" & z$
   'ucase(lcase(ucase('name')))
   'ucase(lcase('NAME'))
   'ucase('name')
   
   If Left$(para$, 3) = "int" Then
        ParseParam = Str$(Int(Val(LCase(ParseParam(Mid(para$, 5, Len(para$) - 5))))))
        Exit Function
    End If
   
   If Left$(para$, 5) = "lcase" Then
      ParseParam = LCase(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
      Exit Function
   End If
   
   If Left$(para$, 5) = "ucase" Then
      ParseParam = UCase(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
      Exit Function
   End If
   
    If Left$(para$, 3) = "len" Then
        ParseParam = Trim$(Str$(Len(ParseParam(GetParam(para$, 0)))))
        Exit Function
    End If
        
    If Left$(para$, 5) = "ltrim" Then
        ParseParam = LTrim(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
        Exit Function
    End If
    
    If Left$(para$, 5) = "rtrim" Then
        ParseParam = RTrim(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
        Exit Function
    End If
    
    If Left$(para$, 4) = "trim" Then
        ParseParam = Trim(ParseParam(Mid$(para$, 6, Len(para$) - 6)))
    End If
   
   If Left$(para$, 3) = "rnd" Then
      d1 = Val(ParseParam(GetParam(para$, 0)))
      d2 = Val(ParseParam(GetParam(para$, 1)))
      j = (Rnd(1) * (d2 - d1)) + d1
      ParseParam = Trim$(Str$(j))
      Exit Function
   End If
   
   If Left$(para$, 4) = "left" Then
        r1$ = ParseParam(GetParam(para$, 0))
        d1 = Val(ParseParam(GetParam(para$, 1)))
        ParseParam = Left$(r1$, d1)
        Exit Function
    End If

    If Left$(para$, 5) = "right" Then
        r1$ = ParseParam(GetParam(para$, 0))
        d1 = Val(ParseParam(GetParam(para$, 1)))
        ParseParam = Right$(r1$, d1)
        Exit Function
    End If
    
    If Left$(para$, 9) = "uridecode" Then
        r1$ = ParseParam(GetParam(para$, 0))
        ParseParam = RidFormatting(r1$)
    End If
    
    If Left$(para$, 7) = "replstr" Then
        ' replstr stringname, find, replace
        r1$ = ParseParam(GetParam(para$, 0))
        r2$ = ParseParam(GetParam(para$, 1))
        r3$ = ParseParam(GetParam(para$, 2))
        temp1$ = ReplaceStr(r1$, r2$, r3$)
        ParseParam = temp1$
    End If
        
   If Left$(para$, 3) = "mid" Then
      rl$ = ParseParam(GetParam(para$, 0))
      d1 = Val(ParseParam(GetParam(para$, 1)))
      d2 = Val(ParseParam(GetParam(para$, 2)))
      ParseParam = Mid$(rl$, d1, d2)
      Exit Function
   End If
   
    If Left$(para$, 6) = "revstr" Then
        r1$ = ParseParam(GetParam(para$, 0))
        ParseParam = RevStr(r1$)
        Exit Function
    End If
        
    If Left$(para$, 3) = "chr" Then
        d1 = Val(ParseParam(GetParam(para$, 0)))
        ParseParam = Chr$(d1)
    End If
   
    If Left$(para$, 3) = "asc" Then
        ds2$ = ParseParam(GetParam(para$, 0))
        ParseParam = Trim$(Str$(Asc(ds2$)))
    End If
    
    If Left$(para$, 3) = "eof" Then
        ds2$ = ParseParam(GetParam(para$, 0))
        ff$ = Trim$(Str$(EOF(file(Val(ds2$)))))
        If ff$ = "True" Then ff$ = "-1"
        If ff$ = "False" Then ff$ = "0"
        ParseParam = ff$
    End If
    
    If Left$(para$, 5) = "instr" Then
        ds2$ = "": ds3$ = "": ds4$ = ""
        ds2$ = ParseParam(GetParam(para$, 0))
        ds3$ = ParseParam(GetParam(para$, 1))
        ds4$ = ParseParam(GetParam(para$, 2))
        cd234$ = Trim$(Str$(InStr(Val(ds2$), ds3$, ds4$)))
        ParseParam = cd234$
    End If
        
    If Left$(para$, 9) = "b64decode" Then
        ds2$ = ""
        ds2$ = ParseParam(GetParam(para$, 0))
        ParseParam = B64.Decode(ds2$)
        Exit Function
    End If
    
    If Left$(para$, 9) = "b64encode" Then
        ds2$ = ""
        ds2$ = ParseParam(GetParam(para$, 0))
        ParseParam = B64.Encode(ds2$)
        Exit Function
    End If
    
    If Left$(para$, 10) = "is2edecode" Then
        ds2$ = "": ds3$ = ""
        ds2$ = ParseParam(GetParam(para$, 0))
        ds3$ = ParseParam(GetParam(para$, 1))
        ParseParam = m_is2e.Decrypt(ds2$, ds3$)
        Exit Function
    End If
    
    If Left$(para$, 10) = "is2eencode" Then
        ds2$ = "": ds3$ = ""
        ds2$ = ParseParam(GetParam(para$, 0))
        ds3$ = ParseParam(GetParam(para$, 1))
        ParseParam = m_is2e.Encrypt(ds2$, ds3$)
        Exit Function
    End If
    
    If Left$(para$, 9) = "preformat" Then
        ds2$ = "": ds3$ = ""
        ds2$ = ParseParam(GetParam(para$, 0))
        ParseParam = UnRidFormatting(ds2$)
        Exit Function
    End If
    
    Exit Function
ParseParamErrCheck:
    WXB CURRENTSOCK, "Scripting Error - " & Err.Description
    sx(sck).Reqok = True
End Function


Public Function IsInteger(pint_name As String) As Integer
   Dim t As Integer
   For t = 0 To toti
      If pint_name$ = inn$(t) & "%" Then IsInteger = 1: Exit Function
   Next t
   IsInteger = 0
End Function

Public Sub CreateInteger(int_name As String, Optional int_data As Integer)
   Dim t As Integer
   For t = 0 To 1000
      If inn$(t) = "" Then inn$(t) = int_name$: ind(t) = int_data: toti = t + 1: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function GetInteger(int_name As String) As Integer
   Dim t As Integer
   For t = 0 To toti
      If inn$(t) & "%" = int_name$ Then GetInteger = ind(t): Exit Function
      If t = 500 Then DoEvents
   Next t
End Function

Public Sub DeleteInteger(int_name As String)
   Dim t As Integer
   For t = 0 To toti
      If inn$(t) & "%" = int_name$ Then inn$(t) = "": ind(t) = 0: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Sub SetInteger(int_name As String, int_data As Integer)
   Dim t As Integer
   For t = 0 To toti
      If inn$(t) & "%" = int_name$ Then ind(t) = int_data: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function IsString(pstring_name As String) As Integer
   Dim t As Integer
   For t = 0 To tots
      If pstring_name$ = stn$(t) & "$" Then IsString = 1: Exit Function
   Next t
   IsString = 0
End Function



Public Sub CreateString(string_name As String, Optional string_data As String)
   Dim t As Integer
   If Right$(string_name$, 1) = "$" Then string_name$ = Left$(string_name$, Len(string_name$) - 1)
   For t = 0 To 1000
      If stn$(t) = "" Then stn$(t) = string_name: std$(t) = string_data: tots = t + 1: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function GetString(string_name As String) As String
   For t = 0 To tots
      If stn$(t) & "$" = string_name$ Then GetString = std$(t): Exit Function
      If t = 500 Then DoEvents
   Next t
End Function

Public Sub DeleteString(string_name As String)
   For t = 0 To tots
      If stn$(t) & "$" = string_name$ Then stn$(t) = "": std$(t) = "": Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Sub SetString(string_name As String, string_data As String)
   'Debug.Print "SetString(" & string_name$ & "," & string_data$ & ")"
   For t = 0 To tots
      If stn$(t) & "$" = string_name$ Then std$(t) = string_data$: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function GetConst(poss_const As String) As String
   Select Case poss_const
      Case "#time"
         r$ = Time$
      Case "#date"
         r$ = Date$
      Case "#timer"
         r$ = Trim$(Str$(Timer))
      Case "#svname"
         r$ = Longbow.ServerName
      Case "#svadmin"
         r$ = Longbow.ServerAdmin
   End Select
   GetConst = r$
End Function

Public Function GetParam(ByVal comd As String, paramnum As Integer) As String
Dim d$(200)
Dim iq As Boolean
Dim dd As Integer
Dim a As Integer
Dim b As Integer
Dim c$
Dim ij As Integer
Dim wc As Integer

dd = InStr(comd$, "(")
comd$ = Right$(comd$, Len(comd$) - dd)
comd$ = Left$(comd$, Len(comd$) - 1)
a = Len(comd$)
iq = False: ij = 0
For b = 1 To a
   c$ = Mid$(comd$, b, 1)
   If c$ = "(" And iq = False Then ij = ij + 1
   If c$ = ")" And iq = False Then ij = ij - 1
   If c$ = "'" Then iq = Not iq
   If c$ = "," And iq = False And ij = 0 Then wc = wc + 1: c$ = ""
   d$(wc) = d$(wc) & c$
Next b
GetParam = d$(paramnum)
End Function

Public Sub SecureCopy(src As String, dest As String)
   On Error GoTo SECCOPYERROR
   FileCopy src, dest
   Exit Sub
SECCOPYERROR:
End Sub

Public Sub ScriptTidy()
    Dim t As Integer
    ' Tidy up the scripting resources when the script ends

    ' Close all open files
    For t = 0 To 10
        If file(t) <> 0 Then Close file(t): filea(t) = 0
    Next t
    
    ' Clear the picture box
        fsu.pic(CURRENTSOCK).Cls
        fsu.di(CURRENTSOCK).Path = Longbow.DefaultRoot
        fsu.fi(CURRENTSOCK).Path = Longbow.DefaultRoot
    
End Sub
