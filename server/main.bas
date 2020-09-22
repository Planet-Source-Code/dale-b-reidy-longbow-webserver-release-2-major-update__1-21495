Attribute VB_Name = "m_main"
Public Type SVMIME
   ext As String
   mtype As String
End Type

Public Type SVVDIR
   virt As String
   real As String
   acti As String
End Type

Public Type SVHOST
   svr As String
   root As String
   acti As String
End Type

Public Type SVDATA
   ServerName As String       ' The server name
   ServerAdmin As String      ' Server admins name
   ListenPort As Integer      ' Listen port for server
   MaxSocks As Integer        ' Maximum sockets
   DefaultRoot As String      ' Default server root
   DocLoc As String           ' Doc file directory
   LogLoc As String           ' Log file directory
   logtype As Integer         ' 0=NONE 1=FULL 2=ERRORS 3=REQUESTS
   IndexFile As String        ' Filename
   SecurityFile As String     ' Filename
   DirListing As Integer      ' 0=NONE 1=SIMPLE 2=GRAPHICAL
   TimerUpdate As Integer     ' Milliseconds For Timer Update
   TimeOut As Integer         ' Seconds To Timeout

End Type

Public Type SVUSERS
   username As String
   password As String
   Active As String           ' 0=NO 1=YES
   Directory As String        ' Users directory
End Type

Public Type SVSOCKET
   Buffer As String
   Header As String
   Reqok As Boolean
   TimeAlive As Long
   Referer As String

End Type

Public Const LONGBOW_SERVER_DETAILS = "LongBow Server 1.0 By Dale Reidy"

Public ServerLogFile As String      ' Server Log File Name

Public ServerStartTime As Long      ' Timer Value from when server was started

Public Longbow As SVDATA

Public B64 As New Base64

Public ipban(200) As String
Public mimes(200) As SVMIME
Public users(2000) As SVUSERS
Public vdirz(60) As SVVDIR
Public vhost(60) As SVHOST

Public SHARES(700) As String ' Shared Directories
Public SHARESL(700) As String ' Shared Directory Rootage (TOPLEVEL|<headed directory>)

Public SERVER_SECURITY_TAG1$, SERVER_SECURITY_TAG2$ ' Server scripting security tags

Public NumReq As Long               ' Number of requests

Public sx() As SVSOCKET
Public cx() As New script

Public Sub InitServer()
   ServerStartTime = Timer
   
'   LoadMimes
'   LoadUsers
'   LoadVDirz
'   LoadHosts
   
   LoadServerDetails
   
   For t = 1 To Longbow.MaxSocks
      
      Load frmmain.ws(t)
      Load fsu.pic(t)
      Load fsu.fi(t)
      Load fsu.di(t)
   
   Next t
   
   ReDim sx(Longbow.MaxSocks)
   ReDim cx(Longbow.MaxSocks)
   

   frmmain.ws(0).LocalPort = Longbow.ListenPort
   frmmain.sxt.Interval = Longbow.TimerUpdate
   StartServer

End Sub

Public Sub StartServer()
If frmmain.ws(0).State <> sckClosed Then frmmain.ws(0).Close
   frmmain.ws(0).Listen
   frmmain.Command1.Enabled = True
   frmmain.Command2.Enabled = False

End Sub

Public Sub StopServer()
   frmmain.ws(0).Close
   frmmain.Command2.Enabled = True
   frmmain.Command1.Enabled = False
End Sub

Public Sub LoadServerDetails()
'   On Error GoTo LOADSERVERERR
    On Error Resume Next
   
   Dim xd As Long
   Dim b$, a$, x$
   Dim d As Long
   
   ' Load the directory viewing color scheme and font face data
   LoadDirViewColorScheme
   
   'Public SERVER_SECURITY_TAG1$, SERVER_SECURITY_TAG2$ ' Server scripting security tags

   Open "..\conf\scriptsec.cfg" For Input As #1
      Line Input #1, SERVER_SECURITY_TAG1$
      Line Input #1, SERVER_SECURITY_TAG2$
   Close
   
   xd = 0
   Open "..\conf\share_dirs.cfg" For Input As #1
    Do Until EOF(1)
        Input #1, SHARES(xd), SHARESL(xd)
        xd = xd + 1
    Loop
    Close 1
    xd = 0
   
   Open "..\conf\http.cfg" For Input As #1
   
   Do Until EOF(1)
      Line Input #1, x$
      d = InStr(x$, "=")
      a$ = Left$(x$, d - 1)
      b$ = Right$(x$, Len(x$) - d)
      
      Select Case a$
         Case "ServerName"
            Longbow.ServerName = b$
         
         Case "ServerAdmin"
            Longbow.ServerAdmin = b$
         
         Case "ListenPort"
            Longbow.ListenPort = Val(b$)
         
         Case "MaxSocks"
            Longbow.MaxSocks = Val(b$)
         
         Case "DefaultRoot"
            Longbow.DefaultRoot = b$
         
         Case "DocLoc"
            Longbow.DocLoc = b$
         
         Case "LogLoc"
            Longbow.LogLoc = b$
         
         Case "LogType"
            Longbow.logtype = Val(b$)
         
         Case "IndexFile"
            Longbow.IndexFile = b$
         
         Case "SecurityFile"
            Longbow.SecurityFile = b$
         
         Case "DirListing"
            Longbow.DirListing = Val(b$)
         
         Case "TimerUpdate"
            Longbow.TimerUpdate = Val(b$)
         
         Case "TimeOut"
            Longbow.TimeOut = Val(b$)
      
      End Select
   Loop
   
   Close 1
   
   xd = 0
   Open "..\conf\mime.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, mimes(xd).ext, mimes(xd).mtype
         xd = xd + 1
      Loop
   Close 1
      
   xd = 0
   Open "..\conf\vdir.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, vdirz(xd).virt, vdirz(xd).real, vdirz(xd).acti
         xd = xd + 1
      Loop
   Close 1
   
   xd = 0
   Open "..\conf\vhost.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, vhost(xd).svr, vhost(xd).root, vhost(xd).acti
         xd = xd + 1
      Loop
   Close 1
   
   xd = 0
   Open "..\conf\users.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, users(xd).username, users(xd).password, users(xd).Active, users(xd).Directory
         xd = xd + 1
      Loop
   Close 1

   'Debug.Print vhost(0).svr
   
   LoadBannedIPS
   
   Exit Sub
LOADSERVERERR:
   Close 1
   MsgBox "!Critial Error!" & vbCrLf & "Error:" & Err.Description & vbCrLf & "Unable to start server", vbCritical, "Longbow Server"
   End
End Sub

Public Sub Cout(text As String)
   
   frmmain.Text1.text = frmmain.Text1.text & text$

End Sub

Public Sub ClrScr()
   
   frmmain.Text1.text = ""

End Sub

Public Sub CloseServer()
   
   
   For t = 0 To Longbow.MaxSocks
      frmmain.ws(t).Close
      If t <> 0 Then Unload frmmain.ws(t)
   Next t

End Sub

Public Sub WXB(socket As Integer, text As String)
  ' Debug.Print socket; text$
   sx(socket).Buffer = sx(socket).Buffer & text$

End Sub

Public Sub WX_FILE(socket As Integer, filename As String)
   ' Writes a file to the socket, this file will be duplicated
   ' exactly, no SSI or modification
   ' USES:IMAGES,NON-ACTIVE HTML, TEXTFILES
   On Error GoTo NOFILE
   dd = FreeFile
   
   Open filename For Binary As #dd
      f1 = LOF(dd)
      f2$ = Space$(f1)
      Get #1, , f2$
   Close dd
   
   sx(socket).Buffer = f2$
   Exit Sub
NOFILE:
   Close dd

End Sub


Public Sub WX_PROPER(sck As Integer, filename As String)
   WX_FILE sck, filename
   ' TO BE CHANGED ;)

End Sub

Public Function SetStringLen(strtochange As String, setlen As Long) As String
   Select Case Len(strtochange$)
   
      
      Case Is > setlen
         If Len(strtochange$) > setlen Then
            SetStringLen = Left$(strtochange$, setlen)
         End If
      
      
      Case setlen
         
         If Len(strtochange$) = setlen Then
            SetStringLen = strtochange$
         End If
      
      
      Case Is < setlen
         
         If Len(strtochange$) < setlen Then
            SetStringLen = strtochange$ & Space$(setlen - Len(strtochange$))
         End If
   
   
   End Select

End Function




Public Sub WLog(TextToLog As String, logtype As Integer)
    
    ' Exit if no logging
    If Longbow.logtype = 0 And logtype <> 0 Then Exit Sub

    ' Exit if error only logging and logging is not for an error
    If logtype = 2 And Longbow.logtype = 1 Then Exit Sub
   
   If ServerLogFile = "" Then
      ServerLogFile = ReplaceStr(Date$ & "_" & Time$, "-", "_")
      ServerLogFile = ReplaceStr(ServerLogFile, ":", "_")
      ServerLogFile = Longbow.LogLoc & "\" & ServerLogFile & ".log"
   End If
   
   dx = FreeFile
   
   Open ServerLogFile For Append As #dx
   Print #dx, Time$ & "," & Date$ & "," & TextToLog$
   Close dx

End Sub

Public Sub AppendIPBan(sck As Integer)
   
   
   For t = 0 To 200
      If ipban(t) = "" Then ipban(t) = frmmain.ws(t).RemoteHostIP: Exit Sub
   Next t
   
   WLog "Socket " & Trim$(Str$(sck)) & " banned IP:" & frmmain.ws(t).RemoteHostIP, 2
End Sub

Public Sub RemoveIPBan(ip As String)
   
   
   For t = 0 To 200
      If ipban(t) = ip$ Then ipban(t) = "": Exit Sub
   Next t
   
   WLog "IP Unbanned " & ip$, 2
End Sub

Public Function IPBanned(sck As Integer) As Integer
   
   
   For t = 0 To 200
      If ipban(t) = frmmain.ws(sck).RemoteHostIP Then IPBanned = 1: Exit Function
   Next t
   
   IPBanned = 0
End Function

Public Sub SaveBannedIPS()
   Dim dxx As Long
   dxx = FreeFile
   
   Open "..\conf\banip.ini" For Output As #dxx
   
   
   For t = 0 To 200
      If ipban(t) <> "" Then Print #1, ipban(t)
   
   Next t
   
   Close dxx
End Sub

Public Sub LoadBannedIPS()
   Dim dxx, xy As Long
   dxx = FreeFile
   xy = 0
   
   Open "..\conf\banip.ini" For Input As #dxx
   
   
   Do Until EOF(dxx)
      Line Input #dxx, ipban(xy)
      xy = xy + 1
   Loop
   
   Close dxx
End Sub

Public Sub AppendIPBanIP(ip As String)
   
   
   For t = 0 To 200
      If ipban(t) = "" Then ipban(t) = ip$: Exit Sub
   Next t
   
   WLog "IP Banned " & ip$, 2
End Sub

Public Sub CLOSEDOWNSERVER()
    ' 1) Save The Configuration Files
    ' 2) Close All Sockets
    ' 3) End The Server Program
    
    'banip.ini
    Open "..\conf\banip.ini" For Output As #3
    For t = 0 To 200
        If ipban(t) <> "" Then Print #3, ipban(t)
    Next t
    Close 3
    
    'dircols.cfg
    Open "..\conf\dircols.cfg" For Output As #3
    Print #3, "dir_backcolor," & DIR_BACKCOLOR
    Print #3, "dir_headcolor," & DIR_HEADCOLOR
    Print #3, "dir_listcolor," & DIR_LISTCOLOR
    Print #3, "dir_barcolor," & DIR_BARCOLOR
    Print #3, "dir_listface," & DIR_LISTFACE
    Print #3, "dir_headface," & DIR_HEADFACE
    Close 3

    'http.cfg
    Open "..\conf\http.cfg" For Output As #3
    Print #3, "ServerName=" & Longbow.ServerName
    Print #3, "ServerAdmin=" & Longbow.ServerAdmin
    Print #3, "ListenPort=" & Trim$(Str$(Longbow.ListenPort))
    Print #3, "MaxSocks=" & Trim$(Str$(Longbow.MaxSocks))
    Print #3, "DefaultRoot=" & Longbow.DefaultRoot
    Print #3, "DocLoc=" & Longbow.DocLoc
    Print #3, "LogLoc=" & Longbow.LogLoc
    Print #3, "LogType=" & Trim$(Str$(Longbow.logtype))
    Print #3, "IndexFile=" & Longbow.IndexFile
    Print #3, "SecurityFile=" & Longbow.SecurityFile
    Print #3, "DirListing=" & Trim$(Str$(Longbow.DirListing))
    Print #3, "TimerUpdate=" & Trim$(Str$(Longbow.TimerUpdate))
    Print #3, "TimeOut=" & Trim$(Str$(Longbow.TimeOut))
    Close 3
    
    'mime.cfg
    Open "..\conf\mime.cfg" For Output As #3
    For t = 0 To 200
        If mimes(t).ext <> "" Then
            Print #3, mimes(t).ext & "," & mimes(t).mtype
        End If
    Next t
    Close 3
    
    'scriptsec.cfg
    Open "..\conf\scriptsec.cfg" For Output As #3
    Print #3, m_main.SERVER_SECURITY_TAG1
    Print #3, m_main.SERVER_SECURITY_TAG2
    Close 3
    
    'users.cfg
    Open "..\conf\users.cfg" For Output As #3
    For t = 0 To 2000
        If users(t).username <> "" Then
            Print #3, users(t).username & "," & users(t).password & "," & LCase$(users(t).Active) & "," & users(t).Directory
        End If
    Next t
    Close 3
    
    'vdir.cfg
    Open "..\conf\vdir.cfg" For Output As #3
    For t = 0 To 60
        If vdirz(t).real <> "" Then
            Print #3, vdirz(t).virt & "," & vdirz(t).real & "," & vdirz(t).acti
        End If
    Next t
    Close 3
    
    'vhost.cfg
    Open "..\conf\vhost.cfg" For Output As #3
    For t = 0 To 60
        If vhost(t).svr <> "" Then
            Print #3, vhost(t).svr & "," & vhost(t).root & "," & vhost(t).acti
        End If
    Next t
    Close 3
        
    'share_dirs.cfg
    Open "..\conf\share_dirs.cfg" For Output As #3
    For t = 0 To 700
        If SHARES(t) <> "" Then
            Print #3, SHARES(t) & "," & SHARESL(t)
        End If
    Next t
    Close 3
    
    
    Cout "Closed."
    
    For t = 0 To Longbow.MaxSocks
        frmmain.ws(t).Close
    Next t
    
    End
End Sub

Public Sub ShowSvEr(text As String)
    frmSvErr.ShowMe text$
End Sub

Public Function UnRidFormatting(text As String) As String
      para = text
      para = ReplaceStr(para, Chr$(34), "%22")
      para = ReplaceStr(para, "<", "%3C")
      para = ReplaceStr(para, ">", "%3E")
      para = ReplaceStr(para, " ", "+")
      para = ReplaceStr(para, "<br>", "%0D%0A")
      para = ReplaceStr(para, "!", "%21")
      para = ReplaceStr(para, "&quot;", "%22")
      para = ReplaceStr(para, " ", "%20")
      para = ReplaceStr(para, "§", "%A7")
      para = ReplaceStr(para, "$", "%24")
      para = ReplaceStr(para, "%", "%25")
      para = ReplaceStr(para, "&", "%26")
      para = ReplaceStr(para, "/", "%2F")
      para = ReplaceStr(para, "(", "%28")
      para = ReplaceStr(para, ")", "%29")
      para = ReplaceStr(para, "=", "%3D")
      para = ReplaceStr(para, "?", "%3F")
      para = ReplaceStr(para, "²", "%B2")
      para = ReplaceStr(para, "³", "%B3")
      para = ReplaceStr(para, "{", "%7B")
      para = ReplaceStr(para, "[", "%5B")
      para = ReplaceStr(para, "]", "%5D")
      para = ReplaceStr(para, "}", "%7D")
      para = ReplaceStr(para, "\", "%5C")
      para = ReplaceStr(para, "ß", "%DF")
      para = ReplaceStr(para, "#", "%23")
      para = ReplaceStr(para, "'", "%27")
      para = ReplaceStr(para, ":", "%3A")
      para = ReplaceStr(para, ",", "%2C")
      para = ReplaceStr(para, ";", "%3B")
      para = ReplaceStr(para, "`", "%60")
      para = ReplaceStr(para, "~", "%7E")
      para = ReplaceStr(para, "+", "%2B")
      para = ReplaceStr(para, "´", "%B4")
    UnRidFormatting = para
End Function

Public Sub CheckInitUpdate()
    On Error Resume Next
    Dim x$, t%, y$, z$, a%, b%
    Dim OLDSECFILE$
    ' Check the configuration files for major changes, and present the user with the Integrity Update
    ' window if major changes are necessary
    
    If InStr(Command$, "-intupdate") Then Exit Sub
    
    If Exists("..\conf\http.old") = 0 Then Exit Sub
    
    Open "..\conf\http.old" For Input As #10
        Do Until EOF(10)
            Line Input #10, x$
                a = InStr(x$, "=")
                y$ = Left$(x$, a - 1)
                z$ = Right$(x$, Len(x$) - a)
                Select Case y$
                    Case "SecurityFile"
                        If LCase$(z$) <> LCase$(Longbow.SecurityFile) Then OLDSECFILE = z$
                End Select
        Loop
    Close 10
        
    If OLDSECFILE$ = "" Then Exit Sub
    
    frmIntUpdate.Show 0
    frmIntUpdate.lblIntUpdate.Caption = "Server Security File Renaming..."
    
    For t = 0 To 700
        If SHARES(t) <> "" Then
            
            If Exists(SHARES(t) & "\" & OLDSECFILE$) = 1 Then
                
                FileCopy SHARES(t) & "\" & OLDSECFILE$, SHARES(t) & "\" & Longbow.SecurityFile
                
                SetAttr SHARES(t) & "\" & Longbow.SecurityFile, vbHidden + vbSystem
                SetAttr SHARES(t) & "\" & OLDSECFILE$, vbNormal
                
                Kill SHARES(t) & "\" & OLDSECFILE$
            End If
        End If
    Next t
        
    Kill "..\conf\http.old"
    
    frmIntUpdate.Hide
End Sub

Public Function SetIfNotNull(text As String, norm As String) As String
    If text = "" Then SetIfNotNull = norm Else SetIfNotNull = text
End Function

Public Function GetCharReps(text As String, ch As String) As Integer
    Dim a%, b%, c%
    a = Len(text$)
    For b = 1 To a
        If Mid$(text$, b, 1) = ch$ Then c = c + 1
    Next b
    GetCharReps = c
End Function
