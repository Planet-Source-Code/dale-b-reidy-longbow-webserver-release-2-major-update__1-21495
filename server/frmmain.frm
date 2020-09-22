VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LongBow Server"
   ClientHeight    =   5205
   ClientLeft      =   2145
   ClientTop       =   3780
   ClientWidth     =   6855
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6855
   Begin VB.CheckBox Check1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   150
      TabIndex        =   13
      ToolTipText     =   "Unlock the Close Server button, this is used as a precaution"
      Top             =   4800
      Width           =   200
   End
   Begin VB.CommandButton cmdDirShare 
      Caption         =   "Director&y Sharing"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5175
      TabIndex        =   14
      ToolTipText     =   "Administrate directory sharing and security"
      Top             =   3600
      Width           =   1590
   End
   Begin VB.CommandButton cmdSvrConf 
      Caption         =   "Server &Config"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5175
      TabIndex        =   9
      ToolTipText     =   "Change general server configurations"
      Top             =   3150
      Width           =   1605
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "&Users"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5175
      TabIndex        =   8
      ToolTipText     =   "Change access priviledges for users"
      Top             =   2700
      Width           =   1605
   End
   Begin VB.CommandButton cmdVHost 
      Caption         =   "Virtual &Hosts"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5175
      TabIndex        =   7
      ToolTipText     =   "Administrate virtual hosts"
      Top             =   2250
      Width           =   1605
   End
   Begin VB.CommandButton cmdVDir 
      Caption         =   "Virtual &Directories"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5175
      TabIndex        =   6
      ToolTipText     =   "Administrate virtual directories"
      Top             =   1800
      Width           =   1605
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Read Log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5025
      TabIndex        =   5
      ToolTipText     =   "Show the current server log"
      Top             =   4725
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sto&p"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3375
      TabIndex        =   1
      ToolTipText     =   "Temporarily stop the server from accepting requests"
      Top             =   4725
      Width           =   1590
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   4575
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   4185
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4980
   End
   Begin VB.Timer sxu 
      Interval        =   10
      Left            =   105
      Top             =   105
   End
   Begin VB.Timer sxt 
      Interval        =   30
      Left            =   105
      Top             =   105
   End
   Begin VB.CommandButton Command2 
      Caption         =   "St&art"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1725
      TabIndex        =   2
      ToolTipText     =   "Allow the server to start allowing connection requests"
      Top             =   4725
      Width           =   1590
   End
   Begin VB.Timer sxz 
      Interval        =   1000
      Left            =   4125
      Top             =   225
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close Ser&ver"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   12
      ToolTipText     =   "Completely close down the server, saving all logs and configuration changes"
      Top             =   4725
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   1710
      Left            =   5100
      Picture         =   "frmmain.frx":27A2
      Top             =   75
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Connected Users"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   2775
      TabIndex        =   11
      Top             =   4350
      Width           =   1335
   End
   Begin VB.Label lblConnUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   300
      Left            =   4125
      TabIndex        =   10
      Top             =   4350
      Width           =   975
   End
   Begin VB.Label lblReq 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   300
      Left            =   900
      TabIndex        =   4
      Top             =   4350
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requests"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   4350
      Width           =   960
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then Command4.Enabled = True
    If Check1.Value <> vbChecked Then Command4.Enabled = False
End Sub



Private Sub cmdDirShare_Click()
    frmDirShare.ShowMe
End Sub

Private Sub cmdSvrConf_Click()
    FileCopy "..\conf\http.cfg", "..\conf\http.old"
    frmServerConfig.ShowMe
End Sub

Private Sub cmdUsers_Click()
    frmUsrAdmin.ShowMe
End Sub

Private Sub cmdVDir_Click()
    frmVDirs.ShowMe
End Sub

Private Sub cmdVHost_Click()
    frmVHosts.ShowMe
End Sub

Private Sub Command1_Click()
 StopServer
End Sub

Private Sub Command2_Click()
   StartServer
End Sub

Private Sub Command3_Click()
 Shell "notepad.exe " & ServerLogFile, vbNormalFocus
End Sub

Public Sub Command4_Click()
    Cout "Closing..."
    ' Save The Configuration Files
    CLOSEDOWNSERVER
    
End Sub

Private Sub Form_Load()
   InitServer
   CheckInitUpdate
   WLog "Server Started", 0
   
   Cout "LongBow Server 1.0b" & vbCrLf
   Cout "-------------------" & vbCrLf
   Cout "Listening On Port " & Trim$(Str$(Longbow.ListenPort)) & vbCrLf
   Cout "Server Config Loaded" & vbCrLf
   Cout "LogFile:" & ServerLogFile & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CloseServer
   End
End Sub

Private Sub sxt_Timer()
   On Error GoTo SXTERR
   Dim t As Integer
   If sxt.Tag = "POTATO" Then Exit Sub
   sxt.Tag = "POTATO"
   Do Until sxt.Enabled = False
    'lblSXT.Caption = lblSXT.Caption + 1
    lblReq.Caption = Trim$(Str$(NumReq))
    
    For t = 1 To Longbow.MaxSocks Step 2
                  DoEvents
                  If sx(t).Header <> "" And sx(t).Reqok = False And sx(t).Buffer = "" And ws(t).State = sckConnected Then
                     ProcessHeader t
                  End If

                     sx(t).TimeAlive = sx(t).TimeAlive + 1
            
                    
                     If sx(t).TimeAlive > (Longbow.TimeOut / 10) Then
                        ws(t).Close
                        sx(t).Buffer = ""
                        sx(t).Header = ""
                        sx(t).Reqok = False
                        sx(t).TimeAlive = 0
                     End If

                  If sx(t).Reqok = True And ws(t).State <> sckConnected Then
                        ws(t).Close
                        sx(t).Buffer = ""
                        sx(t).Header = ""
                        sx(t).Reqok = False
                        sx(t).TimeAlive = 0
                        ws(t).Tag = ""
                  End If

                  If sx(t).Reqok = True And ws(t).State = sckConnected Then
            
                     a = Len(sx(t).Buffer)
            
                     'Debug.Print a
            
                     If a = 0 And frmmain.ws(t).Tag = "LASTPACKET" Then
                        ws(t).Close
            
                        sx(t).Buffer = ""
                        sx(t).Header = ""
                        sx(t).Reqok = False
                        sx(t).TimeAlive = 0
                        ws(t).Tag = ""
                        GoTo RABIDO
                     End If
                     'If a = 0 Then GoTo RABIDO
                     If a > 3000 Then g = 3000 Else g = a: ws(t).Tag = "LASTSEND"
                     r$ = Left$(sx(t).Buffer, g)
                     sx(t).Buffer = Right$(sx(t).Buffer, Len(sx(t).Buffer) - g)

                     ws(t).SendData r$

                     sx(t).TimeAlive = 0
                  End If
RABIDO:
            '         ws(t).SendData sx(t).Buffer
            '         sx(t).Reqok = False
            '         sx(t).Buffer = ""
            
            
                  If sx(t).Reqok = True And ws(t).State <> sckConnected Then
                     ws(t).Close
                     sx(t).Buffer = ""
                     sx(t).Header = ""
                     sx(t).Reqok = False
                     sx(t).TimeAlive = 0
                     ws(t).Tag = ""
                  End If
   Next t
   Loop
   sxt.Tag = ""
   Exit Sub
SXTERR:
   sxt.Tag = ""
   Debug.Print "SXT Error " & Err.Description
End Sub

Private Sub sxu_Timer()
   Dim t As Integer
   For t = 0 To Longbow.MaxSocks Step 2
      'lblSXU.Caption = lblSXU.Caption + 1
      If t <> 0 Then
      
      
      
      DoEvents

      If sx(t).Header <> "" And sx(t).Reqok = False And sx(t).Buffer = "" And ws(t).State = sckConnected Then

         ProcessHeader t
      End If

         sx(t).TimeAlive = sx(t).TimeAlive + 1
      'If ws(t).State = sckConnected Then Debug.Print sx(t).TimeAlive

         If sx(t).TimeAlive > Longbow.TimeOut Then
            ws(t).Close
            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
         End If

      If sx(t).Reqok = True And ws(t).State <> sckConnected Then
            ws(t).Close
            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
            ws(t).Tag = ""
      End If

      If sx(t).Reqok = True And ws(t).State = sckConnected Then

         a = Len(sx(t).Buffer)

         'Debug.Print a

         If a = 0 And frmmain.ws(t).Tag = "LASTPACKET" Then
            ws(t).Close

            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
            ws(t).Tag = ""
            GoTo RABIDO
         End If
         'If a = 0 Then GoTo RABIDO
         If a > 3000 Then g = 3000 Else g = a: ws(t).Tag = "LASTSEND"
         r$ = Left$(sx(t).Buffer, g)
         sx(t).Buffer = Right$(sx(t).Buffer, Len(sx(t).Buffer) - g)

         ws(t).SendData r$

         sx(t).TimeAlive = 0
      End If
RABIDO:
'         ws(t).SendData sx(t).Buffer
'         sx(t).Reqok = False
'         sx(t).Buffer = ""


      If sx(t).Reqok = True And ws(t).State <> sckConnected Then
         ws(t).Close
         sx(t).Buffer = ""
         sx(t).Header = ""
         sx(t).Reqok = False
         sx(t).TimeAlive = 0
         ws(t).Tag = ""
      End If


      End If
      
   Next t
      
End Sub

Private Sub sxz_Timer()
    Dim t, u As Long
    For t = 1 To Longbow.MaxSocks
        If ws(t).State = sckConnected Then u = u + 1
    Next t
    lblConnUsr.Caption = Trim$(Str$(u))
End Sub

Private Sub ws_Close(Index As Integer)
   sx(Index).Buffer = ""
   sx(Index).Header = ""
   sx(Index).Reqok = False
   sx(Index).TimeAlive = 0
   ws(Index).Tag = ""
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   Dim t As Integer
   For t = 1 To Longbow.MaxSocks
      If ws(t).State = sckClosing Then ws(t).Close
      If ws(t).State = sckClosed Then
         ws(t).Accept requestID
         sx(t).Buffer = ""
         sx(t).Header = ""
         sx(t).Reqok = False
         sx(t).TimeAlive = 0
         ws(t).Tag = ""
         If IPBanned(t) = 1 Then
            WriteHTTP t, 403, "-"
            sx(t).Reqok = True
            sx(t).Header = "HAS BEEN BANNED"
            Exit Sub
         End If
         Exit Sub
      End If
   Next t
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   On Error GoTo WSDAO
   ws(Index).GetData sx(Index).Header
   'ProcessHeader Index
   Exit Sub
WSDAO:
         ws(Index).Close
         sx(Index).Buffer = ""
         sx(Index).Header = ""
         sx(Index).Reqok = False
         sx(Index).TimeAlive = 0
         ws(Index).Tag = ""
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   ws(Index).Close
   ' If the listening socket closes, we're buggered, so open it up again, probably causing another error
   ' and the program to crash, but hey, windows crashes, so why can't this? :p
   If Index = 0 Then ws(0).Listen
End Sub

Private Sub ws_SendComplete(Index As Integer)
If ws(Index).Tag = "LASTSEND" Then
   ws(Index).Tag = "LASTPACKET"
End If
End Sub

