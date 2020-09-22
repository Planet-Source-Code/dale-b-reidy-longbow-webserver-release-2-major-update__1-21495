VERSION 5.00
Begin VB.Form frmServerConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Configuration"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update Server"
      Height          =   315
      Left            =   75
      TabIndex        =   32
      Top             =   5100
      Width           =   1740
   End
   Begin VB.CommandButton cmdDirView 
      Caption         =   "Directory View Settings"
      Height          =   315
      Left            =   3000
      TabIndex        =   31
      Top             =   5100
      Width           =   1965
   End
   Begin VB.CommandButton cmdSvrConfClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   5025
      TabIndex        =   1
      Top             =   5100
      Width           =   1065
   End
   Begin VB.Frame fraGeneral 
      BackColor       =   &H00000000&
      Caption         =   "General Server Configuration Administration"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4965
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6015
      Begin VB.Frame fraDirView 
         BackColor       =   &H00000000&
         Caption         =   "Directory View"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   225
         TabIndex        =   28
         Top             =   3150
         Width           =   5340
         Begin VB.OptionButton optGFX 
            BackColor       =   &H00000000&
            Caption         =   "Graphical "
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1650
            TabIndex        =   30
            Top             =   150
            Width           =   1515
         End
         Begin VB.OptionButton optTEXT 
            BackColor       =   &H00000000&
            Caption         =   "Text Only"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   3450
            TabIndex        =   29
            Top             =   150
            Width           =   1290
         End
      End
      Begin VB.TextBox txtDefaultRoot 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   27
         Top             =   1500
         Width           =   3765
      End
      Begin VB.TextBox txtRequestTimeout 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   4200
         TabIndex        =   25
         Top             =   4500
         Width           =   1215
      End
      Begin VB.TextBox txtTimerUpdate 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   23
         Top             =   4500
         Width           =   1215
      End
      Begin VB.TextBox txtSecurityFile 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   21
         Top             =   4125
         Width           =   1815
      End
      Begin VB.TextBox txtIndexFile 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   19
         Top             =   3750
         Width           =   1815
      End
      Begin VB.Frame fraLogType 
         BackColor       =   &H00000000&
         Caption         =   "Log Type"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   465
         Left            =   225
         TabIndex        =   14
         Top             =   2625
         Width           =   5340
         Begin VB.OptionButton optFull 
            BackColor       =   &H00000000&
            Caption         =   "Full"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   3225
            TabIndex        =   17
            Top             =   150
            Width           =   1515
         End
         Begin VB.OptionButton optErrors 
            BackColor       =   &H00000000&
            Caption         =   "Errors"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2025
            TabIndex        =   16
            Top             =   150
            Width           =   1515
         End
         Begin VB.OptionButton optNone 
            BackColor       =   &H00000000&
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   900
            TabIndex        =   15
            Top             =   150
            Width           =   1515
         End
      End
      Begin VB.TextBox txtLogLocation 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   13
         Top             =   2250
         Width           =   3765
      End
      Begin VB.TextBox txtDocLocation 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   11
         Top             =   1875
         Width           =   3765
      End
      Begin VB.TextBox txtMaxSockets 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   4125
         TabIndex        =   9
         Top             =   1125
         Width           =   1215
      End
      Begin VB.TextBox txtListenPort 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   7
         Top             =   1125
         Width           =   1215
      End
      Begin VB.TextBox txtServerAdmin 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   5
         Top             =   750
         Width           =   3765
      End
      Begin VB.TextBox txtServerName 
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1275
         TabIndex        =   3
         Top             =   375
         Width           =   3765
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Default Root"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   26
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Request Timeout"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   2775
         TabIndex        =   24
         Top             =   4500
         Width           =   1440
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Timer Update"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   22
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Security File"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   20
         Top             =   4125
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Index File"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   18
         Top             =   3750
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Log Location"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   12
         Top             =   2250
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Doc Location"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   10
         Top             =   1875
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Max Sockets"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   2925
         TabIndex        =   8
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Listen Port"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   6
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Server Admin"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   4
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Server Name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   75
         TabIndex        =   2
         Top             =   375
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmServerConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSvrConfClose_Click()
    Longbow.DefaultRoot = SetIfNotNull(txtDefaultRoot.text, Longbow.DefaultRoot)
    If optGFX.Value = True Then Longbow.DirListing = 2 Else Longbow.DirListing = 1
    Longbow.DocLoc = SetIfNotNull(txtDocLocation.text, Longbow.DocLoc)
    Longbow.IndexFile = SetIfNotNull(txtIndexFile.text, Longbow.IndexFile)
    Longbow.ListenPort = SetIfNotNull(txtListenPort.text, Trim$(Str$(Longbow.ListenPort)))
    Longbow.LogLoc = SetIfNotNull(txtLogLocation.text, Longbow.LogLoc)
    If optNone.Value = True Then Longbow.LogType = 0
    If optErrors.Value = True Then Longbow.LogType = 1
    If optFull.Value = True Then Longbow.LogType = 2
    Longbow.MaxSocks = SetIfNotNull(txtMaxSockets.text, Trim$(Str$(Longbow.MaxSocks)))
    Longbow.SecurityFile = SetIfNotNull(txtSecurityFile.text, Longbow.SecurityFile)
    Longbow.ServerAdmin = SetIfNotNull(txtServerAdmin.text, Longbow.ServerAdmin)
    Longbow.ServerName = SetIfNotNull(txtServerName.text, Longbow.ServerName)
    Longbow.TimeOut = SetIfNotNull(txtRequestTimeout.text, Trim$(Str$(Longbow.TimeOut)))
    Longbow.TimerUpdate = SetIfNotNull(txtTimerUpdate.text, Trim$(Str$(Longbow.TimerUpdate)))
    Me.Visible = False
End Sub

Public Sub ShowMe()
    txtDefaultRoot.text = Longbow.DefaultRoot
    If Longbow.DirListing = 2 Then optGFX.Value = True Else optTEXT.Value = True
    txtDocLocation.text = Longbow.DocLoc
    txtIndexFile.text = Longbow.IndexFile
    txtListenPort.text = Longbow.ListenPort
    txtLogLocation.text = Longbow.LogLoc
    If Longbow.LogType = 0 Then optNone.Value = True
    If Longbow.LogType = 1 Then optErrors.Value = True
    If Longbow.LogType = 2 Then optFull.Value = True
    txtMaxSockets.text = Longbow.MaxSocks
    txtSecurityFile.text = Longbow.SecurityFile
    txtServerAdmin.text = Longbow.ServerAdmin
    txtServerName.text = Longbow.ServerName
    txtRequestTimeout.text = Longbow.TimeOut
    txtTimerUpdate.text = Longbow.TimerUpdate

    Me.Visible = True
End Sub

Private Sub Command1_Click()
   Dim t As Integer
   frmmain.ws(0).Close
   For t = 1 To Longbow.MaxSocks
      frmmain.ws(t).Close
      Unload frmmain.ws(t)
      Unload fsu.pic(t)
      Unload fsu.fi(t)
      Unload fsu.di(t)
   
   Next t
    InitServer
End Sub
