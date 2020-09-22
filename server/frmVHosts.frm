VERSION 5.00
Begin VB.Form frmVHosts 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Hosts"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4050
      TabIndex        =   6
      Top             =   3675
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2550
      TabIndex        =   5
      Top             =   3075
      Width           =   765
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3375
      TabIndex        =   4
      Top             =   3075
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   825
      TabIndex        =   3
      Top             =   3675
      Width           =   840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Virtual Host Information"
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
      Height          =   2940
      Left            =   2550
      TabIndex        =   2
      Top             =   75
      Width           =   2415
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00000000&
         Caption         =   "Active"
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
         Left            =   150
         TabIndex        =   12
         Top             =   2475
         Width           =   1515
      End
      Begin VB.ComboBox cboDir 
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
         Height          =   345
         Left            =   75
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1800
         Width           =   2265
      End
      Begin VB.TextBox txtHostName 
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
         Left            =   75
         TabIndex        =   9
         Top             =   825
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Directory"
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
         Top             =   1500
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Virtual Host Name"
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
         TabIndex        =   8
         Top             =   525
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Virtual Hosts On Server"
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
      Height          =   3540
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2415
      Begin VB.ListBox lsvhi 
         Height          =   1815
         Left            =   750
         TabIndex        =   7
         Top             =   825
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox lsVH 
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
         Height          =   3210
         Left            =   75
         TabIndex        =   1
         Top             =   225
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmVHosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboDir_LostFocus()
    Dim t As Integer
    cboDir.ForeColor = vbRed
    
    For t = 0 To 700
        If SHARES(t) = cboDir.text Then cboDir.ForeColor = vbGreen
    Next t
End Sub

Private Sub Command1_Click()

    If lsVH.ListIndex = -1 Then Exit Sub
    
    vhost(lsvhi.List(lsVH.ListIndex)).acti = "NO"
    vhost(lsvhi.List(lsVH.ListIndex)).root = ""
    vhost(lsvhi.List(lsVH.ListIndex)).svr = ""
    
    ShowMe
End Sub

Private Sub Command2_Click()
    Dim t As Integer
    
    For t = 0 To 60
        If vhost(t).svr = txtHostName.text Then
            ShowSvEr "Virtual Host Already Exists, Try Update"
            Exit Sub
        End If
    Next t
    
    For t = 0 To 60
        If vhost(t).svr = "" Then
            vhost(t).svr = txtHostName.text
            vhost(t).root = cboDir.text
            If chkActive.Value = 0 Then vhost(t).acti = "no" Else vhost(t).acti = "yes"
            Exit For
        End If
    Next t
    
    ShowMe
        
End Sub

Private Sub Command3_Click()
    If lsVH.ListIndex = -1 Then Exit Sub
    vhost(lsvhi.List(lsVH.ListIndex)).root = cboDir.text
    vhost(lsvhi.List(lsVH.ListIndex)).svr = txtHostName.text
    If chkActive.Value = 0 Then vhost(lsvhi.List(lsVH.ListIndex)).acti = "NO" Else vhost(lsvhi.List(lsVH.ListIndex)).acti = "YES"
    ShowMe
End Sub

Private Sub Command4_Click()
    Me.Visible = False
End Sub

Public Sub ShowMe()
    Dim t As Integer
    lsvhi.Clear
    lsVH.Clear
    cboDir.Clear
    
    For t = 0 To 60
        If vhost(t).svr <> "" Then
            lsVH.AddItem vhost(t).svr
            lsvhi.AddItem t
        End If
    Next t
    
    For t = 0 To 700
        If SHARES(t) <> "" Then
            cboDir.AddItem SHARES(t)
        End If
    Next t
    
    Me.Visible = True
        
End Sub

Private Sub lsVH_Click()
    Dim t As Integer
    
    txtHostName.text = vhost(lsvhi.List(lsVH.ListIndex)).svr
    cboDir.text = vhost(lsvhi.List(lsVH.ListIndex)).root
    If vhost(lsvhi.List(lsVH.ListIndex)).acti = "YES" Then chkActive.Value = 1 Else chkActive.Value = 0
    
    cboDir.ForeColor = vbRed
    
    For t = 0 To 700
        If SHARES(t) = cboDir.text Then cboDir.ForeColor = vbGreen
    Next t
    
End Sub
