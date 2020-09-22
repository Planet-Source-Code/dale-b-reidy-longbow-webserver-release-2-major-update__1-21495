VERSION 5.00
Begin VB.Form frmUsrAdmin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Administration"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5595
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
      Left            =   4425
      TabIndex        =   7
      Top             =   4125
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
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
      Left            =   1500
      TabIndex        =   6
      Top             =   3675
      Width           =   990
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3675
      TabIndex        =   5
      Top             =   3675
      Width           =   990
   End
   Begin VB.CommandButton Command1 
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
      Left            =   450
      TabIndex        =   4
      Top             =   3675
      Width           =   990
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "User Information"
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
      Left            =   150
      TabIndex        =   3
      Top             =   75
      Width           =   2715
      Begin VB.TextBox txtUserdir 
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
         Left            =   150
         TabIndex        =   14
         ToolTipText     =   "Users home directory, accessible by eg, http://server:123/~jones for user Jones"
         Top             =   2250
         Width           =   2415
      End
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
         Top             =   2775
         Width           =   1740
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1500
         Width           =   2415
      End
      Begin VB.TextBox txtUsername 
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
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Users Directory"
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
         TabIndex        =   13
         Top             =   2025
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Password"
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
         TabIndex        =   10
         Top             =   1275
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Username"
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
         TabIndex        =   8
         Top             =   525
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Users"
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
      Left            =   2925
      TabIndex        =   0
      Top             =   75
      Width           =   2490
      Begin VB.ListBox lsuh 
         Height          =   1815
         Left            =   600
         TabIndex        =   2
         Top             =   900
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lsUsrs 
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
         Width           =   2340
      End
   End
End
Attribute VB_Name = "frmUsrAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim t As Integer
    If lsUsrs.ListIndex = -1 Then
        For t = 0 To lsuh.ListCount - 1
            If users(Val(lsuh.List(t))).username = txtUsername.text Then lsUsrs.ListIndex = t: Exit For
        Next t
    End If
    
    users(Val(lsuh.List(lsUsrs.ListIndex))).username = txtUsername.text
    users(Val(lsuh.List(lsUsrs.ListIndex))).password = txtPassword.text
    users(Val(lsuh.List(lsUsrs.ListIndex))).Directory = txtUserdir.text
    
    If chkActive.Value = 0 Then
        users(Val(lsuh.List(lsUsrs.ListIndex))).Active = "no"
    Else
        users(Val(lsuh.List(lsUsrs.ListIndex))).Active = "yes"
    End If
    
    ShowMe
End Sub

Private Sub Command2_Click()
        If lsUsrs.ListIndex = -1 Then Exit Sub
        users(lsUsrs.ListIndex).username = ""
        users(lsUsrs.ListIndex).password = ""
        users(lsUsrs.ListIndex).Directory = ""
        users(lsUsrs.ListIndex).Active = "no"
        ShowMe
End Sub

Private Sub Command3_Click()
    Dim t As Integer
        For t = 0 To 2000
            If users(t).username = txtUsername.text Then
                ShowSvEr "User already exists, try update."
                lsUsrs.ListIndex = t
                lsUsrs_Click
                Exit Sub
            End If
        Next t
        
        For t = 0 To 2000
            If users(t).username = "" Then
                users(t).username = txtUsername.text
                users(t).password = txtPassword.text
                users(t).Directory = txtUserdir.text
                If chkActive.Value = 0 Then users(t).Active = "no" Else users(t).Active = "yes"
                ShowMe
                Exit Sub
            End If
        Next t
End Sub

Private Sub Command4_Click()
    Me.Visible = False
End Sub

Public Sub ShowMe()
    Dim t As Integer
    lsUsrs.Clear
    lsuh.Clear
    For t = 0 To 2000
        If users(t).username <> "" Then
            lsUsrs.AddItem users(t).username
            lsuh.AddItem t
        End If
    Next
    Me.Visible = True
End Sub

Private Sub lsUsrs_Click()
    txtUsername.text = users(Val(lsuh.List(lsUsrs.ListIndex))).username
    txtPassword.text = users(Val(lsuh.List(lsUsrs.ListIndex))).password
    txtUserdir.text = users(Val(lsuh.List(lsUsrs.ListIndex))).Directory
    If LCase$(users(Val(lsuh.List(lsUsrs.ListIndex))).Active) = "yes" Then chkActive.Value = 1 Else chkActive.Value = 0
End Sub
