VERSION 5.00
Begin VB.Form frmVDirs 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Directories"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
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
      Left            =   4875
      TabIndex        =   1
      Top             =   4125
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox Combo1 
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
         Left            =   3150
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   600
         Width           =   2715
      End
      Begin VB.ListBox ls2 
         Height          =   2010
         Left            =   900
         TabIndex        =   11
         Top             =   975
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton cmdUpdate 
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
         Left            =   3150
         TabIndex        =   10
         Top             =   3225
         Width           =   1290
      End
      Begin VB.CommandButton cmdDelete 
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
         Left            =   4575
         TabIndex        =   9
         Top             =   2775
         Width           =   1290
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   3150
         TabIndex        =   8
         Top             =   2775
         Width           =   1290
      End
      Begin VB.CheckBox chkActv 
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
         Left            =   3150
         TabIndex        =   7
         Tag             =   "YESNO"
         Top             =   2100
         Width           =   990
      End
      Begin VB.TextBox txtVDir 
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
         Left            =   3150
         TabIndex        =   6
         Top             =   1425
         Width           =   2715
      End
      Begin VB.TextBox txtRDir 
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
         Left            =   3150
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.ListBox ls1 
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
         Height          =   3660
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   2940
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Virtual Directory"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Real Directory"
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
         Left            =   3375
         TabIndex        =   3
         Top             =   300
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmVDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim t As Integer
    For t = 0 To 60
        If (vdirz(t).virt = txtVDir.text) And (txtVDir.text <> "") Then
            ShowSvEr "Virtual Directory Already Exists With This Name"
            Exit Sub
        End If
    Next t
    If (txtVDir.text = "") Or (Combo1.text = "") Then
        ShowSvEr "All Details Not Entered"
        Exit Sub
    End If
    For t = 0 To 60
        If vdirz(t).virt = "" Then
            If Right$(Combo1.text, 1) = "\" Then
                Combo1.text = Left$(Combo1.text, Len(Combo1.text) - 1)
            End If
            If Right$(txtVDir.text, 1) = "/" Then
                txtVDir.text = Left$(txtVDir.text, Len(txtVDir.text) - 1)
            End If
            vdirz(t).virt = txtVDir.text
            'vdirz(t).real = txtRDir.text
            vdirz(t).real = Combo1.text
            If chkActv.Value = 1 Then vdirz(t).acti = "YES" Else vdirz(t).acti = "NO"
            RefreshList
            Exit Sub
        End If
    Next t
End Sub

Private Sub cmdClose_Click()
    Me.Visible = False
End Sub

Public Sub ShowMe()
    ' Startup Routine To Show The Form
    RefreshList
    Me.Visible = True
End Sub
    
Private Sub cmdDELETE_Click()
    Dim t As Integer
    For t = 0 To 60
        If txtVDir.text = vdirz(t).virt Then
            vdirz(t).acti = ""
            vdirz(t).real = ""
            vdirz(t).virt = ""
            RefreshList
            Exit Sub
        End If
    Next t
    RefreshList
End Sub

Private Sub cmdUpdate_Click()
    Dim t As Integer
    For t = 0 To 60
        If txtVDir.text = vdirz(t).virt Then
            'vdirz(t).real = txtRDir.text
            vdirz(t).real = Combo1.text
            If chkActv.Value = 1 Then vdirz(t).acti = "YES" Else vdirz(t).acti = "NO"
        End If
    Next t
    RefreshList
End Sub



Private Sub Combo1_Click()
    Dim t As Integer
    Dim ISSHARED As Integer
    For t = 0 To 700
        If LCase$(Combo1.text) = LCase$(SHARES(t)) Then ISSHARED = 1
    Next t
    If ISSHARED = 1 Then Combo1.ForeColor = vbGreen Else Combo1.ForeColor = vbRed
End Sub

Private Sub ls1_Click()
    Dim r As Integer
    Dim t As Integer
    Dim s As Integer
    Dim ISSHARED As Integer
    If ls1.ListIndex = -1 Then Exit Sub
    r = ls1.ListIndex
    s = Val(ls2.List(r))
    txtVDir.text = vdirz(s).virt
    'txtRDir.text = vdirz(s).real
    ISSHARED = 0
    Combo1.text = vdirz(s).real
    For t = 0 To 700
        If LCase$(Combo1.text) = LCase$(SHARES(t)) Then ISSHARED = 1
    Next t
    If ISSHARED = 1 Then Combo1.ForeColor = vbGreen Else Combo1.ForeColor = vbRed
    If vdirz(s).acti = "YES" Then chkActv.Value = 1 Else chkActv.Value = 0
End Sub

Public Sub RefreshList()
    Dim t As Integer
    ls1.Clear
    ls2.Clear
    Combo1.Clear
    Combo1.ForeColor = vbGreen
    For t = 0 To 60
        If vdirz(t).virt <> "" Then
            ls1.AddItem vdirz(t).virt
            ls2.AddItem Trim$(Str$(t))
        End If
    Next t

    For t = 0 To 700
        If SHARES(t) <> "" Then
            Combo1.AddItem SHARES(t)
        End If
    Next t
    
End Sub
