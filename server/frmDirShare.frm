VERSION 5.00
Begin VB.Form frmDirShare 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Sharing"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Add New Directory To Share List"
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
      Height          =   1365
      Left            =   3525
      TabIndex        =   15
      Top             =   2850
      Width           =   4140
      Begin VB.CommandButton Command3 
         BackColor       =   &H00000000&
         Caption         =   "Add (Include &Subfolders)"
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
         Left            =   1200
         TabIndex        =   19
         ToolTipText     =   "Also share subfolders under this folder"
         Top             =   825
         Width           =   2265
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         Caption         =   "A&dd"
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
         Left            =   150
         TabIndex        =   18
         ToolTipText     =   "Add this folder without sharing directories below"
         Top             =   825
         Width           =   990
      End
      Begin VB.TextBox txtNEWDIR 
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
         Left            =   825
         TabIndex        =   16
         ToolTipText     =   "Directory to share"
         Top             =   375
         Width           =   3165
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Directory "
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
         TabIndex        =   17
         Top             =   375
         Width           =   1140
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cl&ose"
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
      Left            =   6525
      TabIndex        =   14
      Top             =   4425
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Change Directory Share Settings"
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
      Height          =   2715
      Left            =   3525
      TabIndex        =   1
      Top             =   75
      Width           =   4140
      Begin VB.CommandButton Command4 
         Caption         =   "&Apply Below"
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
         Left            =   2625
         TabIndex        =   20
         ToolTipText     =   "Add these security settings to directories under this one"
         Top             =   1875
         Width           =   1140
      End
      Begin VB.CommandButton cmdCLEAR 
         Caption         =   "C&lear"
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
         TabIndex        =   11
         Top             =   1875
         Width           =   1065
      End
      Begin VB.CommandButton cmdADDUPD 
         Caption         =   "&Change"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1875
         Width           =   1290
      End
      Begin VB.CheckBox chkDIRVIEW 
         BackColor       =   &H00000000&
         Caption         =   "Dir View"
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
         Left            =   1500
         TabIndex        =   9
         Top             =   1500
         Width           =   1065
      End
      Begin VB.CheckBox chkEXECUTE 
         BackColor       =   &H00000000&
         Caption         =   "Execute"
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
         Left            =   1500
         TabIndex        =   8
         Top             =   1275
         Width           =   1065
      End
      Begin VB.CheckBox chkREAD 
         BackColor       =   &H00000000&
         Caption         =   "Read"
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
         Left            =   525
         TabIndex        =   7
         Top             =   1500
         Width           =   1065
      End
      Begin VB.CheckBox chkSECURE 
         BackColor       =   &H00000000&
         Caption         =   "Secure"
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
         Left            =   525
         TabIndex        =   6
         Top             =   1275
         Width           =   1065
      End
      Begin VB.TextBox txtDOMAIN 
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
         Top             =   825
         Width           =   2340
      End
      Begin VB.ComboBox cboUSERS 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cboUSERS"
         Top             =   375
         Width           =   2640
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Location name"
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
         Top             =   825
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Allowed users"
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
         Top             =   450
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Shareable Directory List"
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
      Height          =   4740
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3390
      Begin VB.DirListBox rec 
         Height          =   1440
         Left            =   525
         TabIndex        =   22
         Top             =   2100
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ListBox lsDS 
         Height          =   2595
         Left            =   1125
         TabIndex        =   21
         Top             =   975
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton cmdDELETE 
         Caption         =   "D&elete"
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
         Left            =   75
         TabIndex        =   13
         Top             =   4350
         Width           =   1290
      End
      Begin VB.ListBox lsShDirList 
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
         Height          =   4110
         Left            =   75
         TabIndex        =   12
         Top             =   225
         Width           =   3240
      End
   End
End
Attribute VB_Name = "frmDirShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe()
    'lsShDirList    List of sharable locations
    'cboUSERS       Combo box of users
    
    Dim t As Integer
    
    lsDS.Clear
    lsShDirList.Clear
    cboUSERS.Clear
    For t = 0 To 700
        If SHARES(t) <> "" Then
            lsShDirList.AddItem SHARES(t) & " > " & SHARESL(t)
            lsDS.AddItem Trim$(Str$(t))
        End If
    Next t
    

    cboUSERS.AddItem "any"
    For t = 0 To 2000
        If users(t).username <> "" Then
            If LCase$(users(t).Active = "no") Then cboUSERS.AddItem "!" & users(t).username Else cboUSERS.AddItem users(t).username
        End If
    Next t
    
    Me.Visible = True
End Sub

Private Sub cmdADDUPD_Click()
    Dim x$
    Dim DIR_DOMAIN$, DIR_USERS$, DIR_READ%, DIR_WRITE%, DIR_VIEW%, DIR_EXECUTE%, DIR_SECURITY%
    
    If cboUSERS.text = "" Then cboUSERS.text = "all"
    If txtDOMAIN.text = "" Then txtDOMAIN.text = "Restricted"
    
    DIR_DOMAIN$ = txtDOMAIN.text
    DIR_READ = chkREAD.Value
    DIR_WRITE = 1
    DIR_VIEW = chkDIRVIEW.Value
    DIR_EXECUTE = chkEXECUTE.Value
    DIR_SECURITY = chkSECURE.Value
    DIR_USERS = cboUSERS.text
    
    If lsShDirList.ListIndex = -1 Then Exit Sub
    
    SetAttr SHARES(Val(lsDS.List(lsShDirList.ListIndex))) & "\" & Longbow.SecurityFile, vbNormal
        
    Open SHARES(Val(lsDS.List(lsShDirList.ListIndex))) & "\" & Longbow.SecurityFile For Output As #44
        Print #44, "Users=" & DIR_USERS$
        Print #44, "Domain=" & DIR_DOMAIN$
        If DIR_SECURITY = 1 Then x$ = "yes" Else x$ = "no"
        Print #44, "Secure=" & x$
        If DIR_READ = 1 Then x$ = "yes" Else x$ = "no"
        Print #44, "Read=" & x$
        If DIR_EXECUTE = 1 Then x$ = "yes" Else x$ = "no"
        Print #44, "Execute=" & x$
        If DIR_VIEW = 1 Then x$ = "yes" Else x$ = "no"
        Print #44, "DirView=" & x$
        Print #44, "Write=yes"
    Close 44
    SetAttr SHARES(Val(lsDS.List(lsShDirList.ListIndex))) & "\" & Longbow.SecurityFile, vbHidden + vbSystem
End Sub

Private Sub cmdCLEAR_Click()
    txtDOMAIN.text = ""
    chkSECURE.Value = 0
    chkREAD.Value = 0
    chkDIRVIEW.Value = 0
    chkEXECUTE.Value = 0
    cboUSERS.text = ""
End Sub

Private Sub cmdDELETE_Click()
    On Error GoTo cmddeleteerr
    If lsShDirList.ListIndex = -1 Then Exit Sub

    
    ' Delete the security file for that folder
    
    ' Set file attributes to normal to make sure it can be deleted
    SetAttr SHARES(Val(lsDS.List(lsShDirList.ListIndex))) & "\" & Longbow.SecurityFile, vbNormal
    
    Kill SHARES(Val(lsDS.List(lsShDirList.ListIndex))) & "\" & Longbow.SecurityFile
    

    
    frmDirShare.ShowMe
    
    Exit Sub
cmddeleteerr:
    m_main.ShowSvEr "Error Deleting Shared Directory"
    SHARES(Val(lsDS.List(lsShDirList.ListIndex))) = ""
    SHARESL(Val(lsDS.List(lsShDirList.ListIndex))) = ""
    frmDirShare.ShowMe
End Sub

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Private Sub PLINKYCommand4_Click()
    Dim e$
    Dim t As Long
    e$ = Me.Caption
    Me.Caption = "Modifying Security Data For Recursed Folders"
    rec.Path = SHARES(Val(lsDS.List(lsShDirList.ListIndex)))
    For t = 0 To rec.ListIndex
        'AddShare <NEW SHARE NAME>,<REPLICATE FROM>
        AddShare rec.List(t), SHARES(Val(lsDS.List(lsShDirList.ListIndex)))
    Next t
    Me.Caption = e$
    frmDirShare.ShowMe
End Sub

Private Sub Command2_Click()
    Dim t As Integer
    On Error GoTo CMD2CLICK
    If txtNEWDIR.text = "" Then Exit Sub
    If IsDir(txtNEWDIR.text) = 0 Then Exit Sub
    'If Right$(txtNEWDIR.text, 1) <> "\" Then txtNEWDIR.text = txtNEWDIR.text & "\"
    ' Add To The Shares List
    For t = 0 To 700
        If SHARES(t) = txtNEWDIR.text Then GoTo NOTTOPLEVEL

    Next t
    
    
    For t = 0 To 700
        If SHARES(t) = "" Then
            SHARES(t) = txtNEWDIR.text
            SHARESL(t) = "TOPLEVEL"
            Exit For
        End If
    Next t
    
NOTTOPLEVEL:
    Open txtNEWDIR.text & "\" & Longbow.SecurityFile For Output As #44
        Print #44, "Users=any"
        Print #44, "Domain=Unnamed"
        Print #44, "Secure=yes"
        Print #44, "Read=yes"
        Print #44, "Execute=no"
        Print #44, "DirView=no"
        Print #44, "Write=yes"
    Close 44
    Me.ShowMe
    Exit Sub
CMD2CLICK:
    m_main.ShowSvEr "Error Creating New Share"
    Me.ShowMe
End Sub

Private Sub Command3_Click()
    Dim t As Integer
    Dim l As Long
    'On Error GoTo CMD3CLICK
    If txtNEWDIR.text = "" Then Exit Sub
    If IsDir(txtNEWDIR.text) = 0 Then Exit Sub
    'If Right$(txtNEWDIR.text, 1) <> "\" Then txtNEWDIR.text = txtNEWDIR.text & "\"
    ' Add To The Shares List
    For t = 0 To 700
        If SHARES(t) = txtNEWDIR.text Then GoTo NOTTHETOPLEVEL
    Next t
    
    
    For t = 0 To 700
        If SHARES(t) = "" Then
            SHARES(t) = txtNEWDIR.text
            SHARESL(t) = "TOPLEVEL"
            Exit For
        End If
    Next t
    
NOTTHETOPLEVEL:
    Open txtNEWDIR.text & "\" & Longbow.SecurityFile For Output As #44
        Print #44, "Users=any"
        Print #44, "Domain=Unnamed"
        Print #44, "Secure=yes"
        Print #44, "Read=yes"
        Print #44, "Execute=no"
        Print #44, "DirView=no"
        Print #44, "Write=yes"
    Close 44
    ' NOW MAKE SHARE DETAILS UP FOR DIRECTORIES BELOW THIS ONE
    rec.Path = txtNEWDIR.text
    'Debug.Print rec.Path
    For t = 0 To rec.ListCount - 1
        AddShare rec.List(t), txtNEWDIR.text
    Next t
    
    For t = 0 To 700
        If LTrim$(SHARESL(t)) = LTrim$(txtNEWDIR.text) Then
            ' Modify this locations security file to match the current one
            SetAttr SHARES(l) & "\" & Longbow.SecurityFile, vbNormal
            Kill SHARES(l) & "\" & Longbow.SecurityFile
            FileCopy txtNEWDIR.text & "\" & Longbow.SecurityFile, SHARES(l) & "\" & Longbow.SecurityFile
            SetAttr SHARES(l) & "\" & Longbow.SecurityFile, vbSystem + vbHidden
        End If
    Next t

    
    
    
    Me.ShowMe
    Exit Sub
CMD3CLICK:
    m_main.ShowSvEr "Error Creating New Share"
    Me.ShowMe

End Sub

Private Sub Command4_Click()
    Dim x$, e$
    Dim t As Integer
    Dim l As Integer
    cmdADDUPD_Click
    e$ = Me.Caption
    Me.Caption = "Updating Recursed Directories"
    If lsShDirList.ListIndex = -1 Then Exit Sub
    
    x$ = SHARES(Val(lsDS.List(lsShDirList.ListIndex)))
        
    For t = 0 To 700
        If LTrim$(SHARESL(t)) = LTrim$(x$) Then
            'Debug.Print x$, SHARESL(t), SHARES(t) & "\" & Longbow.SecurityFile
            ' Modify this locations security file to match the current one
            SetAttr SHARES(t) & "\" & Longbow.SecurityFile, vbNormal
            Kill SHARES(t) & "\" & Longbow.SecurityFile
            FileCopy x$ & "\" & Longbow.SecurityFile, SHARES(t) & "\" & Longbow.SecurityFile
            SetAttr SHARES(t) & "\" & Longbow.SecurityFile, vbSystem + vbHidden
        End If
    Next t
    Me.Caption = e$
    Exit Sub
CMD4ERR:
    ShowSvEr "Error Updating Recursed Directories"
    Me.Caption = e$
End Sub

Private Sub lsShDirList_Click()
    On Error GoTo LSBIGERR
    Dim t As Integer
    Dim f$
    Dim DIR_DOMAIN$, DIR_USERS$, DIR_READ%, DIR_WRITE%, DIR_VIEW%, DIR_EXECUTE%, DIR_SECURITY%
    
    DIR_READ = 1: DIR_WRITE = 1: DIR_VIEW = 1: DIR_EXECUTE = 1: DIR_SECURITY = 1
        
    Open SHARES(Val(lsDS.List(lsShDirList.ListIndex))) & "\" & Longbow.SecurityFile For Input As #44
         Do Until EOF(44)
            Line Input #44, f$
            f$ = LCase$(f$)
            Select Case f$
               Case "read=no"
                  DIR_READ = 0
               Case "write=no"
                  DIR_WRITE = 0
               Case "dirview=no"
                  DIR_VIEW = 0
               Case "execute=no"
                  DIR_EXECUTE = 0
               Case "secure=no"
                  DIR_SECURITY = 0
            End Select
            If Left$(f$, 6) = "domain" Then DIR_DOMAIN$ = Right$(f$, Len(f$) - 7)
            If Left$(f$, 5) = "users" Then DIR_USERS$ = Right$(f$, Len(f$) - 6)
         Loop
         Close 44
                  
         chkREAD.Value = DIR_READ
         chkSECURE.Value = DIR_SECURITY
         chkDIRVIEW.Value = DIR_VIEW
         chkEXECUTE.Value = DIR_EXECUTE
         txtDOMAIN.text = DIR_DOMAIN$
         cboUSERS.text = DIR_USERS$

    Exit Sub
LSBIGERR:
    Close 44
    ShowSvEr "Security File Unloadable"
End Sub

Public Sub AddShare(newshare As String, copyfrom As String) 'rec.List(t), SHARES(Val(lsDS.List(lsShDirList.ListIndex)))
    Dim t As Integer
    On Error GoTo AddShareError
    For t = 0 To 700
        If SHARES(t) = "" Then
            SHARES(t) = newshare
            SHARESL(t) = copyfrom
            FileCopy copyfrom & "\" & Longbow.SecurityFile, newshare & "\" & Longbow.SecurityFile
            Exit For
        End If
    Next t
    Exit Sub
AddShareError:
    MsgBox "Error Adding New Directory To Share", vbExclamation, "Longbow Server"

End Sub
