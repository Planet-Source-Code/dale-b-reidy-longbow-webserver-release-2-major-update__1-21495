VERSION 5.00
Begin VB.Form frmSvErr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Longbow Server Message"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   750
      TabIndex        =   1
      Top             =   75
      Width           =   3765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1725
      TabIndex        =   0
      Top             =   675
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   75
      Picture         =   "frmSvErr.frx":0000
      Top             =   150
      Width           =   630
   End
End
Attribute VB_Name = "frmSvErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Public Sub ShowMe(text As String)
    Text1.text = text$
    Me.Visible = True
End Sub
