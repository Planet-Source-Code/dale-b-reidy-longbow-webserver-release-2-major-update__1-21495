VERSION 5.00
Begin VB.Form frmIntUpdate 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Integrity Update"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblIntUpdate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   375
      TabIndex        =   1
      Top             =   0
      Width           =   4965
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   4875
      Picture         =   "frmIntUpdate.frx":0000
      Top             =   450
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   75
      Picture         =   "frmIntUpdate.frx":0427
      Top             =   450
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIntUpdate.frx":084E
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Left            =   900
      TabIndex        =   0
      Top             =   450
      Width           =   3915
   End
End
Attribute VB_Name = "frmIntUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
