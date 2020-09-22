VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim database$       ' Database data
Dim limit$              ' Limited database data
Dim databasefile$  ' Database Filename

Public Function OpenDatabase(databasefilename As String) As Integer
    
    ' Open the database
    
    Dim databasedir$
    Dim ffil1%, lofdb%
    
    Open databasefilename For Binary As #ffil1
        lofdb% = LOF(ffil1%)
        database$ = Space$(lofdb)
        Get #ffil1, , database$
    Close ffil1
    
End Function



