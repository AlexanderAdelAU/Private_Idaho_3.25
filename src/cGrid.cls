VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SelectedCol As Integer
Public SelectedColToSort As Integer
Public SelectedRow As Integer
Public SelectedQuery As String
Public RowSelection As Integer


Private Sub Class_Initialize()
SelectedColToSort = 5
Initialise
End Sub

Public Sub Initialise()
SelectedCol = 5
SelectedRow = 1
SelectedQuery = ""
RowSelection = 0

End Sub


