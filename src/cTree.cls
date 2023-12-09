VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NodeName As String
Public FolderId As Long
Public CancelClick As Boolean
Public CreateFolderName As String
Public FolderName As String
Public NodeID As Long

Public Sub Initialise()
NodeName = ""
FolderId = 0
CancelClick = False
CreateFolderName = ""
FolderName = ""
NodeID = 0
End Sub
