VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMessageType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Encrypted As Integer
Public Signed As Integer
Public DetachedSignature As Integer
Public Unknown As Integer
Public Key As Integer
Public Decrypted As Boolean
Public DecryptedText As Integer

Private Sub Class_Initialize()
Encrypted = 1
DecryptedText = 2
Signed = 3
DetachedSignature = 4
Key = 5
Unknown = 6
Decrypted = False
End Sub
