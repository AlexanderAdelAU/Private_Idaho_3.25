VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "spgpParms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#If Win32 Then
'DefLng A-Z
'#Else
'DefInt A-Z
'#End If

''@B CProblem

'
    Public Sign As Long
    Public SignKeyPass As String
    
''@E CProblem


Public Property Get Buffer() As String
Buffer = BufferIn
End Property

Public Property Let Buffer(ByVal vNewValue As Integer)
mBufferSize = vNewValue
End Property

Public Sub Clear()
Sign = 0
SignKeyPass = ""
End Sub

Private Sub Class_Initialize()

End Sub
