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
    Public Sign As Long
    Public Encrypt As Long
    Public SignAlg As Long
    Public ConventionalEncrypt As Long
    Public ConventionalAlg As Long
    Public Armor As Long
    Public Textmode As Long
    Public Clear As Long
    Public CryptKeyID As String
    Public ConventionalPass As String
    Public SignKeyID As String
    Public SignKeyPass As String
    Public Comment As String
    Public FileIn As String
    Public FileOut As String
    Public Selectprivatekeys As boolena
    
    
''@E CProblem

Public Sub Initialise()
FileIn = ""
FileOut = ""
Sign = 0
SignKeyPass = ""
Encrypt = 1
SignAlg = 0
ConventionalEncrypt = 0
ConventionalAlg = 1
Armor = 0
Textmode = 0
Clear = 0
CryptKeyID = ""
ConventionalPass = ""
Comment = ""
Selectprivatekeys = False
End Sub

Private Sub Class_Initialize()
FileIn = ""
FileOut = ""
Sign = 0
SignKeyPass = ""
Encrypt = 1
SignAlg = 0
ConventionalEncrypt = 0
ConventionalAlg = 1
Armor = 0
Textmode = 0
Clear = 0
CryptKeyID = ""
ConventionalPass = ""
Comment = ""
Selectprivatekeys = False
End Sub


