VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "spgpContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=======================
Public Sign As Long
Public KeyEncrypt As Long
'Public SignAlg As Long
Public ConventionalEncrypt As Long
Public Compress As Long
Public Armor As Long
Public TextMode As Long
Public Clear As Long
Public Comment As String
Public FileIn As String
Public FileOut As String
Public SelectPrivateKeys As Boolean
Public MIME As Long
Public EyesOnly As Long
Public MIMESeparator As String

    
'=============================================
'Main declarations
'===============================================
' note that while some of these constants have the same name as
' variables found in the pgp *.h files, they do not all have the same
' values -- these are not meant to be translations of those files!


'public key algorithms
Private Const PGPPublicKeyAlgorithm_Invalid = 0 'this (0) may not be correct, but how to test it? anyone got an invalid algorithm laying around?
Private Const PGPPublicKeyAlgorithm_RSA = 1
Private Const PGPPublicKeyAlgorithm_RSAEncryptOnly = 2
Private Const PGPPublicKeyAlgorithm_RSASignOnly = 3
Private Const PGPPublicKeyAlgorithm_ElGamal = 4 'A.K.A. Diffie-Hellman
Private Const PGPPublicKeyAlgorithm_DSA = 5


'symmetric ciphers
Private m_PGPCipherAlgorithm As Integer

'hashing algorithms
Private m_PGPHashAlgorithm As Integer


'trust levels
Private Const PGPKeyTrust_Undefined = 0
Private Const PGPKeyTrust_Unknown = 1
Private Const PGPKeyTrust_Never = 2
Private Const PGPKeyTrust_Marginal = 5
Private Const PGPKeyTrust_Complete = 6
Private Const PGPKeyTrust_Ultimate = 7

'Public Algorithm as
'validity levels
Private Const PGPValidity_Unknown = 0
Private Const PGPValidity_Invalid = 1
Private Const PGPValidity_Marginal = 2
Private Const PGPValidity_Complete = 3

'analysis results
Private Const PGPAnalyze_Encrypted = 0            ' Encrypted message
Private Const PGPAnalyze_Signed = 1               ' Signed message
Private Const PGPAnalyze_DetachedSignature = 2    ' Detached signature
Private Const PGPAnalyze_Key = 3                  ' Key data
Private Const PGPAnalyze_Unknown = 4              ' Non-pgp message

' these are for spgpAnalyzeEx only
Private Const PGPAnalyze_EncryptedConventional = 5 ' Like it says
Private Const PGPAnalyze_EncryptedNoKeys = 6      ' Key-encrypted to keys not on local ring

'signature status
Private Const SIGNED_GOOD = 0
Private Const SIGNED_NOT = 1
Private Const SIGNED_BAD = 2
Private Const SIGNED_NO_KEY = 3

'preferences
Private Const PGPPref_PublicKeyring = 0
Private Const PGPPref_PrivateKeyring = 1
Private Const PGPPref_RandomSeedFile = 2
Private Const PGPPref_DefaultKeyID = 3   'not supported (encoded?)

Private mSignKeyPass As String
Private mCryptKeyID As String
Private mConventionalPass As String
Private mSignKeyID As String
    
Public Sub Initialise()


FileIn = ""
FileOut = ""
Sign = 0
mSignKeyPass = ""
KeyEncrypt = 0
ConventionalEncrypt = 0
Armor = 0
TextMode = 0
Clear = 0
mCryptKeyID = ""
mConventionalPass = ""
Comment = ""
SelectPrivateKeys = False
mSignKeyID = ""
MIME = 0
EyesOnly = 0
MIMESeparator = ""
Compress = 0
End Sub

Private Sub Class_Initialize()
Dim HashAlg As New cPGPHashAlgorithm
Dim CipherAlg As New cPGPCipherAlgorithm

m_PGPHashAlgorithm = HashAlg.PGPPublicKeyAlgorithm_Default
m_PGPCipherAlgorithm = CipherAlg.PGPCipherAlgorithm_Default
Call Initialise
End Sub

Public Property Get SignKeyID() As String
SignKeyID = mSignKeyID & Chr(0)
End Property

Public Property Let SignKeyID(ByVal PGPKeyID As String)
Dim pos1 As Integer
Dim pos2 As Integer
If PGPKeyID <> "" Then
    '---------------------------------------------
    'strip out to keep only the bare e-mail address with <...>
    '---------------------------------------------
    PGPKeyID = PGPKeyID
    pos1 = InStr(PGPKeyID, "<")
    pos2 = InStr(PGPKeyID, ">")
    If pos1 > 1 And pos2 > pos1 Then
         mSignKeyID = Mid(PGPKeyID, pos1, pos2 - pos1 + 1)
    Else
         mSignKeyID = "<" & PGPKeyID & ">"
    End If
End If
End Property

Public Property Get CryptKeyID() As String
CryptKeyID = mCryptKeyID & Chr(0)
End Property

Public Property Let CryptKeyID(ByVal PGPKeyID As String)
Dim pos1 As Integer
Dim pos2 As Integer
If PGPKeyID <> "" Then
    '---------------------------------------------
    'strip out to keep only the bare e-mail address
    '---------------------------------------------
    pos1 = InStr(PGPKeyID, "<")
    pos2 = InStr(PGPKeyID, ">")
    If pos1 >= 1 And pos2 > pos1 Then
         mCryptKeyID = Mid(PGPKeyID, pos1, pos2 - pos1 + 1)
    Else
         mCryptKeyID = "<" & PGPKeyID & ">"
    End If
End If
End Property

Public Property Get ConventionalPass() As String
ConventionalPass = mConventionalPass & Chr(0)
End Property

Public Property Let ConventionalPass(ByVal pPass As String)
mConventionalPass = pPass
End Property


Public Property Get SignKeyPass() As String
SignKeyPass = mSignKeyPass & Chr(0)
End Property

Public Property Let SignKeyPass(ByVal vNewpass As String)
mSignKeyPass = vNewpass
End Property
Public Property Get PGPHashAlgorithm() As Integer
PGPHashAlgorithm = m_PGPHashAlgorithm
End Property
Public Property Get PGPCipherAlgorithm() As Integer
PGPCipherAlgorithm = m_PGPCipherAlgorithm
End Property
Public Property Let PGPHashAlgorithm(mValue As Integer)
 m_PGPHashAlgorithm = mValue
End Property

Public Property Let PGPCipherAlgorithm(mValue As Integer)
 m_PGPCipherAlgorithm = mValue
End Property
