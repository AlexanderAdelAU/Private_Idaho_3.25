VERSION 5.00
Begin VB.Form frmAdvancedEncryptionOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advanced Encryption Options"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bOkay 
      Caption         =   "Ok"
      Height          =   405
      Left            =   4290
      TabIndex        =   2
      Top             =   4170
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "HASH  Algorithm"
      Height          =   1755
      Left            =   330
      TabIndex        =   1
      Top             =   2070
      Width           =   5205
      Begin VB.OptionButton optHash 
         Caption         =   "RipeMD - 160"
         Height          =   405
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   900
         Width           =   1575
      End
      Begin VB.OptionButton optHash 
         Caption         =   "SHA1"
         Height          =   405
         Index           =   1
         Left            =   1860
         TabIndex        =   7
         Top             =   900
         Width           =   975
      End
      Begin VB.OptionButton optHash 
         Caption         =   "MD5"
         Height          =   405
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   900
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "This is choses the HASH algorithm which will be used in PGP.  The default is MD5."
         Height          =   495
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   450
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cipher Algorithm"
      Height          =   1485
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5115
      Begin VB.OptionButton optCipher 
         Caption         =   "CAST5"
         Height          =   405
         Index           =   2
         Left            =   3060
         TabIndex        =   5
         Top             =   840
         Width           =   1365
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "3DES"
         Height          =   405
         Index           =   1
         Left            =   1770
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optCipher 
         Caption         =   "IDEA"
         Height          =   405
         Index           =   0
         Left            =   390
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "This algorithm is used when you 'Conventionaly Encrypt' a message.  The default is IDEA."
         Height          =   495
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   270
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmAdvancedEncryptionOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bOkay_Click()
Unload Me
End Sub

Private Sub Form_Load()

Select Case vb2spgpContext.PGPHashAlgorithm
    Case 1
            optHash(0).Value = True
    Case 2
            optHash(1).Value = True
    Case 3
            optHash(2).Value = True
End Select

Select Case vb2spgpContext.PGPCipherAlgorithm
    Case 1
            optCipher(0).Value = True
    Case 2
            optCipher(1).Value = True
    Case 3
            optCipher(2).Value = True
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAdvancedEncryptionOptions = Nothing
End Sub

Private Sub optCipher_Click(Index As Integer)

Dim CipherAlg As New cPGPCipherAlgorithm

'vb2spgpContext.Initialise
Select Case Index
    Case 0
            vb2spgpContext.PGPCipherAlgorithm = CipherAlg.PGPCipherAlgorithm_IDEA
    Case 1
            vb2spgpContext.PGPCipherAlgorithm = CipherAlg.PGPCipherAlgorithm_3DES
    Case 2
            vb2spgpContext.PGPCipherAlgorithm = CipherAlg.PGPCipherAlgorithm_CAST5
End Select
End Sub

Private Sub optHash_Click(Index As Integer)
Dim HashAlg As New cPGPHashAlgorithm


Select Case Index
    Case 0
            vb2spgpContext.PGPHashAlgorithm = HashAlg.PGPPublicKeyAlgorithm_MD5
    Case 1
            vb2spgpContext.PGPHashAlgorithm = HashAlg.PGPHashAlgorithm_SHA
    Case 2
            vb2spgpContext.PGPHashAlgorithm = HashAlg.PGPHashAlgorithm_RIPEMD160
End Select
End Sub
