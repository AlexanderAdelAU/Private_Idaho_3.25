VERSION 5.00
Begin VB.Form frmFileEncryptionOption 
   Caption         =   "File Encryption Options"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4830
      TabIndex        =   1
      Top             =   3570
      Width           =   1215
   End
   Begin VB.Frame Option 
      Caption         =   "Option"
      Height          =   3255
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   5835
      Begin VB.OptionButton optFile 
         Caption         =   "Encrypt a existing file"
         Height          =   405
         Index           =   1
         Left            =   1170
         TabIndex        =   4
         Top             =   2340
         Width           =   3885
      End
      Begin VB.OptionButton optFile 
         Caption         =   "Encrypt the message area to a file"
         Height          =   405
         Index           =   0
         Left            =   1170
         TabIndex        =   3
         Top             =   1980
         Value           =   -1  'True
         Width           =   3885
      End
      Begin VB.Label Label1 
         Caption         =   $"File Encryption Options.frx":0000
         Height          =   1485
         Left            =   240
         TabIndex        =   2
         Top             =   450
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmFileEncryptionOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EncryptToFile As Boolean

Private Sub btnok_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If optFile(0).Value = True Then
    EncryptToFile = True
Else
    EncryptToFile = False
End If
End Sub
