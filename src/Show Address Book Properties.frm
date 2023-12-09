VERSION 5.00
Begin VB.Form frmShowAddressBookProperties 
   Caption         =   "Address Book Properties"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1830
      Width           =   975
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Top             =   960
      Width           =   3195
   End
   Begin VB.TextBox txtDisplayName 
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Top             =   390
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address: "
      Height          =   315
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name: "
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   420
      Width           =   1485
   End
End
Attribute VB_Name = "frmShowAddressBookProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnok_Click()
Unload Me
End Sub



Private Sub Form_Load()
Dim Win As New CWindow
'Win.OnTop(Me) = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmShowAddressBookProperties = Nothing
End Sub


Private Sub txtDisplayName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub
