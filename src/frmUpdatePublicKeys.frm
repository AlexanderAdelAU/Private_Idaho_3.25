VERSION 5.00
Begin VB.Form frmUpdatePublicKeys 
   Caption         =   "Update Public Keys Option"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnNo 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3660
      TabIndex        =   1
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Okay"
      Height          =   315
      Left            =   3660
      TabIndex        =   0
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   60
      Picture         =   "frmUpdatePublicKeys.frx":0000
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUpdatePublicKeys.frx":0646
      Height          =   1815
      Left            =   900
      TabIndex        =   2
      Top             =   300
      Width           =   3795
   End
End
Attribute VB_Name = "frmUpdatePublicKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnNo_Click()
Unload Me
End Sub

Private Sub btnOkay_Click()
On Error GoTo BadPublicKeys
''UpdatePublicKeysFile
Unload Me
Exit Sub
BadPublicKeys:
    MsgBox Err.Description & " Can't update for some reason.  Check for the existence of your public key file!", vbCritical + vbApplicationModal, App.Title
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Load()
 Dim Win As New CWindow
Win.Center Me, Null
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUpdatePublicKeys = Nothing
End Sub
