VERSION 5.00
Begin VB.Form frmScanStatus 
   Caption         =   "Email Scan Status..."
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   180
      Picture         =   "frmScanStatus.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing Message Number: "
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   900
      Width           =   2235
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes Transferred: "
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   4
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label lblBytesTransferred 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while PI scans for your messages.   If the files are large this could take a while..."
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   60
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   900
      Width           =   1575
   End
End
Attribute VB_Name = "frmScanStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Win As New CWindow
    Win.Center Me, Null
    Win.OnTop(Me) = True
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmScanStatus = Nothing
End Sub

