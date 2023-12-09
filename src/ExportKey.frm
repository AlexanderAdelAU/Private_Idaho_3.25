VERSION 5.00
Begin VB.Form frmExportKey 
   Caption         =   "Exported Key"
   ClientHeight    =   4605
   ClientLeft      =   2010
   ClientTop       =   2370
   ClientWidth     =   7290
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   2310
      TabIndex        =   2
      Top             =   4050
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   4050
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3675
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmExportKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  Clipboard.Clear
  Clipboard.SetText Text1.Text
  Response = MsgBox("Copied to the clipboard", vbInformation)
  
End Sub

Private Sub Form_Activate()
  frmExportKey.btnOK.SetFocus
End Sub

