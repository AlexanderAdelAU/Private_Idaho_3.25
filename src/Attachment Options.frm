VERSION 5.00
Begin VB.Form frmAttachmentOptions 
   Caption         =   "Attachment Options"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1515
      Left            =   930
      TabIndex        =   3
      Top             =   2610
      Width           =   5115
      Begin VB.OptionButton Option1 
         Caption         =   "Save it to disk."
         Height          =   435
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   750
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Open the attachment."
         Height          =   435
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   330
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   5130
      TabIndex        =   2
      Top             =   4290
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   345
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   4290
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File to be opened: "
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   150
      Width           =   1365
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "File to open or save: "
      Height          =   315
      Left            =   2430
      TabIndex        =   6
      Top             =   150
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   1935
      Left            =   840
      TabIndex        =   0
      Top             =   570
      Width           =   5235
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   60
      Picture         =   "Attachment Options.frx":0000
      Top             =   30
      Width           =   675
   End
End
Attribute VB_Name = "frmAttachmentOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public iRes As Integer
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    If Option1.Item(0).Value Then
        iRes = vbYes
    Else
        iRes = vbNo
    End If
Else
    iRes = vbCancel
End If
Unload Me

End Sub

Private Sub Form_Load()
Dim msg As String
    
    msg = "This file is encrypted.  Would you like to view or save the attachment? " & vbCrLf & vbCrLf
    msg = msg & "Select your option, and press OK." & vbCrLf & vbCrLf
    msg = msg & "Note - if you launch or open an encrypted message ( a file with an .asc extension), then a decrypted copy WILL "
    msg = msg & "be located in your system's temporary directory " & "(" & TempPathLocation & "), " & "or in the FileSafe directory."
    msg = msg & "You should delete it if you do not wish it to be seen by others. "
    Label1 = msg
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set frmAttachmentOptions = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
'If Option1.Item(1).Value = True Then
'End If
End Sub

Private Sub option_Click(Index As Integer)

End Sub
