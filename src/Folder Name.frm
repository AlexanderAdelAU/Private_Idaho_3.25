VERSION 5.00
Begin VB.Form frmFolderName 
   Caption         =   "Group or Folder Name"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCommand 
      Caption         =   "Cancel"
      Height          =   405
      Index           =   1
      Left            =   3390
      TabIndex        =   3
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "OK"
      Height          =   405
      Index           =   0
      Left            =   2310
      TabIndex        =   2
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txtFolderName 
      Height          =   345
      Left            =   420
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   3795
   End
   Begin VB.Label lblNamePrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Name"
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Top             =   240
      Width           =   4065
   End
End
Attribute VB_Name = "frmFolderName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FolderName As String
Private Sub btnCommand_Click(Index As Integer)
If Index = 0 Then
    FolderName = txtFolderName
Else
    'frmMain.Tree(giTreeIndex).CreateFolderName = ""
    FolderName = ""
End If
'Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
Win.Center Me, Null
txtFolderName = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set frmFolderName = Nothing
End Sub


