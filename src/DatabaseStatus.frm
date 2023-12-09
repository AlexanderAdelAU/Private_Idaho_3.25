VERSION 5.00
Begin VB.Form frmDatabaseStatus 
   Caption         =   "Database Status"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   2790
      TabIndex        =   1
      Top             =   900
      Width           =   885
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   3705
   End
End
Attribute VB_Name = "frmDatabaseStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Set frmDatabaseStatus = Nothing
End Sub
