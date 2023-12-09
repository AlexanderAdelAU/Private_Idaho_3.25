VERSION 5.00
Begin VB.Form frmRemailerCodes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remailer Codes"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   405
      Left            =   3510
      TabIndex        =   1
      Top             =   2760
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2025
      Left            =   510
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "RemailerCodes.frx":0000
      Top             =   510
      Width           =   4275
   End
End
Attribute VB_Name = "frmRemailerCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "# = response in less than 5 minutes." & vbCrLf
Text1.Text = Text1.Text & "* = response in less than 1 hour." & vbCrLf
Text1.Text = Text1.Text & "+ = response in less than 4 hours." & vbCrLf
Text1.Text = Text1.Text & "- = response in less than 24 hours." & vbCrLf
Text1.Text = Text1.Text & ". = response in less than 2 days." & vbCrLf
Text1.Text = Text1.Text & "_ = response in more than 2 days." & vbCrLf

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmRemailerCodes = Nothing
End Sub

