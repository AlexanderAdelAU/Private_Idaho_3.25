VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmInputBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Box"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblPrompt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   660
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Input Box.frx":0000
      Top             =   90
      Width           =   5325
   End
   Begin VB.CommandButton btnButton 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   4860
      TabIndex        =   2
      Top             =   2610
      Width           =   1065
   End
   Begin VB.CommandButton btnButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   3630
      TabIndex        =   1
      Top             =   2610
      Width           =   1065
   End
   Begin VB.TextBox txtInputText 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   630
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2070
      Width           =   5325
   End
   Begin Threed.SSCheck chkShowPassword 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   131074
      BackStyle       =   1
      Caption         =   "Show Typing"
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   60
      Picture         =   "Input Box.frx":0006
      Top             =   90
      Width           =   570
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnButton_Click(Index As Integer)
If Index = 1 Then txtInputText = ""
Me.Hide
End Sub

Private Sub chkShowPassword_Click(Value As Integer)
If Value = True Then
    txtInputText.PasswordChar = ""
Else
    txtInputText.PasswordChar = "*"
End If
End Sub

