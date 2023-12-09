VERSION 5.00
Begin VB.Form frmPOPPassword 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POP e-mail login"
   ClientHeight    =   2100
   ClientLeft      =   1440
   ClientTop       =   1665
   ClientWidth     =   5100
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   5100
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2250
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1020
      Width           =   4125
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter your e-mail account password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   270
      TabIndex        =   1
      Top             =   150
      Width           =   4095
   End
End
Attribute VB_Name = "frmPOPPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If gPassPhrase <> "1" And gPassPhrase <> "2" Then
        MailConnector.AccountPassword = "@_@"
    End If
    If gPassPhrase = "2" Then
        End
    End If
    If gPassPhrase = "1" Then
        gPassPhrase = ""
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error GoTo PassErr
    frmPOPPassword.Hide
    If Text2.Text = "" Then
        Beep
        Text2.Text = "Can't be blank"
    Else
        If gPassPhrase = "1" Or gPassPhrase = "2" Then
            gPassPhrase = Text2.Text
        Else
            MailConnector.AccountPassword = Text2.Text
            frmMain.POP1.User = MailConnector.AccountName
            frmMain.POP1.Password = MailConnector.AccountPassword
        End If
        Unload Me
    End If
    
    Exit Sub

PassErr:
    MsgBox Err.Description & " (Password Error..)"
    Err.Clear
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
    Win.Center Me, Null
    Win.OnTop(Me) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPOPPassword = Nothing
End Sub


