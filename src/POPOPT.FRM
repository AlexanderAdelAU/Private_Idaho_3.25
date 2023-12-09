VERSION 5.00
Begin VB.Form Form28 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POP Options"
   ClientHeight    =   1920
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   4920
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
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   4920
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete PGP messages from server after download"
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
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim SectionName As String
    Dim KeyName As String
    Dim KeyValue As String
    SectionName = "Options"
    KeyName = "ServerDelete"
    If Check1.Value = 0 Then
        KeyValue = "0"
    Else
        KeyValue = "1"
    End If
    WriteProfile SectionName, KeyName, KeyValue
    Unload Form28
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim SectionName As String
    Dim tmpstr As String
    
    SectionName = "Options"
    tmpstr = ReadProfile(SectionName, "ServerDelete")
    If tmpstr = "1" Or tmpstr = "" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form28 = Nothing
End Sub
