VERSION 5.00
Begin VB.Form frmMixmasterHome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MIXMASTER Location"
   ClientHeight    =   4815
   ClientLeft      =   2760
   ClientTop       =   1995
   ClientWidth     =   6300
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
   Icon            =   "MixmasterHome.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4815
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   780
      TabIndex        =   5
      Top             =   1680
      Width           =   3315
   End
   Begin VB.DriveListBox Drive1 
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
      Left            =   780
      TabIndex        =   4
      Top             =   1200
      Width           =   3315
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   4230
      Pattern         =   "*.exe"
      TabIndex        =   3
      Top             =   1260
      Width           =   1515
   End
   Begin VB.CommandButton OK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4230
      Width           =   945
   End
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   4230
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "MixmasterHome.frx":000C
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lblIntro 
      BackStyle       =   0  'Transparent
      Caption         =   $"MixmasterHome.frx":0652
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Index           =   3
      Left            =   810
      TabIndex        =   2
      Top             =   150
      Width           =   5025
   End
End
Attribute VB_Name = "frmMixmasterHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim msg As String
On Error GoTo Drive1Error
Dir1.Path = Drive1.Drive
Exit Sub

Drive1Error:
Beep
If Err.Number = 68 Or Err.Number = 71 Then
    msg = "Error #" & Str$(Err) & " No floppy in drive!"
    MsgBox msg, vbExclamation, App.Title
Else
    msg = "Error #" & Str$(Err)
End If
Err.Clear
End Sub

Private Sub Form_Load()
On Error Resume Next
Screen.MousePointer = vbDefault
Dir1.Path = gMixPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMixmasterHome = Nothing
End Sub

Private Sub OK_Click()
Dim SectionName As String
SectionName = "Remailer Info"
gMixPath = Dir1.Path
WriteProfile SectionName, "MixmasterPath", gMixPath
Screen.MousePointer = vbDefault
Unload Me
End Sub



