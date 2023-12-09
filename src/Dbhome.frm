VERSION 5.00
Begin VB.Form frmDBHome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Location"
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
   Icon            =   "Dbhome.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4815
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   2505
      Left            =   810
      TabIndex        =   5
      Top             =   1470
      Width           =   3315
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   4230
      Pattern         =   "*.mdb"
      TabIndex        =   4
      Top             =   1020
      Width           =   1515
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   810
      TabIndex        =   2
      Top             =   1020
      Width           =   3285
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
      Left            =   1680
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
      Left            =   2730
      TabIndex        =   1
      Top             =   4230
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "Dbhome.frx":000C
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Private Idaho can't find the PGP directory.   Select the location where the database can be found and then press ok."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   810
      TabIndex        =   3
      Top             =   300
      Width           =   4905
   End
End
Attribute VB_Name = "frmDBHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private DriveNum As Integer


Private Sub Cancel_Click()
End
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
If Err = 68 Or Err = 71 Then
    msg = "Error #" & Str$(Err) & " No floppy in drive!"
    MsgBox msg, vbExclamation, App.Title
Else
    msg = "Error #" & Str$(Err)
End If
On Error GoTo 0
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbArrow
Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, _
                        SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub OK_Click()
'DbPath = Dir1.Path
SaveSetting App.Title, PATHS, DBLOCATION, Dir1.Path
Unload Me
'Set frmDBHome = Nothing
Screen.MousePointer = vbDefault
End Sub



