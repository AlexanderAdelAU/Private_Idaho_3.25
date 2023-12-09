VERSION 5.00
Begin VB.Form frmNymServerStats 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nym Server Stats"
   ClientHeight    =   3390
   ClientLeft      =   1740
   ClientTop       =   1440
   ClientWidth     =   6960
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
   LinkTopic       =   "Form31"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   6960
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Next"
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
      Left            =   4545
      TabIndex        =   6
      Top             =   2850
      Width           =   975
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
      Left            =   5640
      TabIndex        =   5
      Top             =   2850
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "name"
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "latency"
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
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "up time"
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
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a nym server to use and click next.  Note only nym.alias format are support at present."
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6315
   End
End
Attribute VB_Name = "frmNymServerStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Nym.Server = List1.List(List1.ListIndex)
    If Nym.Server = "" Then
        Beep
        Exit Sub
    End If
    Nym.Server = Mid$(Nym.Server, 1, InStr(Nym.Server, " ") - 1)
   ' gIsNewNym = IsNewNym(Nym.Server)
    Unload Me
    frmCreateNymStep2.Show
End Sub

Private Sub Command2_Click()
    SetWindowPos frmNymServerStats.hWnd, -2, 0, 0, 0, 0, &H50 'This will put "Always Visible" off!
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Win As New CWindow
    Win.OnTop(Me) = True
    Win.Center Me, Null
    
    'If gShowNymStatus = 1 Then
       ' Command1.Visible = False
       ' Me.Caption = "Nym Server Stats from Remailer info."
       ' Label1.Caption = FileDateTime(App.Path + "\remailer.txt")
    'End If
    Load frmRemailerList
    frmRemailerList.ClearRemailerSort
    'This will just sort the list for cypher punk remailers with NewNym
    gSortRemailer.newnym = 1
    frmRemailerList.SortRemailersByType
    frmRemailerList.FillNymList
    Unload frmRemailerList
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmNymServerStats = Nothing
End Sub
