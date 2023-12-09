VERSION 5.00
Begin VB.Form Form26 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple secret keys"
   ClientHeight    =   2850
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5160
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   5160
   Begin VB.ListBox List1 
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
      Height          =   1590
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4815
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
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
      Left            =   4080
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select the Private Key you'd like to use."
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    gPGPKeyID = List1.List(List1.ListIndex)
    Unload Form26
End Sub

Private Sub Command2_Click()
    Label1.Caption = ""
    Unload Me
    gCancelAction = True
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim tmpstr As String
Dim SectionName As String
    
    SectionName = "PGP Options"
    For i = 1 To 8
        tmpstr = ReadProfile(SectionName, "User ID " + Format(i))
        If tmpstr <> "" Then
            List1.AddItem tmpstr
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form26 = Nothing
End Sub

Private Sub List1_DblClick()
    gPGPKeyID = List1.List(List1.ListIndex)
    Unload Form26
End Sub

