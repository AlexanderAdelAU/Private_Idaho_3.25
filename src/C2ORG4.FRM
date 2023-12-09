VERSION 5.00
Begin VB.Form frmCreateNymStep4 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nym - Step 4"
   ClientHeight    =   4125
   ClientLeft      =   7065
   ClientTop       =   390
   ClientWidth     =   5025
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
   LinkTopic       =   "Form20"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   5025
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   $"C2ORG4.frx":0000
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
         Height          =   1065
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   4155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   $"C2ORG4.frx":0091
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
         Height          =   885
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4155
      End
   End
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
      Left            =   2640
      TabIndex        =   0
      Top             =   3720
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
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"C2ORG4.frx":0123
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
      Height          =   885
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmCreateNymStep4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

           
    'If gRemailerType = STANDARD_EMAIL Then
     '   frmCreateNymStep6.Show
     '   Unload Me
   ' Else
        If Not gRemailerType = STANDARD_EMAIL Then
            If Not gRemailerType = REMAILER_CYPHERPUNK Then
                MsgBox "Only Cypherpunk type remailers can be used for reply blocks.", 16, gPiStr
                Exit Sub
            End If
        End If
        frmCreateNymStep5.Show
        Unload Me
   ' End If
End Sub

Private Sub Command2_Click()
    gIsNewNym = False
    gNymState = gNYM_IDLE
    Unload Me
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
    Win.Center Me, Null
    Win.OnTop(Me) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCreateNymStep4 = Nothing
End Sub
