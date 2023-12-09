VERSION 5.00
Begin VB.Form frmCreateNymStep3 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nym - Step 3"
   ClientHeight    =   5430
   ClientLeft      =   7065
   ClientTop       =   390
   ClientWidth     =   6105
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
   LinkTopic       =   "Form19"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   6105
   Visible         =   0   'False
   Begin VB.OptionButton optReplyBlock 
      Caption         =   "Use Email Reply Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   14
      Top             =   1920
      Value           =   -1  'True
      Width           =   2385
   End
   Begin VB.OptionButton optReplyBlock 
      Caption         =   "Use NewsGroup Reply Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3060
      TabIndex        =   13
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Frame fReplyBlock 
      Caption         =   "Use your Emal Address as your Reply Block"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   0
      Left            =   540
      TabIndex        =   10
      Top             =   2400
      Width           =   5055
      Begin VB.TextBox txtEmailAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Width           =   3525
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Your real email address"
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
         Height          =   225
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   2535
      End
   End
   Begin VB.Frame fReplyBlock 
      Caption         =   "Use Newsgroup as your Reply Block"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Index           =   1
      Left            =   540
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtNewsGroupEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Text            =   "mail2news@anon.lcs.mit.edu"
         Top             =   510
         Width           =   3645
      End
      Begin VB.TextBox txtNewsGroupEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Text            =   "alt.anonymous.messages"
         Top             =   1140
         Width           =   3645
      End
      Begin VB.TextBox txtNewsGroupEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Text            =   "Something you recognise"
         Top             =   1710
         Width           =   3645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mail2News Service"
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
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   2535
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NewsGroup Name"
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
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   2535
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reply Block Subject"
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
         Height          =   225
         Index           =   13
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   2535
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
      Left            =   3570
      TabIndex        =   0
      Top             =   4740
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
      Left            =   4620
      TabIndex        =   1
      Top             =   4740
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"C2ORG3.frx":0000
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
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   90
      Width           =   5505
   End
End
Attribute VB_Name = "frmCreateNymStep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Nym.EmailAddress = txtEmailAddress.Text
    
    Nym.NewsGroupReplyEmail = txtNewsGroupEmail(0)
    Nym.NewsGroupReplyGroup = txtNewsGroupEmail(1)
    Nym.NewsGroupReplySubject = txtNewsGroupEmail(2)
      
    If optReplyBlock(1).Value = True Then
        Nym.UseNewsGroupReply = True
    Else
         Nym.UseNewsGroupReply = False
    End If
   
    frmCreateNymStep4.Show
    Unload Me
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
Set frmCreateNymStep3 = Nothing
End Sub

Private Sub optReplyBlock_Click(Index As Integer)
Select Case Index
    Case 0
        fReplyBlock(0).Visible = True
        fReplyBlock(1).Visible = False
        Nym.UseNewsGroupReply = False
    Case 1
        fReplyBlock(1).Visible = True
        fReplyBlock(0).Visible = False
        Nym.UseNewsGroupReply = True
End Select
End Sub
