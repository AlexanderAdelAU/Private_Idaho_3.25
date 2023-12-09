VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmCreateNymStep6 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nym - Step 6"
   ClientHeight    =   4725
   ClientLeft      =   7050
   ClientTop       =   375
   ClientWidth     =   5610
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
   LinkTopic       =   "Form22"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   5610
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   300
      TabIndex        =   4
      Top             =   2220
      Width           =   4665
      Begin Threed.SSCheck chkDontUseRemailer 
         Height          =   465
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   820
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   255
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Don't use remailer to send the 'configuration' request.  I will send it direct to the nym server from my account."
      End
      Begin Threed.SSCheck chkManual 
         Height          =   585
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   1032
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   255
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "I know what I am doing.  So let me manually encrypt the configuration request to the Nym Server"
         Value           =   1
      End
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Finish"
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
      Left            =   2850
      TabIndex        =   3
      Top             =   4110
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
      Left            =   3930
      TabIndex        =   0
      Top             =   4110
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"C2ORG6.frx":0000
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
      Height          =   915
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4785
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"C2ORG6.frx":0101
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
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4725
   End
End
Attribute VB_Name = "frmCreateNymStep6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDontUseRemailer_Click(Value As Integer)
If Value = True Then
    Nym.DontUseRemailer = True
Else
    Nym.DontUseRemailer = False
End If
End Sub

Private Sub chkManual_Click(Value As Integer)
If Value = True Then
    gManualEncrypt = True
Else
    gManualEncrypt = False
End If
End Sub

Private Sub Command2_Click()
      gNymState = gNYM_IDLE
      Unload Me
End Sub

Private Sub Command3_Click()
    
    'If nymstate is create then listindex = 0
    Me.Hide
    Nym.NymState = gNYMCONFIG
    gRemailerSelectCaption = "You can now chose single or multiple Remailers to be used in your Reply Block.  These are the Remailers that will carry your reply back to you.  The last one in the chain will be the one that delivers to your email address."
    PIForm(gActivePIInstance).ProcessNymCommand Nym.NymState, 0
    If gCancelAction Then
        gCancelAction = False
        PIForm(gActivePIInstance).txtTo.Text = ""
    Else
        gRemailerSelectCaption = "This remailer is optional and is used so that the Nym server will not know from whom the request came.  You do not need to select a remailer here if you want to send the message directly to the Nym Server."
        PIForm(gActivePIInstance).SendNymMessage
        If Not gCancelAction Then Unload frmPI
    End If
    Unload Me
    gCancelAction = False
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
    Win.Center Me, Null
    'Win.OnTop(Me) = True
    chkManual.Value = ssCBUnchecked ' this will force an update...
   'Nym.DontUseRemailer = False
   
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateNymStep6 = Nothing
End Sub

Private Sub SSCheck1_Click(Value As Integer)

End Sub
