VERSION 5.00
Begin VB.Form frmStatus 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Private Idaho - Net Status"
   ClientHeight    =   2850
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   5835
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
   LinkTopic       =   "Form33"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   5835
   Begin VB.Frame fMode 
      Caption         =   "Scanning Out Box"
      Height          =   1755
      Left            =   270
      TabIndex        =   1
      Top             =   450
      Width           =   5145
      Begin VB.Label lblBytesTransferred 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "lblBytesTransferred"
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
         Left            =   1650
         TabIndex        =   4
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Transferred: "
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
         Left            =   180
         TabIndex        =   3
         Top             =   1110
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label lblTransferStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "lblTransferStatus"
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
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Width           =   4740
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   300
      Left            =   4710
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2430
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "STATUS.frx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    On Error GoTo CancelErrHandler
    'check state
    Screen.MousePointer = 0
    Select Case MailConnector.ServerState
    'disconnect from pop server
    Case POPSTATE
        frmMain.SMTP1.Action = 0
        frmMain.POP1.Action = 2
        ' frmMain.POP1.WinsockLoaded = False
        'disconnect from smtp server
    Case SMTPSTATE
        frmMain.SMTP1.Action = 0
        frmMain.SMTP1.Action = 4
        frmMain.SMTP1.Action = 2
        frmMain.SMTP1.WinsockLoaded = False
        MsgBox "Mail not sent.", vbApplicationModal, App.Title
        'disconnect from http server
    Case HTTPSTATE
        frmMain.SMTP1.Action = 0
        
    Case IMAPSTATE
        frmMain.IMAP1.Action = 0
        frmMain.IMAP1.Action = 3

    End Select
    MailConnector.ServerState = 0
    Unload Me
    Exit Sub

CancelErrHandler:
    Select Case MailConnector.ServerState
        Case SMTPSTATE
            MsgBox "Mail not sent.", vbApplicationModal, App.Title
    End Select
    MailConnector.ServerState = 0
    frmStatus.Hide
    MsgBox Err.Description & "(in Status)", vbApplicationModal, App.Title
    Err.Clear
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
    Win.Center Me, Null
    Win.OnTop(Me) = True
    lblBytesTransferred = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmStatus = Nothing
End Sub

Private Sub lblBytesTransferred_Change()
If Val(lblBytesTransferred) > 0 Then lblstatus.Visible = True
End Sub

