VERSION 5.00
Begin VB.Form frmUSENETGateways 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USENET Gateways"
   ClientHeight    =   6645
   ClientLeft      =   2400
   ClientTop       =   1245
   ClientWidth     =   8835
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
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6645
   ScaleWidth      =   8835
   Begin VB.Frame Frame1 
      Caption         =   "Use Remailer?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   5790
      TabIndex        =   18
      Top             =   4590
      Width           =   2565
      Begin VB.OptionButton optRemailer 
         Caption         =   "No"
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
         Left            =   1380
         TabIndex        =   20
         Top             =   660
         Width           =   705
      End
      Begin VB.OptionButton optRemailer 
         Caption         =   "Yes"
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
         Left            =   300
         TabIndex        =   19
         Top             =   660
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Select if you wish to send this request through a remailer."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   21
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.TextBox txtOtherHeadersID 
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
      TabIndex        =   17
      Top             =   2670
      Width           =   1335
   End
   Begin VB.TextBox txtOtherHeaders 
      Height          =   285
      Left            =   1590
      TabIndex        =   16
      ToolTipText     =   "eg X-Header: This is a header"
      Top             =   2670
      Width           =   6735
   End
   Begin VB.TextBox txtReferences 
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
      Left            =   1590
      TabIndex        =   14
      Top             =   2340
      Width           =   6735
   End
   Begin VB.TextBox txtMessageID 
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
      Left            =   1590
      TabIndex        =   11
      Top             =   2010
      Width           =   6735
   End
   Begin VB.TextBox txtSubject 
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
      Left            =   1590
      TabIndex        =   10
      Top             =   1680
      Width           =   6735
   End
   Begin VB.TextBox txtNewsGroups 
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
      Left            =   1590
      TabIndex        =   5
      Top             =   1350
      Width           =   6735
   End
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
      TabIndex        =   2
      Top             =   4050
      Width           =   5385
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ok"
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
      Height          =   345
      Left            =   5790
      TabIndex        =   1
      Top             =   5850
      Width           =   1065
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
      Height          =   345
      Left            =   7260
      TabIndex        =   0
      Top             =   5850
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"EUGATE.frx":0000
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
      Height          =   405
      Index           =   7
      Left            =   90
      TabIndex        =   22
      Top             =   0
      Width           =   8205
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Newsgroups:"
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
      Index           =   6
      Left            =   150
      TabIndex        =   15
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "References: "
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
      Index           =   5
      Left            =   150
      TabIndex        =   13
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Message-ID:"
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
      Index           =   4
      Left            =   150
      TabIndex        =   12
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"EUGATE.frx":00C7
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
      Height          =   1515
      Index           =   2
      Left            =   5730
      TabIndex        =   9
      Top             =   3090
      Width           =   2805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject: "
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
      Index           =   3
      Left            =   150
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblstatus 
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   7
      Top             =   5850
      Width           =   5355
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"EUGATE.frx":01A1
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
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   480
      Width           =   8205
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"EUGATE.frx":034E
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
      Height          =   795
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   3120
      Width           =   5355
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "last updated: "
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
      Left            =   5760
      TabIndex        =   4
      Top             =   4350
      Width           =   2355
   End
End
Attribute VB_Name = "frmUSENETGateways"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If List1.Text = "" Then
        Beep
        lblStatus = "Select a gateway from the list!"
        DoEvents
        Exit Sub
    End If
    PIForm(gActivePIInstance).txtTo.Text = List1.Text 'PIForm(gActivePIInstance).txtTo.Text & List1.Text
    MailHeader(1).Value = txtNewsGroups
    MailHeader(2).Value = txtsubject
    If txtMessageID = "" Then
        Randomize
        MailHeader(3).Value = Int(1000000000 * Rnd) & "@nowhere.net"
    Else
        MailHeader(3).Value = txtMessageID
    End If
    MailHeader(4).Value = txtReferences
    MailHeader(5).Value = txtOtherHeaders
    MailHeader(5).ID = txtOtherHeadersID
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
    PIForm(gActivePIInstance).USENETGate.Checked = False
    gNewsgroupType = 0
    gCancelAction = True
End Sub

Private Sub Form_Load()
Dim FileNum As Integer
Dim Item As String
    'Set default values
    gRemailerType = REMAILER_CYPHERPUNK
    gCancelAction = False
    MailHeader(0).Value = ""
    MailHeader(0).ID = ""
    If MailHeader(1).Value = "" Then
        txtNewsGroups = "alt.test"
    Else
        txtNewsGroups = MailHeader(1).Value
    End If
   ' txtsubject = MailHeader(2).Value
    
    If MailHeader(2).Value = "" Then
        txtsubject = IIf(PIForm(gActivePIInstance).txtsubject = "", "No Subject", PIForm(gActivePIInstance).txtsubject)
    Else
        txtsubject = MailHeader(2).Value
    End If
    
    If MailHeader(3).Value = "" Then
        Randomize
        txtMessageID = Int(Now() * Rnd * 1000000) & 10000 * Now() & "@nowhere.net"
    Else
        txtMessageID = MailHeader(3).Value
    End If
    
    txtReferences = MailHeader(4).Value
    txtOtherHeaders = MailHeader(5).Value
    txtOtherHeadersID = MailHeader(5).ID
    
    If iFileExists(App.Path + "\USENET.TXT") Then
        FileNum = FreeFile
        Open App.Path & "\USENET.TXT" For Input As FileNum
        Line Input #FileNum, Item
        Label2.Caption = Label2.Caption + Item
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            List1.AddItem Item
        Loop
        Close #FileNum
    Else
        MsgBox "USENET.TXT not present in the same directory as " & gPiStr + ".  Type in the USENET gateways you wish to use.", 48, gPiStr
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUSENETGateways = Nothing
End Sub


Private Sub optRemailer_Click(Index As Integer)
If optRemailer(0).Value = True Then
    gRemailerType = REMAILER_CYPHERPUNK 'PI.UseCypherPunk
Else
    gRemailerType = STANDARD_EMAIL 'PIForm(gActivePIInstance).DontUseRemailer
End If
End Sub
