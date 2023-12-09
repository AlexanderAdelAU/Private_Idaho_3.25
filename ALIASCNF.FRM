VERSION 5.00
Begin VB.Form frmAliasConfiguration 
   HelpContextID   =  0
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alias.net Configuration"
   ClientHeight    =   3855
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5520
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
   LinkTopic       =   "Form35"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   5520
   Begin VB.CommandButton Command1 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Command1                                                                                                                         *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   4200
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Command2                                                                                                                         *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Text1                                                                                                                            *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CheckBox Check6 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Check6                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Allow fingering"
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
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CheckBox Check5 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Check5                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Disable nym account"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CheckBox Check4 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Check4                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Fixed size messages"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CheckBox Check3 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Check3                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Automatic public key encryption"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Check2                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Automatic PGP signing"
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
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Check1                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Automatic acknowledgement"
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
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Label2                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Send e-mail to help@nym.alias.net for information about these options."
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      WhatsThisHelpID =  0  'WhatsThisHelpID for: Label1                                                                                                                           *** HelpWriter *** 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Full name: "
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
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "frmAliasConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim temp As String
Dim tmpstr As String

    'based on the value of the check boxes, constructs
    'a string to pass to the gNym server
    If Check1.Value = 1 Then
        temp = "+acksend "
    Else
        temp = "-acksend "
    End If
    If Check2.Value = 1 Then
        temp = temp + "+signsend "
    Else
        temp = temp + "-signsend "
    End If
    If Check3.Value = 1 Then
        temp = temp + "+cryptrecv "
    Else
        temp = temp + "-cryptrecv "
    End If
    If Check4.Value = 1 Then
        temp = temp + "+fixedsize "
    Else
        temp = temp + "-fixedsize "
    End If
    If Check5.Value = 1 Then
        temp = temp + "+disable "
    Else
        temp = temp + "-disable "
    End If
    If Check6.Value = 1 Then
        temp = temp + "+fingerkey "
    Else
        temp = temp + "-fingerkey "
    End If
    'stuff the gloabl variable with the string settings
    gNymConfigStr = temp & "name=" & Chr$(34) & Text1.Text & Chr$(34)   '"Asc("""")
    'call the configuration routine
    frmMultiNyms.NymAliasConfig
    'set the to: box to the gNym server
    frmPI.txtTo.Text = gNymServer
    tmpstr = InStr(1, gNymServer, "@")
    'fie$ = Mid$(nymServer, tmpStr, (Len(nymServer) - tmpStr) + 1)
    temp = gPGPKeyID
    gPGPKeyID = gNym + Mid$(gNymServer, tmpstr, (Len(gNymServer) - tmpstr) + 1)
    gIsNewNym = False
    Me.Hide
    'encrypt the configuration message
    EncryptMessage gNymServer, "-eatws"
    gPGPKeyID = temp
    Unload Me
End Sub

Private Sub Form_Load()
'*** Code added by HelpWriter ***
    SetApphelp Me.hWnd
'***********************************
    Text1.Text = gfullNym
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*** Code added by HelpWriter ***
    QuitHelp
'***********************************
Set frmAliasConfiguration = Nothing
End Sub
