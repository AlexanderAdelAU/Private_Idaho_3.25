VERSION 5.00
Begin VB.Form frmPGPOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PGP Options"
   ClientHeight    =   6165
   ClientLeft      =   1890
   ClientTop       =   525
   ClientWidth     =   7560
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
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   7560
   Begin VB.CommandButton bclear 
      Caption         =   "Clear"
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
      Left            =   6420
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select Default PGP User Key"
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
      Left            =   180
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame Opltions 
      Caption         =   "Options"
      Height          =   3465
      Left            =   150
      TabIndex        =   4
      Top             =   2040
      Width           =   6945
      Begin VB.CheckBox chkEncryptSentMessagesToSelf 
         Caption         =   "Encrypt Sent Messages to Self"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   3060
         Width           =   3375
      End
      Begin VB.CheckBox chkLoadPassPhrase 
         Caption         =   "Load passphrase into memory at start-up"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   6
         Top             =   960
         Width           =   4455
      End
      Begin VB.CheckBox chkAutoDecryptKeyRings 
         Caption         =   "Auto encrypt/decrypt key rings"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   2010
         Width           =   4455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"EUPGPOP.frx":0000
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
         Height          =   735
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   2340
         Width           =   6525
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"EUPGPOP.frx":0101
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
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   6405
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"EUPGPOP.frx":01B5
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
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1380
         Width           =   6405
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3990
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5040
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtDefaultKey 
      Enabled         =   0   'False
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
      Left            =   2580
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"EUPGPOP.frx":026F
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
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7065
   End
End
Attribute VB_Name = "frmPGPOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempPath As String
Dim TempKeyID As String

Private Sub bclear_Click()
txtDefaultKey = ""
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
   ' Dim tmpstr As String
    Dim SectionName As String
    SectionName = "PGP Options"
    
     WriteProfile SectionName, "Default Key ID", txtDefaultKey
  
    If chkLoadPassPhrase.Value = vbChecked Then
        WriteProfile SectionName, "Passphrase", "true"
    Else
        WriteProfile SectionName, "Passphrase", "false"
    End If
    If chkAutoDecryptKeyRings.Value = vbChecked Then
        WriteProfile SectionName, "AutoEncrypt", "true"
    Else
        WriteProfile SectionName, "AutoEncrypt", "false"
    End If
    If chkEncryptSentMessagesToSelf.Value = vbChecked Then
        WriteProfile SectionName, "EncryptToSelf", "true"
    Else
        WriteProfile SectionName, "EncryptToSelf", "false"
    End If
    Unload Me
End Sub



Private Sub Command4_Click()
'Only show private keys as this is used for signing
MousePointer = vbHourglass
DoEvents
vb2spgpContext.SelectPrivateKeys = True
PIForm(gActivePIInstance).mKeyRingIDs_Click
txtDefaultKey = Key.UserID
MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Dim SectionName As String
    Dim i As Integer
    Dim tmpstr As String
           
    SectionName = "PGP Options"
    
    gPGPKeyID = ReadProfile(SectionName, "Default Key ID")
    If gPGPKeyID = "" Then
        txtDefaultKey = ""
    Else
        txtDefaultKey = gPGPKeyID
    End If
    
    tmpstr = ReadProfile(SectionName, "Passphrase")
    If tmpstr = "true" Then
        chkLoadPassPhrase.Value = 1
    Else
        chkLoadPassPhrase.Value = 0
    End If
    
    'tmpstr = ReadProfile(SectionName, "AutoEncrypt")
   ' If tmpstr = "true" Then
     '   chkAutoDecryptKeyRings.Value = 1
   ' Else
    '    chkAutoDecryptKeyRings.Value = 0
   ' End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPGPOptions = Nothing
End Sub

