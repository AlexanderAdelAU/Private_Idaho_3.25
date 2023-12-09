VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateNymStep2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creating a Nym - Step 2"
   ClientHeight    =   5910
   ClientLeft      =   870
   ClientTop       =   855
   ClientWidth     =   7335
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
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5910
   ScaleWidth      =   7335
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6810
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nym Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   270
      TabIndex        =   8
      Top             =   960
      Width           =   6585
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2610
         MaxLength       =   25
         TabIndex        =   1
         Top             =   450
         Width           =   3675
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2610
         TabIndex        =   2
         Top             =   810
         Width           =   3645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Your Nym is:"
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
         Height          =   285
         Left            =   1500
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nym e.g.: Casper"
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
         Index           =   0
         Left            =   1230
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alias e.g. Casper the Ghost"
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
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label lblNym 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label6"
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
         Left            =   2610
         TabIndex        =   9
         Top             =   1200
         Width           =   3615
      End
   End
   Begin VB.Frame fNymKeys 
      Caption         =   "Create Nym Keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   270
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton btnBrowse 
         Caption         =   "Browse"
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
         Left            =   5430
         TabIndex        =   15
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtFilePath 
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
         Left            =   120
         TabIndex        =   7
         Text            =   "C:\Program Files\Network Associates\"
         Top             =   1290
         Width           =   5175
      End
      Begin VB.CommandButton bKeys 
         Caption         =   "Create Keys for Nym"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2250
         TabIndex        =   6
         Top             =   1710
         Width           =   1845
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   $"C2ORG2.frx":0000
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
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   6375
      End
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
      Left            =   5730
      TabIndex        =   3
      Top             =   5340
      Width           =   1155
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
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   5340
      Width           =   1155
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Height          =   405
      Left            =   240
      TabIndex        =   14
      Top             =   5340
      Width           =   3735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"C2ORG2.frx":0150
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
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   90
      Width           =   6375
   End
End
Attribute VB_Name = "frmCreateNymStep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pNymServer As String

Private Sub bKeys_Click()
Dim FileFunctions As New cFileFunctions
Dim objFiles As New Collection
Dim objFile As File
Dim iResponse As Long
Dim SectionName As String
Dim sPath As String
Dim pid As Long

SectionName = "PGP Options"
On Error GoTo bKeysError
Command1.Enabled = False
bKeys.Enabled = False
sPath = ReadProfile(SectionName, "PGPKeys.exe Location")
If InStr(1, sPath, ".exe") = 0 And InStr(1, txtFilePath, ".exe") = 0 Then
    'Search for the file
    Me.MousePointer = vbHourglass
    lblStatus = "Looking for PGPKeys.exe ...please be patient."
    DoEvents
    Set objFiles = FileFunctions.FindFile("PGPKeys.exe", "C:\Program Files\Network Associates")
    Me.MousePointer = vbDefault
    DoEvents
    'Add the results to the listbox
    For Each objFile In objFiles
        If UCase(objFile.name) = UCase("PGPKeys.exe") Then
            sPath = objFile.Path
            'iResponse = ShellExecute(Me.hWnd, "open", sPath, vbNullString, CurDir, SW_SHOW)
            pid = Shell(sPath, vbNormalFocus)
            If Not pid = 0 Then
                Err.Raise 6006, "Error executing PGPKeys.exe"
            Else
                lblStatus = "Writing location to registry..."
                WriteProfile SectionName, "PGPKeys.exe Location", objFile.Path
            End If
            Exit For
        End If
    Next
Else
     txtFilePath.Enabled = False
     'If InStr(1, sPath, ".exe") = 0 And InStr(1, txtFilePath, ".exe") = 0 Then
      sPath = txtFilePath
     Me.MousePointer = vbHourglass
     'iResponse = ShellExecute(Me.hWnd, "open", sPath, vbNullString, CurDir, SW_SHOW)
     
     pid = Shell(sPath, vbNormalFocus)
     If pid = 0 Then
                Err.Raise 6006, "Error executing PGPKeys.exe", iResponse
     End If
     DoEvents
     lblStatus = "When you have closed PGPKeys, click on the Next button."
     WriteProfile SectionName, "PGPKeys.exe Location", sPath
     Command1.Enabled = True
     Me.MousePointer = vbDefault
End If
Set objFile = Nothing
Set objFiles = Nothing
Set FileFunctions = Nothing
bKeys.Enabled = True
Exit Sub
bKeysError:
    Me.MousePointer = vbDefault
    txtFilePath.Enabled = True
    bKeys.Enabled = True
    WriteProfile SectionName, "PGPKeys.exe Location", ""
    MsgBox "Error detected while trying to find or execute the PGPKeys.exe file.  Error reported as: " & Err.Description, vbApplicationModal + vbCritical, "File Location Error"
    Err.Clear
End Sub

Private Sub btnBrowse_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Find PGPKeys.exe"
CommonDialog1.Flags = &H2& + &H4&
CommonDialog1.Filter = "PGP Files (PGP*.exe)|PGP*.exe"
'"Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.CancelError = True
CommonDialog1.InitDir = txtFilePath
CommonDialog1.Action = 1
txtFilePath = CommonDialog1.FileName
ChDir App.Path
End Sub

Private Sub Command1_Click()
   
    Dim tmpstr As String
    Dim i As Integer
    Dim msg As String
   
   If Not KeyOnKeyRing(lblNym) Then
        msg = "You do not have the Public Key associated with this Nym on your keyring.  If you wish to proceed with this Nym then launch the application 'PGPKeys' that comes with PGP and then create a key that matches: " & vbCrLf & vbCrLf
        msg = msg & Text1.Text & " " & lblNym & "."
        msg = msg & vbCrLf & vbCrLf & "Once this is complete proceed by pressing the 'Next' button."
        MsgBox msg, vbApplicationModal + vbCritical, "Nym PublicKey Required"
        fNymKeys.Visible = True
        tmpstr = ReadProfile("PGP Options", "PGPKeys.exe Location")
        If Not tmpstr = "" Then
            txtFilePath.Enabled = False
            txtFilePath = tmpstr
        End If
        Exit Sub
    End If
   
   If Text1.Text <> "" And Text2.Text <> "" Then
        Nym.ID = lblNym
        Nym.name = Text1.Text
        gIsNewNym = True
        gNymState = gNYMCONFIG
        frmCreateNymStep3.Show
        Unload Me
    Else
        MsgBox "You need to provide a nym and full nym to continue.", vbApplicationModal, "nym"
        gIsNewNym = False
    End If
End Sub

Private Sub Command2_Click()
    gIsNewNym = False
    Unload Me
End Sub



Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Form_Load()
    Dim Win As New CWindow
    Dim i As Integer
    Win.Center Me, Null
    'Win.OnTop(Me) = True
    'If gPGPVersion = PGP5x Then
       ' bkyes.Enabled = False
       ' Command3.Caption = "Not availble for PGP5x"
    'End If
    i = InStr(1, Nym.Server, "@")
    pNymServer = Mid$(Nym.Server, i, (Len(Nym.Server) - i) + 1)
    lblNym = "......" & pNymServer
    lblStatus = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateNymStep2 = Nothing
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
lblNym = Mid(Text2.Text, 1, Len(Text2.Text)) & pNymServer 'IIf(Shift = 0, LCase(Chr(KeyCode)), Chr(KeyCode)) & lblfullnym
End Sub
