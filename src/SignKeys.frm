VERSION 5.00
Begin VB.Form frmSignKeys 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select User Addressee"
   ClientHeight    =   4980
   ClientLeft      =   1545
   ClientTop       =   1890
   ClientWidth     =   9450
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
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4980
   ScaleWidth      =   9450
   Begin VB.CheckBox chkManual 
      Caption         =   "Manual Entry"
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
      Height          =   315
      Left            =   6810
      TabIndex        =   14
      Top             =   3180
      Width           =   1575
   End
   Begin VB.TextBox txtPassPhrase 
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
      Height          =   735
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "SignKeys.frx":0000
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtUserID 
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
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton btnUserID 
      BackColor       =   &H00C0C0C0&
      Caption         =   "--->"
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
      Index           =   1
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "Click to select"
      Top             =   2640
      Width           =   675
   End
   Begin VB.TextBox txtUserID 
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
      Index           =   0
      Left            =   5640
      TabIndex        =   8
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton btnUserID 
      Caption         =   "--->"
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
      Index           =   0
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Click to select"
      Top             =   1560
      Width           =   675
   End
   Begin VB.ComboBox cmdAddressBook 
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
      TabIndex        =   4
      Text            =   "cmbAddressBook"
      ToolTipText     =   "Select Address Book"
      Top             =   1140
      Width           =   2775
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
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
      Left            =   8130
      TabIndex        =   1
      Top             =   4500
      Width           =   975
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
      Left            =   7020
      TabIndex        =   0
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Other ID to use to certify the above ID. (Normally this is your ID)"
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
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   13
      Top             =   2160
      Width           =   2955
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pass Phrase"
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
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   12
      Top             =   3210
      Width           =   1125
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Keyring"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2475
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "User ID to be certified or signed."
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
      Index           =   1
      Left            =   5640
      TabIndex        =   6
      Top             =   1230
      Width           =   2475
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "This funtion will certify and sign a user ID from you keyrings.  Select a key to certify or sign"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   6435
   End
End
Attribute VB_Name = "frmSignKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fPubKeySelected As Boolean

Private Sub btnUserID_Click(Index As Integer)
    txtUserID(Index) = List1.List(List1.ListIndex)
End Sub

Private Sub cmdAddressBook_Click()
If cmdAddressBook.ListIndex = 0 Then
    LoadPubKeysOut
    fPubKeySelected = True
Else
    LoadPrivateAddressBook
    fPubKeySelected = False
End If
txtUserID(0) = ""
txtUserID(1) = ""
txtPassPhrase = ""
End Sub

Private Sub Command1_Click()

    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim UserID1 As String
    Dim UserID2 As String
    Dim cmd As String
    Dim ReferenceDateTime As String
    
    pos1 = 0
    pos2 = 0
    pos1 = InStr(txtUserID(0), "<") + 1
    pos2 = InStr(txtUserID(0), ">")
    If pos2 <= pos1 Then
        UserID1 = txtUserID(0)
    Else
        UserID1 = Chr$(34) & Mid(txtUserID(0), pos1, pos2 - pos1) + Chr$(34)
    End If
    
    pos1 = 0
    pos2 = 0
    pos1 = InStr(txtUserID(1), "<") + 1
    pos2 = InStr(txtUserID(1), ">")
    If pos2 <= pos1 Then
        UserID2 = txtUserID(1)
    Else
        UserID2 = Chr$(34) & Mid(txtUserID(1), pos1, pos2 - pos1) + Chr$(34)
    End If
   cmd = gPGPPath & "\PGP -KS " & UserID1 & " -U " & UserID2 '& " -z " & Chr(34) & txtPassPhrase & Chr(34)
   Me.Hide
   DoEvents
   gPGPResponse.res(0) = "Y"
   If Not chkManual = vbChecked And Len(txtPassPhrase) > 0 Then
        gPGPResponse.res(1) = txtPassPhrase
        gPGPResponse.Count = 2
    Else
        gPGPResponse.Count = 1
    End If
   If iFileExists(gPGPPath & "\PUBRING.PGP") Then
        ReferenceDateTime = FileDateTime(gPGPPath & "\PUBRING.PGP")
    End If
   ExecCmd (cmd)
   gPGPResponse.Count = 0
   Me.Show
    If Not iFileChanged(gPGPPath & "\PUBRING.PGP", ReferenceDateTime) Then
        MsgBox "PUBRING.PGP was not updated.  It is likely that the User ID you selected has already been certified.", vbApplicationModal + vbExclamation, App.Title
    Else
        UpdatePublicKeysFile
        frmPI.ShowStatus ("Key signature and verification was successful.")
    End If
   Unload Me
End Sub

Private Sub Command2_Click()
     Unload Me
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
Dim FileNum As Integer
Dim TextLine As String
Dim tmpstr As String
Dim Pos As Integer
    
Win.Center Me, Null
Win.OnTop(Me) = True
    
    fPubKeySelected = True
    'Default to Pubkeys.out
    cmdAddressBook.AddItem "Public Keys"
    LoadPubKeysOut
    'cmdAddressBook.AddItem "Private Address Book"
    'LoadPrivateAddressBook
    cmdAddressBook.ListIndex = 0
    txtUserID(0) = ""
     txtUserID(1) = ""
     txtPassPhrase = ""
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSelectUserID = Nothing
End Sub

Private Sub lblUserID_Click()

End Sub

Private Sub List1_DblClick()
   Exit Sub
   'fix this later
    Dim pos1 As Integer
    Dim pos2 As Integer
    'Clear existing Key ID
    gKeyID = ""
    pos1 = 0
    pos2 = 0
    
    pos1 = InStr(List1.Text, "<") + 1
    pos2 = InStr(List1.Text, ">")
    If Len(gKeyID) > 0 Then
        gKeyID = gKeyID + " " + Chr$(34) + Mid(List1.Text, pos1, pos2 - pos1) + Chr$(34)
    Else
        gKeyID = Chr$(34) + Mid(List1.Text, pos1, pos2 - pos1) + Chr$(34)
    End If
    Unload Me
End Sub

Private Sub LoadPubKeysOut()
Dim FileNum As Integer
Dim TextLine As String
Dim tmpstr As String
Dim Pos As Integer

'---------------------------------------------
    'Load the list from PUBKEYS.OUT
    '---------------------------------------------
    FileNum = FreeFile
    List1.Clear
   ' Label2.Caption = Label2.Caption & gKeyID & "."
    Open App.Path + "\pubkeys.out" For Input As FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, TextLine
        tmpstr = TextLine
        If Mid$(TextLine, 1, 3) = "pub" Then
            Pos = InStr(TextLine, " ")
            TextLine = Mid$(TextLine, Pos + 1, Len(TextLine))
            Pos = InStr(TextLine, " ")
            TextLine = Mid$(TextLine, Pos + 1, Len(TextLine))
            Pos = InStr(TextLine, " ")
            TextLine = Mid$(TextLine, Pos + 1, Len(TextLine))
            Pos = InStr(TextLine, " ")
            TextLine = Mid$(TextLine, Pos + 1, Len(TextLine))
            If Mid$(tmpstr, 6, 1) = " " Then
                Pos = InStr(TextLine, " ")
                TextLine = Mid$(TextLine, Pos + 1, Len(TextLine))
            End If
            List1.AddItem TextLine
        End If
    Wend
    Close #FileNum
End Sub

Private Sub LoadPrivateAddressBook()
'get the address list
Dim Item As String
Dim FileNum As Integer
    FileNum = FreeFile
    List1.Clear
    If iFileExists(App.Path + "\ADDRESS.TXT") Then
        Open App.Path + "\ADDRESS.TXT" For Input As FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            List1.AddItem Item
        Loop
        Close FileNum
    End If
End Sub

Private Sub List2_Click()

End Sub

Private Sub txtPassPhrase_Change()
chkManual.Value = vbUnchecked
If Len(txtPassPhrase) = 0 Then chkManual.Value = vbChecked
End Sub
