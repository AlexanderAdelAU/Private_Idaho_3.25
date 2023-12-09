VERSION 5.00
Begin VB.Form frmRemailerOtions 
   Caption         =   "Remailer Options"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Mixmaster Info"
      Height          =   2475
      Left            =   270
      TabIndex        =   9
      Top             =   2340
      Width           =   7155
      Begin VB.CommandButton btnAdd 
         Caption         =   "...Add"
         Height          =   315
         Index           =   4
         Left            =   6210
         TabIndex        =   15
         Top             =   1830
         Width           =   735
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "...Add"
         Height          =   315
         Index           =   3
         Left            =   6210
         TabIndex        =   14
         Top             =   1170
         Width           =   735
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "...Add"
         Height          =   315
         Index           =   2
         Left            =   6210
         TabIndex        =   13
         Top             =   540
         Width           =   735
      End
      Begin VB.ComboBox cmdMixPubRingURL 
         Height          =   315
         Left            =   420
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1830
         Width           =   5625
      End
      Begin VB.ComboBox cmdMixType2URL 
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1200
         Width           =   5625
      End
      Begin VB.ComboBox cmdMixListURL 
         Height          =   315
         Left            =   420
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   570
         Width           =   5625
      End
      Begin VB.Label Label1 
         Caption         =   "Get Mixmaster Public Rings from this URL"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1590
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Get Mixmaster Type 2 Info from this URL:"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   990
         Width           =   3645
      End
      Begin VB.Label Label1 
         Caption         =   "Get Mixmaster List from this URL:"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   330
         Width           =   2835
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remailer Info"
      Height          =   2085
      Left            =   270
      TabIndex        =   2
      Top             =   60
      Width           =   7125
      Begin VB.ComboBox cmbRemailerURL 
         Height          =   315
         Left            =   420
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   630
         Width           =   5565
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "...Add"
         Height          =   315
         Index           =   0
         Left            =   6180
         TabIndex        =   5
         Top             =   660
         Width           =   735
      End
      Begin VB.ComboBox cmbKeysURL 
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1380
         Width           =   5625
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "...Add"
         Height          =   315
         Index           =   1
         Left            =   6180
         TabIndex        =   3
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Get Remailer Info from this URL:"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Get Remailer Keys from this URL:"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1110
         Width           =   2835
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6510
      TabIndex        =   1
      Top             =   4920
      Width           =   915
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5550
      TabIndex        =   0
      Top             =   4950
      Width           =   915
   End
End
Attribute VB_Name = "frmRemailerOtions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Loading As Boolean

Private Sub btnAdd_Click(Index As Integer)
Select Case Index
Case 0
    gRemailerTypeURL = 0
    frmRemailerURL.lblListType = "Standard Remailer"
Case 1
    frmRemailerKeysURL.Show
    Exit Sub
Case 2
    gRemailerTypeURL = 1
    frmRemailerURL.lblListType = "URLs for Mixmaster List"
Case 3
    gRemailerTypeURL = 2
    frmRemailerURL.lblListType = "URLs for Type 2 Mixmaster"
Case 4
    gRemailerTypeURL = 3
    frmRemailerURL.lblListType = "URLs for Mixmaster Public Rings"
End Select
frmRemailerURL.Show
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOkay_Click()
 Dim SectionName As String
 SectionName = "Net Info"
 
'First Remailer
gRemailerInfoURL = cmbRemailerURL.Text
WriteProfile SectionName, "RemailerInfoURL", gRemailerInfoURL
gPGPKeysURL = cmbKeysURL
WriteProfile SectionName, "PGPKeysURL", gPGPKeysURL

'Now Mixmaster
gMixListURL = cmdMixListURL.Text 'ReadProfile(SectionName, "MixListURL")
WriteProfile SectionName, "MixListURL", gMixListURL
gMixType2URL = cmdMixType2URL.Text 'ReadProfile(SectionName, "MixType2URL")
WriteProfile SectionName, "MixType2URL", gMixType2URL
gMixPubRingURL = cmdMixPubRingURL.Text 'ReadProfile(SectionName, "MixPubRingURL")
WriteProfile SectionName, "MixPubRingURL", gMixPubRingURL

Unload Me
End Sub

Private Sub Form_Activate()
If Loading Then Exit Sub
FillInfoURLCombo
FillKeysURLCombo
'
FillMixListURLCombo
FillMixType2URLCombo
FillMixPubRingURLCombo

End Sub

Private Sub Form_Load()

Loading = True

FillInfoURLCombo
FillKeysURLCombo
'
FillMixListURLCombo
FillMixType2URLCombo
FillMixPubRingURLCombo

Loading = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmRemailerOptions = Nothing
End Sub

Private Sub FillInfoURLCombo()
Dim SectionName As String
Dim FileNum As Integer
Dim Item As String
cmbRemailerURL.Clear
FileNum = FreeFile
''

 SectionName = "Net Info"
 
If iFileExists(App.Path + "\InfoURL.TXT") Then
    Open App.Path + "\InfoURL.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmbRemailerURL.AddItem Item
        Loop
    Else
        gRemailerInfoURL = "http://www.publius.net/rlist"
        cmbRemailerURL.AddItem gRemailerInfoURL
    End If
    'add the preffered one first on the list
    gRemailerInfoURL = ReadProfile(SectionName, "RemailerInfoURL")
    cmbRemailerURL.Text = gRemailerInfoURL
    Close FileNum
Else
    gRemailerInfoURL = "http://www.publius.net/rlist"
    cmbRemailerURL.AddItem gRemailerInfoURL
End If
cmbRemailerURL.ListIndex = 0

End Sub

Private Sub FillKeysURLCombo()
Dim FileNum As Integer
Dim Item As String

cmbKeysURL.Clear
FileNum = FreeFile
If iFileExists(App.Path + "\KeysURL.TXT") Then
    Open App.Path + "\KeysURL.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmbKeysURL.AddItem Item
        Loop
    Else
        gPGPKeysURL = "http://www.publius.net/pgpkeys"
        cmbKeysURL.AddItem gPGPKeysURL
        cmbKeysURL.ListIndex = 0
    End If
    Close FileNum
    'add the preffered one first on the list
    gPGPKeysURL = ReadProfile("Net Info", "PGPKeysURL")
    cmbKeysURL.Text = gPGPKeysURL
Else
    gPGPKeysURL = "http://www.publius.net/pgpkeys"
    cmbKeysURL.AddItem gPGPKeysURL
End If
cmbKeysURL.ListIndex = 0

End Sub

Public Sub FillMixListURLCombo()
Dim SectionName As String
Dim FileNum As Integer
Dim Item As String
cmdMixListURL.Clear
FileNum = FreeFile
If iFileExists(App.Path + "\MixListURL.TXT") Then
    Open App.Path + "\MixListURL.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmdMixListURL.AddItem Item
        Loop
    Else
        gMixListURL = "http://anon.efga.org/anon/mlist.html"
        cmdMixListURL.AddItem gMixListURL
    End If
    Close FileNum
    'add the preffered one first on the list
    gMixListURL = ReadProfile("Net Info", "MixListURL")
    cmdMixListURL.Text = gMixListURL
Else
    gMixListURL = "http://anon.efga.org/anon/mlist.html"
    cmdMixListURL.AddItem gMixListURL
End If
cmdMixListURL.ListIndex = 0
End Sub

Public Sub FillMixType2URLCombo()
Dim SectionName As String
Dim FileNum As Integer
Dim Item As String
cmdMixType2URL.Clear
FileNum = FreeFile
If iFileExists(App.Path + "\MixType2URL.TXT") Then
    Open App.Path + "\MixType2URL.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmdMixType2URL.AddItem Item
        Loop
    Else
        gMixType2URL = "http://anon.efga.org/anon/type2.list"
        cmdMixType2URL.AddItem gMixType2URL
    End If
    Close FileNum
    'add the preffered one first on the list
    gMixType2URL = ReadProfile("Net Info", "MixType2URL")
    cmdMixType2URL.Text = gMixType2URL
Else
   gMixType2URL = "http://anon.efga.org/anon/type2.list"
    cmdMixType2URL.AddItem gMixType2URL
End If
cmdMixType2URL.ListIndex = 0
End Sub

Public Sub FillMixPubRingURLCombo()
Dim SectionName As String
Dim FileNum As Integer
Dim Item As String
cmdMixPubRingURL.Clear
FileNum = FreeFile
If iFileExists(App.Path + "\MixPubRingURL.TXT") Then
    Open App.Path + "\MixPubRingURL.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmdMixPubRingURL.AddItem Item
        Loop
    Else
        gMixPubRingURL = "http://anon.efga.org/anon/pubring.mix"
        cmdMixPubRingURL.AddItem gMixPubRingURL
    End If
    Close FileNum
    'add the preffered one first on the list
    gMixPubRingURL = ReadProfile("Net Info", "MixPubRingURL")
    cmdMixPubRingURL.Text = gMixPubRingURL
Else
   gMixPubRingURL = "http://anon.efga.org/anon/pubring.mix"
    cmdMixPubRingURL.AddItem gMixPubRingURL
End If
cmdMixPubRingURL.ListIndex = 0
End Sub

