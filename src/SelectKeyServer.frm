VERSION 5.00
Begin VB.Form frmSelectKeyServer 
   Caption         =   "Select a Key Server"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAdd 
      Caption         =   "...Add"
      Height          =   315
      Index           =   1
      Left            =   6840
      TabIndex        =   7
      Top             =   1620
      Width           =   735
   End
   Begin VB.ComboBox cmbKeyServerAddress 
      Height          =   315
      Left            =   540
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1620
      Width           =   6135
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "...Add"
      Height          =   315
      Index           =   0
      Left            =   6840
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
   Begin VB.ComboBox cmbKeyServerURL 
      Height          =   315
      Left            =   540
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   540
      Width           =   6135
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   2280
      Width           =   915
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Submit Key to this server Address:"
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1260
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Get Key from this server URL:"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4095
   End
End
Attribute VB_Name = "frmSelectKeyServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Loading As Boolean

Private Sub btnAdd_Click(index As Integer)
    frmKeyServer.Show
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOkay_Click()
Dim SectionName As String
SectionName = "Net Info"
 
gGetKeyURL = cmbKeyServerURL.Text
WriteProfile SectionName, "GetKeyURL", gGetKeyURL

gSubKeyURL = cmbKeyServerAddress.Text
WriteProfile SectionName, "SubmitKeyURL", gSubKeyURL

Unload Me
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
If Loading Then Exit Sub
FillServerURLCombo
FillServerAddressCombo
End Sub

Private Sub Form_Load()

Loading = True

FillServerURLCombo
FillServerAddressCombo

Loading = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSelectKeyServer = Nothing
End Sub

Private Sub FillServerURLCombo()
Dim SectionName As String
Dim FileNum As Integer
Dim Item As String
cmbKeyServerURL.Clear
FileNum = FreeFile
If iFileExists(App.Path + "\KeyServerURL.TXT") Then
    Open App.Path + "\KeyServerURL.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmbKeyServerURL.AddItem Item
        Loop
    Else
        gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
        cmbKeyServerURL.AddItem gGetKeyURL
    End If
    Close FileNum
Else
    gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
    cmbKeyServerURL.AddItem gGetKeyURL
End If
cmbKeyServerURL.ListIndex = 0

End Sub
Private Sub FillServerAddressCombo()
Dim SectionName As String
Dim FileNum As Integer
Dim Item As String
cmbKeyServerAddress.Clear
FileNum = FreeFile
If iFileExists(App.Path + "\KeyServerAddr.TXT") Then
    Open App.Path + "\KeyServerAddr.TXT" For Input As FileNum
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            cmbKeyServerAddress.AddItem Item
        Loop
    Else
        gSubKeyURL = "pgp-public-keys@pgp.ai.mit.edu"
        cmbKeyServerAddress.AddItem gSubKeyURL
    End If
    Close FileNum
Else
    gSubKeyURL = "pgp-public-keys@pgp.ai.mit.edu"
    cmbKeyServerAddress.AddItem gSubKeyURL
End If
cmbKeyServerAddress.ListIndex = 0
End Sub
