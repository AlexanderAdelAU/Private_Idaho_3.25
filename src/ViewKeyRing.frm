VERSION 5.00
Begin VB.Form frmViewKeyRing 
   Caption         =   "Keys"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Keys"
      Height          =   3255
      Left            =   240
      TabIndex        =   18
      Top             =   630
      Width           =   7065
      Begin VB.OptionButton optExport 
         Caption         =   "Public Key"
         Height          =   225
         Index           =   1
         Left            =   4620
         TabIndex        =   22
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optExport 
         Caption         =   "Private Key"
         Height          =   225
         Index           =   0
         Left            =   3360
         TabIndex        =   21
         Top             =   2880
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   2535
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   270
         Width           =   6765
      End
      Begin VB.Label Label8 
         Caption         =   "To export a key double-click on the id."
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   2880
         Width           =   2955
      End
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "Cancel"
      Height          =   405
      Index           =   1
      Left            =   6120
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "Select Key"
      Height          =   405
      Index           =   0
      Left            =   4770
      TabIndex        =   15
      Top             =   5730
      Width           =   1215
   End
   Begin VB.TextBox edCreated 
      Height          =   285
      Left            =   1170
      TabIndex        =   9
      Top             =   4590
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Key Properties"
      Height          =   1575
      Left            =   210
      TabIndex        =   0
      Top             =   3930
      Width           =   7095
      Begin VB.TextBox edValidity 
         Height          =   285
         Left            =   6120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edTrust 
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edPrivate 
         Height          =   285
         Left            =   6120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox edBits 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edAlg 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edFingerprint 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   330
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Validity:"
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Trust:"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Created:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Private or Public?"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Key Size:"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Algorithm:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fingerprint:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label lblContext 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   330
      TabIndex        =   17
      Top             =   60
      Width           =   7095
   End
End
Attribute VB_Name = "frmViewKeyRing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnCommand_Click(Index As Integer)

If Index = 0 Then
    If List1.SelCount = 1 Then
        gPGPKeyID = Key.UserID
        gCancelAction = False
    Else
        gCancelAction = True
    End If
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim BufferOut As String
Dim i As Long
Dim Count As Long
Dim BufferLen As Long
  
On Error Resume Next
BufferLen = spgpKeyRingCount() * 1024
BufferOut = String(BufferLen, Chr(0))
i = spgpKeyRingID(BufferOut, BufferLen)

If Not i = 0 Then
    Beep
    MsgBox "Keyring not found...", vbApplicationModal + vbCritical, "KeyRing Error"
Else
    Count = CountCRLF(BufferOut)
    Call ChopKeyProps(BufferOut, Count)
    For i = 0 To UBound(KeyArray())
        If vb2spgpContext.SelectPrivateKeys Then
            If KeyArray(i).Private Then If KeyArray(i).KeyID <> vbNullString Then List1.AddItem KeyArray(i).KeyID & Chr(9) & KeyArray(i).UserID
        Else
            If KeyArray(i).KeyID <> vbNullString Then List1.AddItem KeyArray(i).KeyID & Chr(9) & KeyArray(i).UserID
        End If
    Next i
End If
End Sub

' extract the Key ID from the selected list item
' call keyprops to get the key's properties
Private Sub List1_Click()
Dim i As Long
Dim BufferIn As String
Dim KeyProperties As String

KeyProperties = String(KEYPROPS_BUFFER_SIZE, Chr(0))
'BufferIn = Space(4096)

' first 10 characters will be the key id (0x12345678)
BufferIn = Mid(List1.List(List1.ListIndex), 1, 10) & Chr(0)

' keyprops takes either key id(s) or user id(s)
' and returns the key's properties
i = spgpKeyProps(BufferIn, KeyProperties, Len(KeyProperties))

' parse the returned property-string into a TKey_Data record
Key = ParseKeyData(KeyProperties)
'Clipboard.SetText key.
  edFingerprint.Text = Key.Fingerprint
  edAlg.Text = Key.KeyAlgorithm
  edBits.Text = Key.Bits
  
  If Key.Private = True Then
    edPrivate = "Private"
  Else
    edPrivate = "Public"
  End If
  
  edCreated.Text = Key.DateTimeStr
  edTrust.Text = Key.Trust
  edValidity.Text = Key.Validity
  'Do this just in case
  gPGPKeyID = Key.UserID
  btnCommand(0).Enabled = True
End Sub

Private Sub List1_DblClick()
vb2spgpContext.SelectPrivateKeys = IIf(optExport(0).Value = True, 1, 0)
'Mid(List1.List(List1.ListIndex), 1, 10)
frmExportKey.Text1.Text = GetKey(Mid(List1.List(List1.ListIndex), 1, 10), True)
frmExportKey.Show vbModal
End Sub

