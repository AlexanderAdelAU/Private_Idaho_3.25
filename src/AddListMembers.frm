VERSION 5.00
Begin VB.Form frmAddListMembers 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select List Members"
   ClientHeight    =   4815
   ClientLeft      =   1545
   ClientTop       =   1890
   ClientWidth     =   8670
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
   ScaleHeight     =   4815
   ScaleWidth      =   8670
   Begin VB.CommandButton btnDone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
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
      Height          =   315
      Index           =   1
      Left            =   7380
      TabIndex        =   9
      Top             =   4230
      Width           =   975
   End
   Begin VB.ComboBox cmdAddressBook 
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
      Height          =   315
      ItemData        =   "AddListMembers.frx":0000
      Left            =   210
      List            =   "AddListMembers.frx":0002
      TabIndex        =   7
      Text            =   "cmbAddressBook"
      ToolTipText     =   "Select Address Book"
      Top             =   750
      Width           =   3345
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "<- Remove"
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
      Left            =   3810
      TabIndex        =   4
      Top             =   2070
      Width           =   975
   End
   Begin VB.CommandButton btnTo 
      Caption         =   " Add ->"
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
      Left            =   3810
      TabIndex        =   3
      ToolTipText     =   "Click to select"
      Top             =   1320
      Width           =   945
   End
   Begin VB.ListBox ToList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   4860
      TabIndex        =   2
      Top             =   750
      Width           =   3525
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
      Height          =   2595
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1290
      Width           =   3435
   End
   Begin VB.CommandButton btnDone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
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
      Index           =   0
      Left            =   6210
      TabIndex        =   0
      Top             =   4230
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  You can only create member lists that come from either your personal address book, or your public key ring, NOT BOTH."
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
      Height          =   615
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   60
      Width           =   3435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the members to add to this list: "
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
      Left            =   5130
      TabIndex        =   6
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label lblListName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblListName"
      Height          =   255
      Left            =   4980
      TabIndex        =   5
      Top             =   420
      Width           =   3285
   End
End
Attribute VB_Name = "frmAddListMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'rivate fPubKeySelected As Boolean
Private fLoading As Boolean
'Values for m_MemberTYpe - this refers to members of groups,
'ie does a group come from the Private Contacts or from the PGP Keyring etc
'0 not set in database
'1 - Members from Private Contacts
'2 - Members from PGP keyring
Private m_MemberType As Integer

Private Sub btnDone_Click(Index As Integer)
Dim rsMembers As Recordset
Dim rsList As Recordset

Dim i As Integer
Dim lMemberID As Long
'Dim iStartID As Integer
'Dim iEndID As Integer

On Error GoTo MemberError
If Index = 1 Then
    Unload Me
    Exit Sub
End If

Set rsList = DB.OpenRecordset(lblListName, dbOpenDynaset)
Set rsMembers = DB.OpenRecordset("Master List", dbOpenDynaset)
rsMembers.FindFirst "[List Name] = " & "'" & lblListName & "'"
rsMembers.Edit
rsMembers("Member Type") = m_MemberType
rsMembers.Update
rsMembers.Close


'Remove previous members and fill with new ones..
While Not rsList.EOF
    rsList.Delete
    rsList.MoveNext
Wend

For i = 0 To ToList.ListCount - 1
       rsList.AddNew
        rsList("Member ID") = 1 'this is temporary..
        rsList("Contact Email") = ToList.List(i) 'StripEMailAddress(ToList.List(i))  'rsMembers("Address Book ID")
        rsList.Update
    DoEvents
Next
rsList.Close

'Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
'rs.FindFirst "[List Name] = " & "'" & lblListName & "'"
'If rs.EOF Then Err.Raise 6003, "Add Members", "The group selected is not in the database."

'Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
'rs.FindFirst "[List Name] = " & "'" & lblListName & "'"

Unload Me
Exit Sub
MemberError:
    MsgBox "An error has occured: " & Err.Description, vbApplicationModal + vbCritical
    Err.Clear
    Unload Me
End Sub

Private Sub btnRemove_Click()
On Error Resume Next

If ToList.ListIndex = -1 Then Exit Sub
    ToList.RemoveItem (ToList.ListIndex)
'End If
If ToList.ListCount = 0 Then
    DeleteAllMembers
    m_MemberType = 0
    cmdAddressBook.Enabled = True
End If
End Sub

Private Sub btnTo_Click()

m_MemberType = cmdAddressBook.ListIndex + 1
cmdAddressBook.Enabled = False

If List1.List(List1.ListIndex) = "" Then Exit Sub
    ToList.AddItem List1.List(List1.ListIndex)
End Sub

Private Sub cmdAddressBook_Click()
Dim bResponse As Integer

List1.Clear
ToList.Clear

Select Case cmdAddressBook.ListIndex

    Case 1
        If Not PGP_SDKPresent Or gPGPVersion = NoPGP Then
            MsgBox "PI either can't find PGP SDK in Windows System directory or you have disabled PGP.  You must install PGP to be able select from the public key list", vbApplicationModal + vbCritical
            Exit Sub
        End If
        LoadPubKeysOut
        PopulateMemberList
        
    Case 0
        LoadPrivateAddressBook
        PopulateMemberList
        
        
End Select
End Sub

Private Sub Form_Activate()
Dim rs As Recordset
On Error GoTo ActivateError
If fLoading Then
    cmdAddressBook.Enabled = True
    Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
    rs.FindFirst "[List Name] = " & "'" & lblListName & "'"
    If rs.NoMatch Then Exit Sub
    Select Case rs("Member Type")
        Case 0
            m_MemberType = 0 ' Not set - Undefinded
            cmdAddressBook.ListIndex = 0
        Case 1
            m_MemberType = 1 'From Private Address Book
            cmdAddressBook.ListIndex = 0
            cmdAddressBook.Enabled = False
        Case 2
            m_MemberType = 2 ' From PGP Keyring
            cmdAddressBook.ListIndex = 1
            cmdAddressBook.Enabled = False
    End Select
    
    rs.Close
    fLoading = False
    'CreateListTable = True
End If
Exit Sub
ActivateError:
    MsgBox "An error has occured: " & Err.Description, vbCritical + vbApplicationModal
    Err.Clear
End Sub

Private Sub Form_Load()
Dim rs As Recordset
'Dim Win As New CWindow

'Win.Center Me, Null
'Win.OnTop(Me) = True
 
On Error GoTo BadList
'First Fill in the contacts
cmdAddressBook.AddItem "Private Address Book"
cmdAddressBook.AddItem "Public Keys"
'cmdAddressBook.AddItem "Mail Group"
fLoading = True

'Okay fill in the members of this list into the to box

BadList:
Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAddListMembers = Nothing
End Sub

Private Sub List1_DblClick()

    Call btnTo_Click

   
End Sub

Private Sub LoadPubKeysOut()
Dim BufferOut As String
Dim i As Long
Dim Count As Long
Dim BufferLen As Long
    
    BufferLen = spgpKeyRingCount() * 256
    BufferOut = String(BufferLen, Chr(0))
    i = spgpKeyRingID(BufferOut, BufferLen)
    Count = CountCRLF(BufferOut)
    Call ChopKeyProps(BufferOut, Count)
    
    For i = 0 To UBound(KeyArray())
        If KeyArray(i).KeyID <> "" Then List1.AddItem KeyArray(i).UserID 'KeyArray(i).KeyID & Chr(9) & KeyArray(i).UserID
    Next i
End Sub

Private Sub LoadPrivateAddressBook()
Dim rs As Recordset
Dim ItemRecord As String

    Set rs = DB.OpenRecordset("Contacts", dbOpenDynaset)
    List1.Clear
    If rs.EOF Then Exit Sub
        While Not rs.EOF
            ItemRecord = IIf(IsNull(rs("Contact Name")), "", rs("Contact Name") & " ")
            ItemRecord = ItemRecord & rs("Contact Email") '& "(ID=" & rs("Address Book ID") & ")"
            List1.AddItem ItemRecord 'rs("Contact Name") & vbTab & rs("Contact Email")
            rs.MoveNext
        Wend
    rs.Close
End Sub


Private Sub ToList_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
    If ToList.ListCount = 0 Then Exit Sub
    PIForm(gActivePIInstance).AddressList.RemoveContact ToList.List(ToList.ListIndex), CONTACT_TO_LIST
    ToList.RemoveItem (ToList.ListIndex)
    
End If
End Sub



Private Sub PopulateMemberList()
Dim rsList As Recordset
Dim rsMembers As Recordset
Dim lMemberID As Long

On Error GoTo BadPopulate
Set rsList = DB.OpenRecordset(lblListName, dbOpenDynaset)
'Set rsMembers = DB.OpenRecordset("Contacts", dbOpenDynaset)

While Not rsList.EOF
    'lMemberID = rsList("Member ID")
    'rsMembers.FindFirst "[Member ID] = " & lMemberID   '" & "'" & Tree.FolderName & "'"
    ToList.AddItem rsList("Contact Email")
    rsList.MoveNext
    DoEvents
Wend
rsList.Close
'rsMembers.Close
Exit Sub
BadPopulate:
    MsgBox "Can't populate member list: " & Err.Description, vbApplicationModal + vbCritical
    Err.Clear

End Sub

Public Sub DeleteAllMembers()
Dim rsList As Recordset

On Error Resume Next
Set rsList = DB.OpenRecordset(lblListName, dbOpenDynaset)

'Remove previous members and fill with new ones..
While Not rsList.EOF
    rsList.Delete
    rsList.MoveNext
Wend
rsList.Close
End Sub
