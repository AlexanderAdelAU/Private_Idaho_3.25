VERSION 5.00
Begin VB.Form frmSelectUserID 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select User Addressee"
   ClientHeight    =   5790
   ClientLeft      =   1545
   ClientTop       =   1890
   ClientWidth     =   8925
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
   ScaleHeight     =   5790
   ScaleWidth      =   8925
   Begin VB.CommandButton btnRemoveContact 
      Caption         =   "Remove Contact"
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
      Left            =   2430
      TabIndex        =   11
      Top             =   5100
      Width           =   1575
   End
   Begin VB.CommandButton btnNewContact 
      Caption         =   "New Contact"
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
      Left            =   450
      TabIndex        =   10
      Top             =   5100
      Width           =   1635
   End
   Begin VB.CommandButton btnCC 
      Caption         =   "CC->"
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
      Left            =   4230
      TabIndex        =   9
      ToolTipText     =   "Click to select"
      Top             =   3450
      Width           =   585
   End
   Begin VB.ListBox CCList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   4890
      TabIndex        =   8
      Top             =   3390
      Width           =   3855
   End
   Begin VB.CommandButton btnTo 
      Caption         =   "To->"
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
      TabIndex        =   7
      ToolTipText     =   "Click to select"
      Top             =   1560
      Width           =   585
   End
   Begin VB.ListBox ToList 
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
      Height          =   1620
      Left            =   4860
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
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
      ItemData        =   "SelectUserID.frx":0000
      Left            =   240
      List            =   "SelectUserID.frx":0002
      TabIndex        =   5
      Text            =   "cmbAddressBook"
      ToolTipText     =   "Select Address Book"
      Top             =   1140
      Width           =   3825
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
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
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
      Left            =   7830
      TabIndex        =   2
      Top             =   5280
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
      Left            =   6630
      TabIndex        =   0
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "You did not specify a key to encrypt with in the 'To:' box"
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
      TabIndex        =   4
      Top             =   180
      Width           =   4395
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a User ID from the list to encrypt the message with."
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
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   540
      Width           =   5175
   End
End
Attribute VB_Name = "frmSelectUserID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fPubKeySelected As Boolean
Private m_MemberType As Integer

Private Sub btnCC_Click()
Dim ccAddressee As String
Dim ContactEmailAddress As String
Dim ContactFullName As String
Dim VarArray As Variant
Dim lIndex As Long

'
'This tag is used to determine the direction in the list1.doubleclick
'
If List1.ListIndex = -1 Then Exit Sub
btnTo.Tag = ""
btnCC.Tag = "Selected"

Select Case cmdAddressBook.ListIndex
    Case 1
        'Use Public Keys
        CCList.Enabled = True
        CCList.FontUnderline = True
        ContactEmailAddress = StripEMailAddress(List1.List(List1.ListIndex))
        ContactFullName = StripFullName(List1.List(List1.ListIndex))
        VarArray = Array(ContactFullName & " " & ContactEmailAddress, ContactFullName, ContactEmailAddress, CONTACT_ON_PGPKEYRING)
        lIndex = PIForm(gActivePIInstance).AddressList.AddContact(VarArray, CONTACT_CC_LIST)
        CCList.AddItem IIf(ContactFullName = "", ContactEmailAddress, ContactFullName) 'Display name
    Case 0
        'Use Address Book
        CCList.FontUnderline = True
        ccAddressee = PIForm(gActivePIInstance).LookUpContactRecord(List1.List(List1.ListIndex), CONTACT_CC_LIST)
        If ccAddressee = "" Then CCList.FontUnderline = False
        CCList.AddItem ccAddressee
    Case 2
        'Use Mail List
        CCList.Enabled = True
        CCList.FontUnderline = True
        Dim sGroupName As String
        'Add the Group name
        sGroupName = List1.List(List1.ListIndex)
        VarArray = Array(sGroupName, "", "", CONTACT_IN_MAILGROUP)
        lIndex = PIForm(gActivePIInstance).AddressList.AddContact(VarArray, CONTACT_CC_LIST)
        If Not lIndex = 0 Then CCList.AddItem List1.List(List1.ListIndex)
End Select



End Sub


Private Sub btnNewContact_Click()
frmEditAddressBookEntry.Show vbModal
LoadPrivateAddressBook
End Sub

Private Sub btnRemoveContact_Click()
Dim rs As Recordset
Dim i As Integer
  On Error Resume Next
  
  Set rs = DB.OpenRecordset("Contacts", dbOpenDynaset)
  
  If Not rs.EOF Then
    rs.MoveFirst
  Else
    Beep
    Exit Sub
  End If
  'Dim s As String
  's = rs("Contact name")
  For i = 0 To List1.ListIndex - 1
    rs.MoveNext
  Next
    rs.Edit
    rs.Delete
    rs.Update
    rs.Close
    LoadPrivateAddressBook
End Sub

Private Sub btnTo_Click()
Dim toAddressee As String
Dim VarArray As Variant
'Dim MemberArray As Variant
Dim ContactEmailAddress As String
Dim ContactFullName As String
Dim lIndex As Long
Dim i As Integer

'This tag is used to determine the direction in the list1.doubleclick
btnTo.Tag = "Selected"
'

If List1.ListIndex = -1 Then Exit Sub
'If List1.List(List1.ListIndex) = "" Then Exit Sub
Select Case cmdAddressBook.ListIndex
    Case 1
        'Use Public Keys
        ToList.Enabled = True
        ToList.FontUnderline = True
        ContactEmailAddress = StripEMailAddress(List1.List(List1.ListIndex))
        ContactFullName = StripFullName(List1.List(List1.ListIndex))
        VarArray = Array(ContactFullName & " " & ContactEmailAddress, ContactFullName, ContactEmailAddress, CONTACT_ON_PGPKEYRING)
        lIndex = PIForm(gActivePIInstance).AddressList.AddContact(VarArray, CONTACT_TO_LIST)
        ToList.AddItem IIf(ContactFullName = "", ContactEmailAddress, ContactFullName) 'Display name
    Case 0
        'Use Address Book
        ToList.FontUnderline = True
        toAddressee = PIForm(gActivePIInstance).LookUpContactRecord(List1.List(List1.ListIndex), CONTACT_TO_LIST)
        If toAddressee = "" Then ToList.FontUnderline = False
        ToList.AddItem toAddressee
    Case 2
        'Use Mail List
        Dim MemberArray As Variant
        Dim sGroupName As String
        Dim DisplayName As String
        Dim iMailGroupIndex
        
        ToList.Enabled = True
        ToList.FontUnderline = True
                
        'MemberArray = GetMembersOfGroup(List1.List(List1.ListIndex))
        'Add the Group name
        sGroupName = List1.List(List1.ListIndex)
        VarArray = Array(sGroupName, "", "", CONTACT_IN_MAILGROUP)
        lIndex = PIForm(gActivePIInstance).AddressList.AddContact(VarArray, CONTACT_TO_LIST)
        If Not lIndex = 0 Then ToList.AddItem List1.List(List1.ListIndex)
End Select

End Sub

Private Sub CCList_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyDelete Then
    If CCList.ListCount = 0 Then Exit Sub
    PIForm(gActivePIInstance).AddressList.RemoveContact ToList.List(ToList.ListIndex), CONTACT_CC_LIST
    CCList.RemoveItem (CCList.ListIndex)
    
End If

End Sub



Private Sub cmdAddressBook_Click()
List1.Clear
ToList.Clear
CCList.Clear
Select Case cmdAddressBook.ListIndex

    Case 1 'If cmdAddressBook.ListIndex = 1 Then
        If Not PGP_SDKPresent Or gPGPVersion = NoPGP Then
            MsgBox "Can't find PGP SDK in Windows System directory or you have disabled PGP.  You must install PGP to be able select from the public key list", vbApplicationModal + vbCritical
            Exit Sub
        End If
        LoadPubKeysOut
        fPubKeySelected = True
        m_MemberType = 2
        btnNewContact.Enabled = False
        btnRemoveContact.Enabled = False
        FillListsWithExistingContacts (CONTACT_TO_LIST)
        FillListsWithExistingContacts (CONTACT_CC_LIST)
    Case 0
        'Private Address Book
        LoadPrivateAddressBook
        m_MemberType = 1
        fPubKeySelected = False
        btnNewContact.Enabled = True
        btnRemoveContact.Enabled = True
        FillListsWithExistingContacts (CONTACT_TO_LIST)
        FillListsWithExistingContacts (CONTACT_CC_LIST)
        
    Case 2
        LoadMailGroupList List1
        fPubKeySelected = False
        m_MemberType = 3
        btnNewContact.Enabled = False
        btnRemoveContact.Enabled = False
        FillListsWithExistingContacts (CONTACT_TO_LIST)
        FillListsWithExistingContacts (CONTACT_CC_LIST)
        
End Select
End Sub

Private Sub Command1_Click()

Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer
Dim VarArray() As Variant
    
PIForm(gActivePIInstance).txtTo.SelStart = 0
PIForm(gActivePIInstance).txtTo.Text = ""

    VarArray = PIForm(gActivePIInstance).AddressList.GetAllContacts(CONTACT_TO_LIST)
    For i = 1 To VarArray(0, CONTACT_TO_LIST, 0)
        If i > 1 Then
            PIForm(gActivePIInstance).txtTo.SelText = "; "
            PIForm(gActivePIInstance).txtTo.SelUnderline = True
            PIForm(gActivePIInstance).txtTo.SelText = VarArray(1, CONTACT_TO_LIST, i)
        Else
            PIForm(gActivePIInstance).txtTo.SelUnderline = True
            PIForm(gActivePIInstance).txtTo.SelText = VarArray(1, CONTACT_TO_LIST, i)
        End If
        PIForm(gActivePIInstance).txtTo.SelUnderline = False
    Next
    
    PIForm(gActivePIInstance).txtCC.SelStart = 0
    PIForm(gActivePIInstance).txtCC.Text = ""
    VarArray = PIForm(gActivePIInstance).AddressList.GetAllContacts(CONTACT_CC_LIST)
    For i = 1 To VarArray(0, CONTACT_CC_LIST, 0)
        If i > 1 Then
            PIForm(gActivePIInstance).txtCC.SelText = "; "
            PIForm(gActivePIInstance).txtCC.SelUnderline = True
            PIForm(gActivePIInstance).txtCC.SelText = VarArray(1, CONTACT_CC_LIST, i)
        Else
            PIForm(gActivePIInstance).txtCC.SelUnderline = True
            PIForm(gActivePIInstance).txtCC.SelText = VarArray(1, CONTACT_CC_LIST, i)
        End If
        PIForm(gActivePIInstance).txtCC.SelUnderline = False
    Next
   
'For i = 0 To ToList.ListCount - 1
   ' PIForm(gActivePIInstance).txtTo.SelUnderline = False
    'If i > 0 Then
            
          '  PIForm(gActivePIInstance).txtTo.SelText = ", "
           ' PIForm(gActivePIInstance).txtTo.SelUnderline = True
           ' PIForm(gActivePIInstance).txtTo.SelText = ToList.List(i)
   ' Else
        '    PIForm(gActivePIInstance).txtTo.SelUnderline = True
          '  PIForm(gActivePIInstance).txtTo.SelText = ToList.List(i)
            
   ' End If
   ' PIForm(gActivePIInstance).txtTo.SelUnderline = False
'Next
'Now CC List
'PIForm(gActivePIInstance).txtCC.SelStart = 0
'PIForm(gActivePIInstance).txtCC.Text = ""
'For i = 0 To CCList.ListCount - 1
   ' PIForm(gActivePIInstance).txtCC.SelUnderline = False
    'If i > 0 Then
            
          '  PIForm(gActivePIInstance).txtCC.SelText = ", "
           ' PIForm(gActivePIInstance).txtCC.SelUnderline = True
           ' PIForm(gActivePIInstance).txtCC.SelText = CCList.List(i)
   ' Else
         '   PIForm(gActivePIInstance).txtCC.SelUnderline = True
         '   PIForm(gActivePIInstance).txtCC.SelText = CCList.List(i)
            
    'End If
    'PIForm(gActivePIInstance).txtCC.SelUnderline = False
'Next


Unload Me

End Sub

Private Sub Command2_Click()
     Unload Me
    gCancelAction = True
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
Dim FileNum As Integer
Dim TextLine As String
Dim tmpstr As String
Dim Pos As Integer
    
Win.Center Me, Null
'Win.OnTop(Me) = True
    
    ToList.Clear
    CCList.Clear
    fPubKeySelected = False
    'Default to Pubkeys.out
    cmdAddressBook.AddItem "Private Address Book"
    cmdAddressBook.AddItem "Public Keys"
    cmdAddressBook.AddItem "Mail Group"
   ' LoadPrivateAddressBook
    DoEvents
    cmdAddressBook.ListIndex = 0
    
    'This tag is used to determine the direction in the list1.doubleclick
    btnTo.Tag = "Selected" 'Default
    'FillListsWithExistingContacts (CONTACT_TO_LIST)
    'FillListsWithExistingContacts (CONTACT_CC_LIST)
    Exit Sub

KeyError:
    MsgBox "Not able to load User Id selection form.  Error is: " & Err.Description, vbApplicationModal, gPiStr
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSelectUserID = Nothing
End Sub

Private Sub List1_DblClick()
If btnTo.Tag = "Selected" Then
    Call btnTo_Click
Else
    Call btnCC_Click
End If
   
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
            ItemRecord = IIf(IsNull(rs("Contact Name")), "", rs("Contact Name"))
            ItemRecord = ItemRecord & " " & rs("Contact Email")
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

Public Sub FillListsWithExistingContacts(iAddressList As Integer)
 Dim VarArray As Variant
 Dim i As Long
 On Error Resume Next
 'the (0,0) index holds the number of entries
 
 Select Case m_MemberType
 
    Case 1
        VarArray = PIForm(gActivePIInstance).AddressList.GetAllPrivateAddressBookContacts(iAddressList)
    Case 2
        VarArray = PIForm(gActivePIInstance).AddressList.GetAllPGPContacts(iAddressList)
    Case 3
        VarArray = PIForm(gActivePIInstance).AddressList.GetAllMailGroupContacts(iAddressList)
 End Select
 
 If iAddressList = CONTACT_TO_LIST Then
    ToList.Clear
    ToList.FontUnderline = True
    For i = 1 To VarArray(0, iAddressList, 0)
        ToList.AddItem VarArray(1, iAddressList, i)
    Next
    ToList.FontUnderline = False
    ToList.Enabled = True
End If
If iAddressList = CONTACT_CC_LIST Then
    CCList.Clear
    CCList.FontUnderline = True
    For i = 1 To VarArray(0, iAddressList, 0)
        CCList.AddItem VarArray(1, iAddressList, i)
    Next
    CCList.FontUnderline = False
    CCList.Enabled = True
End If
'End Select
End Sub


