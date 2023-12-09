VERSION 5.00
Begin VB.Form frmEditAddressBook 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create or Edit Address Book Entries"
   ClientHeight    =   5700
   ClientLeft      =   1545
   ClientTop       =   1890
   ClientWidth     =   5430
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
   ScaleHeight     =   5700
   ScaleWidth      =   5430
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
      Left            =   4230
      TabIndex        =   4
      Top             =   5310
      Width           =   975
   End
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
      Left            =   2970
      TabIndex        =   3
      Top             =   4770
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
      TabIndex        =   2
      Top             =   4770
      Width           =   1635
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
      ItemData        =   "Edit Address Book.frx":0000
      Left            =   240
      List            =   "Edit Address Book.frx":0002
      TabIndex        =   1
      Text            =   "cmbAddressBook"
      ToolTipText     =   "Select Address Book"
      Top             =   600
      Width           =   4815
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
      TabIndex        =   0
      Top             =   1020
      Width           =   4845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Using this dialog box you can add, edit and remove contacts from your address book."
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
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click on address to edit it."
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
      Left            =   1470
      TabIndex        =   5
      Top             =   4470
      Width           =   2595
   End
End
Attribute VB_Name = "frmEditAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fPubKeySelected As Boolean
Private m_MemberType As Integer
Private Sub btnNewContact_Click()
frmEditAddressBookEntry.btnSave.Enabled = True
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
  For i = 0 To List1.ListIndex - 1
    rs.MoveNext
  Next
    rs.Edit
    rs.Delete
    rs.Update
    rs.Close
    LoadPrivateAddressBook
End Sub



Private Sub cmdAddressBook_Click()
List1.Clear

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
       ' FillListsWithExistingContacts (CONTACT_TO_LIST)
        'FillListsWithExistingContacts (CONTACT_CC_LIST)
    Case 0
        'Private Address Book
        LoadPrivateAddressBook
        m_MemberType = 1
        fPubKeySelected = False
        btnNewContact.Enabled = True
        btnRemoveContact.Enabled = True
       ' FillListsWithExistingContacts (CONTACT_TO_LIST)
        'FillListsWithExistingContacts (CONTACT_CC_LIST)
        
    Case 2
        LoadMailGroupList List1
        fPubKeySelected = False
        m_MemberType = 3
        btnNewContact.Enabled = False
        btnRemoveContact.Enabled = False
       ' FillListsWithExistingContacts (CONTACT_TO_LIST)
       ' FillListsWithExistingContacts (CONTACT_CC_LIST)
        
End Select
End Sub

Private Sub Command1_Click()


Unload Me

End Sub



Private Sub Form_Load()
Dim Win As New CWindow

    
    On Error Resume Next
    Win.Center Me, Null
   ' Win.OnTop(Me) = True
    
    cmdAddressBook.AddItem "Private Address Book"
    cmdAddressBook.AddItem "Public Keys"
    cmdAddressBook.AddItem "Mail Group"
    LoadPrivateAddressBook
    DoEvents
    cmdAddressBook.ListIndex = 0
    
   
    Exit Sub

'KeyError:
   ' MsgBox "Not able to load User Id selection form.  Error is: " & Err.Description, vbApplicationModal, gPiStr
    'Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSelectUserID = Nothing
End Sub

Private Sub List1_DblClick()
'Dim s As String
's = List1.List(List1.ListIndex)
If cmdAddressBook.ListIndex = 0 Then frmEditAddressBookEntry.btnSave.Enabled = True
frmEditAddressBookEntry.txtName(0) = StripFullName(List1.List(List1.ListIndex))
frmEditAddressBookEntry.txtName(1) = StripEMailAddress(List1.List(List1.ListIndex))
frmEditAddressBookEntry.Show vbModal
'List1.Clear
LoadPrivateAddressBook
   
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






