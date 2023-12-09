VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEmailAddressList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private AddressArray()
Private m_Index(4) As Integer ' This holds the number of entries for each List TYpe
Private m_DisplayName
'Private m_TotalContacts
Private Const DISPLAY_NAME As Integer = 1
Private Const FULL_NAME As Integer = 2
Private Const EMAIL_ADDRESS As Integer = 3
Private Const CONTACT_ATTRIBUTES As Integer = 4

Public Function AddContact(ContactID As Variant, iAddresseeList As Integer) As Long
Dim l_index As Integer
Dim i As Integer
'
'First check to see it it exists
' Return 0 if already there
If Not ContactID(3) = CONTACT_IN_MAILGROUP Then
   For i = 1 To m_Index(iAddresseeList)
        If AddressArray(EMAIL_ADDRESS, iAddresseeList, i) = Trim(ContactID(2)) Then
            AddContact = 0
            Exit Function
        End If
    Next
Else
    For i = 1 To m_Index(iAddresseeList)
        If AddressArray(DISPLAY_NAME, iAddresseeList, i) = Trim(ContactID(0)) Then
            AddContact = 0
            Exit Function
        End If
    Next
End If

m_Index(iAddresseeList) = m_Index(iAddresseeList) + 1
m_Index(CONTACT_ALL_LIST) = m_Index(CONTACT_TO_LIST) + m_Index(CONTACT_CC_LIST)
l_index = m_Index(iAddresseeList)
ReDim Preserve AddressArray(4, CONTACT_ARRAY_LIMIT, m_Index(CONTACT_ALL_LIST))
AddressArray(DISPLAY_NAME, iAddresseeList, l_index) = ContactID(0) 'Display Name
AddressArray(FULL_NAME, iAddresseeList, l_index) = ContactID(1) 'FullName
AddressArray(EMAIL_ADDRESS, iAddresseeList, l_index) = ContactID(2) 'Email Address
'Status 0 new contact
'Status 1 in Contact Database
'Status 2 on PGP KeyRing
AddressArray(CONTACT_ATTRIBUTES, iAddresseeList, l_index) = ContactID(3)

'Now also add to ALL list for non mail groups
'If Not ContactID(3) = CONTACT_IN_MAILGROUP Then
    'AddressArray(DISPLAY_NAME, CONTACT_ALL_LIST, m_Index(CONTACT_ALL_LIST)) = ContactID(0) 'Display Name
    'AddressArray(FULL_NAME, CONTACT_ALL_LIST, m_Index(CONTACT_ALL_LIST)) = ContactID(1) 'FullName
    'AddressArray(EMAIL_ADDRESS, CONTACT_ALL_LIST, m_Index(CONTACT_ALL_LIST)) = ContactID(2) 'Email Address
    'AddressArray(CONTACT_ATTRIBUTES, CONTACT_ALL_LIST, m_Index(CONTACT_ALL_LIST)) = ContactID(3) 'Status 0
'End If

AddContact = m_Index(iAddresseeList) 'Return this
End Function

Private Sub Class_Initialize()
'AddressArray is organised as follows
' Array(VV,TT,II)
' First element is the field value
'   1: DISPLAY NAME (This is normally used for mail groups as the Full Name for
'      a contact is the same as the Full Name
'   2: FULL NAME
'   3: EMAIL ADDRESS
'   4: ATTRIBUTE:
'       CONTACT_IN_DB = 1
'       CONTACT_ON_PGPKEYRING = 2
'       CONTACT_IN_MAILGROUP = 3
' Second element is the type of Address list ie
'   TO Field, CC Field, and combined ie ALL
' Third element is the Index
' Lastly, the element (0,TT,0) is the number of elements in the Type TT

ReDim AddressArray(4, CONTACT_ARRAY_LIMIT, 0)
m_Index(0) = 0
m_Index(1) = 0
m_Index(2) = 0
m_Index(3) = 0
End Sub
Public Sub Initialise()
ReDim AddressArray(4, CONTACT_ARRAY_LIMIT, 0)

m_Index(0) = 0
m_Index(1) = 0
m_Index(2) = 0
m_Index(3) = 0
End Sub

Public Function GetContactEmailAddress(ContactName As String, iAddresseeList As Integer) As String
Dim i As Integer
For i = 1 To m_Index(iAddresseeList)
    If Trim(AddressArray(DISPLAY_NAME, iAddresseeList, i)) = Trim(ContactName) Then
        GetContactEmailAddress = AddressArray(3, iAddresseeList, i)
        Exit Function
    End If
Next
GetContactEmailAddress = ""
End Function
Public Function GetContactDisplayName(ContactName As String, iAddresseeList As Integer) As String
Dim i As Integer
For i = 1 To m_Index(iAddresseeList)
    If Trim(AddressArray(DISPLAY_NAME, iAddresseeList, i)) = Trim(ContactName) Then
        GetContactDisplayName = AddressArray(DISPLAY_NAME, iAddresseeList, i)
        Exit Function
    End If
Next
GetContactDisplayName = "" 'Not there
End Function
Public Function RemoveContact(ContactName As String, iAddresseeList As Integer) As String
Dim i As Integer
Dim J As Integer
Dim Found As Boolean
On Error Resume Next
RemoveContact = "" 'Default is error condition
For i = 1 To m_Index(iAddresseeList)
    If Trim(AddressArray(DISPLAY_NAME, iAddresseeList, i)) = Trim(ContactName) Then
        Found = True
        Exit For
    End If
Next

If Found Then
    For J = i To m_Index(iAddresseeList) - 1
        AddressArray(1, iAddresseeList, J) = AddressArray(1, iAddresseeList, J + 1) 'Display Name
        AddressArray(2, iAddresseeList, J) = AddressArray(2, iAddresseeList, J + 1)  'FullName
        AddressArray(3, iAddresseeList, J) = AddressArray(3, iAddresseeList, J + 1)  'Email Address
        AddressArray(4, iAddresseeList, J) = AddressArray(4, iAddresseeList, J + 1)  'status 0 is new, Status 1 in Contact Database                                                                              'Status 2 on PGP KeyRing    Next
    Next
    m_Index(iAddresseeList) = m_Index(iAddresseeList) - 1
    m_Index(CONTACT_ALL_LIST) = m_Index(CONTACT_TO_LIST) + m_Index(CONTACT_CC_LIST)
    
    ReDim Preserve AddressArray(4, 2, m_Index(CONTACT_ALL_LIST))
    RemoveContact = ContactName
End If
End Function
Public Function GetContactAttributes(ContactName As String, iAddresseeList As Integer) As Integer
Dim i As Integer
For i = 1 To m_Index(iAddresseeList)
    If Trim(AddressArray(3, iAddresseeList, i)) = Trim(ContactName) Then
        GetContactAttributes = AddressArray(4, iAddresseeList, i)
        Exit Function
    End If
Next
End Function
Public Function GetContactType1(ContactName As String, iAddresseeList As Integer) As String
Dim i As Integer
Exit Function ' this does nothing at presetn
For i = 1 To m_Index(iAddresseeList)
    If Trim(AddressArray(1, iAddresseeList, 1)) = Trim(ContactName) Then
        GetContactType1 = AddressArray(4, iAddresseeList, i)
        Exit Function
    End If
Next
End Function
Public Function GetCount(ContactName As String, iAddresseeList As Integer) As Integer

GetCount = m_Index(iAddresseeList)

End Function
Public Function GetAllContacts(iAddresseeList As Integer) As Variant
AddressArray(0, iAddresseeList, 0) = m_Index(iAddresseeList) 'm_Index
GetAllContacts = AddressArray
End Function

Public Function GetAllPrivateAddressBookContacts(iAddresseeList As Integer) As Variant

Dim PGPArray() As Variant
Dim iNumNonPGP As Integer
Dim i As Integer
iNumNonPGP = 0
ReDim Preserve PGPArray(4, iAddresseeList, 0)
PGPArray(0, iAddresseeList, 0) = 0
For i = 1 To m_Index(iAddresseeList)
    If AddressArray(CONTACT_ATTRIBUTES, iAddresseeList, i) = CONTACT_IN_DB Then
        iNumNonPGP = iNumNonPGP + 1
        ReDim Preserve PGPArray(4, iAddresseeList, iNumNonPGP)
        PGPArray(1, iAddresseeList, iNumNonPGP) = AddressArray(1, iAddresseeList, i)  'Display Name
        PGPArray(2, iAddresseeList, iNumNonPGP) = AddressArray(2, iAddresseeList, i)  'FullName
        PGPArray(3, iAddresseeList, iNumNonPGP) = AddressArray(3, iAddresseeList, i) 'Email Address
        PGPArray(4, iAddresseeList, iNumNonPGP) = AddressArray(4, iAddresseeList, i) 'status 0 is new, Status 1 in Contact Database
        PGPArray(0, iAddresseeList, 0) = iNumNonPGP
    End If
Next
GetAllPrivateAddressBookContacts = PGPArray
End Function
Public Function GetAllMailGroupContacts(iAddresseeList As Integer) As Variant

Dim MailGroupArray() As Variant
Dim iNumMailGroupMembers As Integer
Dim i As Integer
iNumMailGroupMembers = 0
ReDim Preserve MailGroupArray(4, iAddresseeList, 0)
MailGroupArray(0, iAddresseeList, 0) = 0
For i = 1 To m_Index(iAddresseeList)
    If AddressArray(CONTACT_ATTRIBUTES, iAddresseeList, i) = CONTACT_IN_MAILGROUP Then
        iNumMailGroupMembers = iNumMailGroupMembers + 1
        ReDim Preserve MailGroupArray(4, iAddresseeList, iNumMailGroupMembers)
        MailGroupArray(1, iAddresseeList, iNumMailGroupMembers) = AddressArray(1, iAddresseeList, i)  'Display Name
        MailGroupArray(2, iAddresseeList, iNumMailGroupMembers) = AddressArray(2, iAddresseeList, i)  'FullName
        MailGroupArray(3, iAddresseeList, iNumMailGroupMembers) = AddressArray(3, iAddresseeList, i) 'Email Address
        MailGroupArray(4, iAddresseeList, iNumMailGroupMembers) = AddressArray(4, iAddresseeList, i) 'status 0 is new, Status 1 in Contact Database
        MailGroupArray(0, iAddresseeList, 0) = iNumMailGroupMembers
    End If
Next
GetAllMailGroupContacts = MailGroupArray
End Function
Public Function GetAllMailGroups(iAddresseeList As Integer) As Variant
GetNumberMailGroups
Dim MailGroupArray() As Variant
Dim iNumMailGroups As Integer
Dim i As Integer
iNumMailGroups = 0
ReDim Preserve MailGroupArray(0) 'MailGroupArray(4, iAddresseeList, 0)
'MailGroupArray(0, iAddresseeList, 0) = 0
For i = 1 To m_Index(iAddresseeList)
    If AddressArray(CONTACT_ATTRIBUTES, iAddresseeList, i) = CONTACT_IN_MAILGROUP Then
        iNumMailGroups = iNumMailGroups + 1
        ReDim Preserve MailGroupArray(iNumMailGroups) 'MailGroupArray(4, iAddresseeList, iNumMailGroups)
        'MailGroupArray(DISPLAY_NAME, iAddresseeList, iNumMailGroups) = AddressArray(1, iAddresseeList, i)  'Display Name
        'MailGroupArray(0, iAddresseeList, 0) = iNumMailGroups
        MailGroupArray(iNumMailGroups) = AddressArray(DISPLAY_NAME, iAddresseeList, i)  'Display Name
        
        MailGroupArray(0) = iNumMailGroups
    End If
Next
GetAllMailGroups = MailGroupArray
End Function
Public Function GetNumberMailGroups() As Integer
'Dim MailGroupArray() As Variant
Dim iNumMailGroups As Integer
Dim i As Integer

iNumMailGroups = 0
For i = 1 To m_Index(CONTACT_TO_LIST)
    If AddressArray(CONTACT_ATTRIBUTES, CONTACT_TO_LIST, i) = CONTACT_IN_MAILGROUP Then
        iNumMailGroups = iNumMailGroups + 1
    End If
Next
For i = 1 To m_Index(CONTACT_CC_LIST)
    If AddressArray(CONTACT_ATTRIBUTES, CONTACT_CC_LIST, i) = CONTACT_IN_MAILGROUP Then
        iNumMailGroups = iNumMailGroups + 1
    End If
Next
GetNumberMailGroups = iNumMailGroups
End Function
Public Function GetNumberPGPContacts() As Integer
'Dim MailGroupArray() As Variant
Dim iNumPGPContacts As Integer
Dim i As Integer

iNumPGPContacts = 0
For i = 1 To m_Index(CONTACT_TO_LIST)
    If AddressArray(CONTACT_ATTRIBUTES, CONTACT_TO_LIST, i) = CONTACT_ON_PGPKEYRING Then
        iNumPGPContacts = iNumPGPContacts + 1
    End If
Next
For i = 1 To m_Index(CONTACT_CC_LIST)
    If AddressArray(CONTACT_ATTRIBUTES, CONTACT_CC_LIST, i) = CONTACT_ON_PGPKEYRING Then
        iNumPGPContacts = iNumPGPContacts + 1
    End If
Next
GetNumberPGPContacts = iNumPGPContacts
End Function
Public Function GetNumberNonPGPContacts() As Integer
'Dim MailGroupArray() As Variant
Dim iNumNonPGPContacts As Integer
Dim i As Integer

iNumNonPGPContacts = 0
For i = 1 To m_Index(CONTACT_TO_LIST)
    If Not AddressArray(CONTACT_ATTRIBUTES, CONTACT_TO_LIST, i) = CONTACT_ON_PGPKEYRING Then
        iNumNonPGPContacts = iNumNonPGPContacts + 1
    End If
Next
For i = 1 To m_Index(CONTACT_CC_LIST)
    If Not AddressArray(CONTACT_ATTRIBUTES, CONTACT_CC_LIST, i) = CONTACT_ON_PGPKEYRING Then
        iNumNonPGPContacts = iNumNonPGPContacts + 1
    End If
Next
GetNumberNonPGPContacts = iNumNonPGPContacts
End Function
Public Function GetAllPGPContacts(iAddresseeList As Integer) As Variant
Dim PGPArray() As Variant
Dim iNumPGP As Integer
Dim i As Integer
iNumPGP = 0
ReDim Preserve PGPArray(4, iAddresseeList, 0)
PGPArray(0, iAddresseeList, 0) = 0

Select Case iAddresseeList

    Case CONTACT_TO_LIST, CONTACT_CC_LIST
        For i = 1 To m_Index(iAddresseeList)
            If AddressArray(4, iAddresseeList, i) = CONTACT_ON_PGPKEYRING Then
                iNumPGP = iNumPGP + 1
                ReDim Preserve PGPArray(4, iAddresseeList, iNumPGP)
                PGPArray(1, iAddresseeList, iNumPGP) = AddressArray(1, iAddresseeList, i)  'Display Name
                PGPArray(2, iAddresseeList, iNumPGP) = AddressArray(2, iAddresseeList, i)  'FullName
                PGPArray(3, iAddresseeList, iNumPGP) = AddressArray(3, iAddresseeList, i) 'Email Address
                PGPArray(4, iAddresseeList, iNumPGP) = AddressArray(4, iAddresseeList, i) 'status 0 is new, Status 1 in Contact Database
                PGPArray(0, iAddresseeList, 0) = iNumPGP
            End If
        Next
    Case CONTACT_ALL_LIST
        For i = 1 To m_Index(CONTACT_TO_LIST)
            If AddressArray(4, CONTACT_TO_LIST, i) = CONTACT_ON_PGPKEYRING Then
                iNumPGP = iNumPGP + 1
                ReDim Preserve PGPArray(4, iAddresseeList, iNumPGP)
                PGPArray(1, iAddresseeList, iNumPGP) = AddressArray(1, CONTACT_TO_LIST, i)  'Display Name
                PGPArray(2, iAddresseeList, iNumPGP) = AddressArray(2, CONTACT_TO_LIST, i)  'FullName
                PGPArray(3, iAddresseeList, iNumPGP) = AddressArray(3, CONTACT_TO_LIST, i) 'Email Address
                PGPArray(4, iAddresseeList, iNumPGP) = AddressArray(4, CONTACT_TO_LIST, i) 'status 0 is new, Status 1 in Contact Database
                PGPArray(0, iAddresseeList, 0) = iNumPGP
            End If
        Next
        For i = 1 To m_Index(CONTACT_CC_LIST)
            If AddressArray(4, CONTACT_CC_LIST, i) = CONTACT_ON_PGPKEYRING Then
                iNumPGP = iNumPGP + 1
                ReDim Preserve PGPArray(4, iAddresseeList, iNumPGP)
                PGPArray(1, iAddresseeList, iNumPGP) = AddressArray(1, CONTACT_CC_LIST, i)  'Display Name
                PGPArray(2, iAddresseeList, iNumPGP) = AddressArray(2, CONTACT_CC_LIST, i)  'FullName
                PGPArray(3, iAddresseeList, iNumPGP) = AddressArray(3, CONTACT_CC_LIST, i) 'Email Address
                PGPArray(4, iAddresseeList, iNumPGP) = AddressArray(4, CONTACT_CC_LIST, i) 'status 0 is new, Status 1 in Contact Database
                PGPArray(0, iAddresseeList, 0) = iNumPGP
            End If
        Next
End Select
GetAllPGPContacts = PGPArray
End Function

Public Function GetListAllNonPGPContacts(iAddresseeList As Integer) As String
Dim i As Integer
Dim sAddressList As String

For i = 1 To m_Index(iAddresseeList)
    If Not AddressArray(CONTACT_ATTRIBUTES, iAddresseeList, i) = CONTACT_ON_PGPKEYRING Then
        If i > 1 Then sAddressList = sAddressList & ", "
            sAddressList = LTrim(sAddressList & AddressArray(FULL_NAME, iAddresseeList, i) & " <" & AddressArray(EMAIL_ADDRESS, iAddresseeList, i)) & ">"
    End If
    'sAddressList = ""
Next
GetListAllNonPGPContacts = sAddressList
End Function
Public Function GetListAllPGPContacts(iAddresseeList As Integer) As String
Dim i As Integer
Dim sAddressList As String
'sAddressList = ""

For i = 1 To m_Index(iAddresseeList)
    If AddressArray(4, iAddresseeList, i) = CONTACT_ON_PGPKEYRING Then
        If i > 1 Then sAddressList = sAddressList & ", "
           ' sAddressList = sAddressList & IIf(IsNull(AddressArray(2, iAddresseeList, i)) = True, "", AddressArray(2, iAddresseeList, i))
            'If sAddressList = "" Then
              '  sAddressList = AddressArray(3, iAddresseeList, i)
            'Else
                sAddressList = LTrim(sAddressList & AddressArray(FULL_NAME, iAddresseeList, i) & " <" & AddressArray(EMAIL_ADDRESS, iAddresseeList, i)) & ">"
            'End If
    End If
Next
GetListAllPGPContacts = sAddressList
End Function
Public Function GetListAllContacts(iAddresseeList As Integer) As String
Dim i As Integer
Dim sAddressList As String

For i = 1 To m_Index(iAddresseeList)
    'If Not AddressArray(4, iAddresseeList, i) = CONTACT_ON_PGPKEYRING Then
        If Not sAddressList = "" Then sAddressList = sAddressList & ", "
            sAddressList = sAddressList & IIf(IsNull(AddressArray(2, iAddresseeList, i)) = True, "", AddressArray(2, iAddresseeList, i))
            If sAddressList = "" Then
                sAddressList = AddressArray(3, iAddresseeList, i)
            Else
                sAddressList = sAddressList & " <" & AddressArray(3, iAddresseeList, i) & ">"
        End If
    'End If
Next
GetListAllContacts = sAddressList
End Function
