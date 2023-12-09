VERSION 5.00
Begin VB.Form frmAddressBook 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Private Address Book Entries"
   ClientHeight    =   3060
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   6180
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
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   6180
   Begin VB.TextBox txtName 
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
      Index           =   1
      Left            =   1830
      TabIndex        =   5
      Top             =   1380
      Width           =   3705
   End
   Begin VB.TextBox txtName 
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
      Left            =   1830
      TabIndex        =   4
      Top             =   870
      Width           =   3705
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4560
      TabIndex        =   1
      Top             =   2070
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Save"
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
      Left            =   3390
      TabIndex        =   0
      Top             =   2070
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   180
      Picture         =   "AddressBook.frx":0000
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your new contact details in the fields below and then press okay."
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
      Index           =   2
      Left            =   1170
      TabIndex        =   6
      Top             =   60
      Width           =   4395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
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
      Left            =   330
      TabIndex        =   3
      Top             =   1470
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
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
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   1245
   End
   Begin VB.Menu AddressEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu AddressEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu AdressEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu AddressEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu WebMenu 
      Caption         =   "&Web"
      Visible         =   0   'False
      Begin VB.Menu WebAnon 
         Caption         =   "Anonymous &jump to URL"
         Shortcut        =   ^J
      End
   End
End
Attribute VB_Name = "frmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
Dim rs As Recordset
    'On Error GoTo BadEntry
    'Set rs = DB.OpenRecordset("Contacts", dbOpenDynaset)
    'rs.AddNew
    'rs("Contact Name") = txtName(0)
    'If InStr(1, txtName(1), "@") = 0 Then
       ' MsgBox "This is not a valid email address.", vbApplicationModal + vbCritical
       ' Exit Sub
   ' End If
   ' rs("Contact Email") = txtName(1) '"<" & txtName(1) & ">"
   ' rs.Update
    'Unload Me

   UpdateContactRecord (txtName(0) & " " & txtName(1))
Unload Me
Exit Sub
BadEntry:
    MsgBox "New contact was not entered. Reason was: " & Err.Description, vbApplicationModal + vbCritical
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
    
Win.Center Me, Null
Win.OnTop(Me) = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAddressBook = Nothing
End Sub

