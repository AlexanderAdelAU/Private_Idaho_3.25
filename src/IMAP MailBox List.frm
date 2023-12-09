VERSION 5.00
Object = "{33337313-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "imap40.ocx"
Begin VB.Form frmIMAPMailBoxList 
   Caption         =   "IMAP MailBox List"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   390
      TabIndex        =   1
      Top             =   1020
      Width           =   7725
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   435
      Left            =   6660
      TabIndex        =   0
      Top             =   4110
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Mailbox , Separator  , Flags"
      Height          =   435
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This only list your mailboxes on your IMAP server.  Future releases of PI will incorporate more features here."
      Height          =   435
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   90
      Width           =   6495
   End
   Begin IMAPLibCtl.IMAP IMAP2 
      Left            =   7740
      Top             =   270
      LocalFile       =   ""
      Mailbox         =   "Inbox"
      MailServer      =   ""
      MessageSet      =   ""
      Password        =   ""
      SearchCriteria  =   ""
      User            =   ""
      WinsockLoaded   =   -1  'True
   End
End
Attribute VB_Name = "frmIMAPMailBoxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_Click()

End Sub

Private Sub btnok_Click()
Unload Me
End Sub

Private Sub Form_Load()
Exit Sub
'frmMain.ConnectIMAP4
IMAP2.MailServer = "mail.itech.net.au"

IMAP2.Mailbox = """*"""
IMAP2.Action = 15 'a_ListMailboxes
IMAP2.Mailbox = """INBOX"""
IMAP2.Action = 4 'a_SelectMailbox
If Not IMAP2.MessageCount = 0 Then
    IMAP2.MessageSet = "1:" & IMAP2.MessageCount
    'IMAP2.MessageSet = IMAP2.MessageId
    IMAP2.Action = 20 'a_GetMessageHeaders
    Text1 = IMAP2.MessageHeaders
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmIMAPMailBoxList = Nothing
End Sub
