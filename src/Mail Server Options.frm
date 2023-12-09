VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmMailServerOptions 
   Caption         =   "Mail Server Options"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Mail Server"
      TabPicture(0)   =   "Mail Server Options.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdMailServerApply"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtReplyTo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDNSServerName"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "MailServerOptionHelp(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MailServerOptionHelp(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Send and Receive Options"
      TabPicture(1)   =   "Mail Server Options.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblHelp"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "btnRetrievalOptionsApply"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Create Mail List"
      TabPicture(2)   =   "Mail Server Options.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstMailList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtList"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnAdd"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "btnRemove"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnAddMembers"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton MailServerOptionHelp 
         Caption         =   "<--- Help"
         Height          =   315
         Index           =   0
         Left            =   6600
         TabIndex        =   50
         Top             =   1620
         Width           =   1035
      End
      Begin VB.CommandButton MailServerOptionHelp 
         Caption         =   "<--- Help"
         Height          =   315
         Index           =   1
         Left            =   6600
         TabIndex        =   49
         Top             =   1980
         Width           =   1035
      End
      Begin VB.TextBox txtDNSServerName 
         Height          =   285
         Left            =   2340
         TabIndex        =   46
         Top             =   1980
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2340
         TabIndex        =   45
         Top             =   1620
         Width           =   4095
      End
      Begin VB.CommandButton btnAddMembers 
         Caption         =   "Add/Remove Members"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -73560
         TabIndex        =   42
         Top             =   5100
         Width           =   2055
      End
      Begin VB.CommandButton btnRemove 
         Caption         =   "Remove Mail Group"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -70320
         TabIndex        =   41
         Top             =   5100
         Width           =   1935
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add Group/List"
         Height          =   375
         Left            =   -68160
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtList 
         Height          =   375
         Left            =   -73560
         TabIndex        =   39
         Top             =   960
         Width           =   5175
      End
      Begin VB.ListBox lstMailList 
         Height          =   3375
         Left            =   -73560
         TabIndex        =   38
         ToolTipText     =   "Right click to modify"
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txtReplyTo 
         Height          =   285
         Left            =   2340
         TabIndex        =   35
         Top             =   1260
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Incoming and Outgoing Server Details"
         Height          =   2295
         Left            =   210
         TabIndex        =   21
         Top             =   2580
         Width           =   8115
         Begin VB.CheckBox chkAuthenticationRequired 
            Caption         =   "Your Server Requires Authentication"
            Height          =   375
            Left            =   6060
            TabIndex        =   55
            Top             =   1620
            Width           =   1875
         End
         Begin VB.TextBox txtSMTPPort 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   54
            Text            =   "25"
            Top             =   735
            Width           =   735
         End
         Begin VB.CommandButton MailServerOptionHelp 
            Caption         =   "<--- Help"
            Height          =   315
            Index           =   3
            Left            =   6600
            TabIndex        =   52
            Top             =   1110
            Width           =   975
         End
         Begin VB.CommandButton MailServerOptionHelp 
            Caption         =   "<--- Help"
            Height          =   315
            Index           =   2
            Left            =   6600
            TabIndex        =   51
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkDefaultPort 
            Caption         =   "Use Default Port"
            Height          =   315
            Left            =   6030
            TabIndex        =   34
            Top             =   330
            Width           =   1545
         End
         Begin VB.TextBox txtMailPort 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   32
            Text            =   "143"
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox cmbConnectOption 
            Height          =   315
            ItemData        =   "Mail Server Options.frx":0054
            Left            =   2160
            List            =   "Mail Server Options.frx":005E
            TabIndex        =   30
            Top             =   330
            Width           =   1485
         End
         Begin VB.TextBox txtSMTPServerName 
            Height          =   285
            Left            =   2190
            TabIndex        =   28
            Tag             =   "Note: Use this if you want to send mail through a different server."
            Top             =   750
            Width           =   2895
         End
         Begin VB.TextBox Text8 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2190
            PasswordChar    =   "*"
            TabIndex        =   24
            Top             =   1830
            Width           =   2895
         End
         Begin VB.TextBox txtMailServerName 
            Height          =   285
            Left            =   2190
            TabIndex        =   23
            Top             =   1110
            Width           =   2895
         End
         Begin VB.TextBox Text5 
            Height          =   300
            Left            =   2190
            TabIndex        =   22
            Top             =   1470
            Width           =   2895
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Port to use: "
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   4080
            TabIndex        =   33
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Connect Using"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   31
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP server name"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   150
            TabIndex        =   29
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Your Account Password"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   27
            Top             =   1890
            Width           =   1755
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Incoming Server Name"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1170
            Width           =   1695
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Your Account Name"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   25
            Top             =   1530
            Width           =   1515
         End
      End
      Begin VB.CommandButton btnRetrievalOptionsApply 
         Caption         =   "Apply"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68250
         TabIndex        =   17
         Top             =   5130
         Width           =   1335
      End
      Begin VB.CommandButton cmdMailServerApply 
         Caption         =   "Apply"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6810
         TabIndex        =   16
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2340
         TabIndex        =   11
         Top             =   540
         Width           =   4095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2340
         TabIndex        =   10
         Top             =   900
         Width           =   4095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mail Server Scan Interval"
         Height          =   1545
         Left            =   -74040
         TabIndex        =   5
         Top             =   4020
         Width           =   4995
         Begin VB.TextBox txtScan 
            Height          =   285
            Left            =   1530
            TabIndex        =   6
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Scan every:"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "minutes."
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   8
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "The server will be scanned for either PGP or normal messages depending on the time you enter.  Zero (0) means don't scan."
            Height          =   555
            Index           =   2
            Left            =   150
            TabIndex        =   7
            Top             =   300
            Width           =   4335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mail Server Options"
         Height          =   2775
         Left            =   -74040
         TabIndex        =   2
         Top             =   660
         Width           =   5505
         Begin VB.CheckBox chkOption 
            Caption         =   "Bypass your ISP - Send Messages direct to Destination Server"
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   3
            Left            =   330
            TabIndex        =   47
            Top             =   2160
            Width           =   4995
         End
         Begin VB.CheckBox chkOption 
            Caption         =   "Block messages from List"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   330
            TabIndex        =   43
            Top             =   1890
            Width           =   3195
         End
         Begin VB.OptionButton optEmailScan 
            Caption         =   "Don't Retrieve Messages, just send"
            Height          =   285
            Index           =   2
            Left            =   330
            TabIndex        =   20
            Top             =   810
            Width           =   3105
         End
         Begin VB.OptionButton optEmailScan 
            Caption         =   "Retrieve all Messages"
            Height          =   285
            Index           =   1
            Left            =   330
            TabIndex        =   19
            Top             =   540
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.OptionButton optEmailScan 
            Caption         =   "Retrieve PGP Messages only"
            Height          =   285
            Index           =   0
            Left            =   330
            TabIndex        =   18
            Top             =   270
            Width           =   2865
         End
         Begin VB.CheckBox chkOption 
            Caption         =   "Leave Messages on Server"
            Height          =   315
            Index           =   0
            Left            =   330
            TabIndex        =   4
            Top             =   1260
            Width           =   3195
         End
         Begin VB.CheckBox chkOption 
            Caption         =   "Preview Messages before Downloading"
            Height          =   315
            Index           =   1
            Left            =   330
            TabIndex        =   3
            Top             =   1590
            Width           =   3195
         End
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   53
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblHelp 
         Caption         =   "lblHelp"
         ForeColor       =   &H00008000&
         Height          =   3675
         Left            =   -68400
         TabIndex        =   48
         Top             =   780
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DNS Server Address"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   44
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Can be blank"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   37
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reply To"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   1290
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail address"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"Mail Server Options.frx":006F
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   210
         TabIndex        =   13
         Top             =   4920
         Width           =   7905
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NNTP (News server name)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1620
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   315
      Left            =   7170
      TabIndex        =   0
      Top             =   6420
      Width           =   1245
   End
   Begin VB.Menu mnuMailGroupPopUpMenu 
      Caption         =   "MailGroupPopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteMailGroup 
         Caption         =   "Delete Mail Group"
      End
      Begin VB.Menu mnuRenameMailGroup 
         Caption         =   "Rename Mail Group"
      End
   End
End
Attribute VB_Name = "frmMailServerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pOptionsTab As Integer
'Private Loading As Boolean
Private m_EmailScanInterval As Integer
Private m_TimerInterval As Integer
Private m_RetrievePGPMessages As Boolean
Private m_RetrievePlainMessages As Boolean
Private m_RetrieveAllMessages As Boolean
Private m_RetrieveNoMessages As Boolean


Private Sub btnAdd_Click()
Dim sReponse As String

sReponse = CreateListTable(txtList)
If Not sReponse = "" Then
    MsgBox "Mail Group list table could not be created.  Make sure you have used valid characters in the name etc.  Actual error returned is: " & sReponse, vbCritical + vbApplicationModal, "Bad List Name"
    Exit Sub
End If
LoadMailGroupList lstMailList
End Sub

Private Sub btnAddMembers_Click()
If Not lstMailList.List(lstMailList.ListIndex) = "" Then
    frmAddListMembers.lblListName = lstMailList.List(lstMailList.ListIndex)
    frmAddListMembers.Show vbModal
End If
btnAddMembers.Enabled = False
End Sub

Private Sub btnRemove_Click()
If Not lstMailList.ListIndex = -1 Then
    DeleteListTable lstMailList.List(lstMailList.ListIndex)
    LoadMailGroupList lstMailList
End If
btnRemove.Enabled = False
End Sub

Private Sub btnRetrievalOptionsApply_Click()
Dim SectionName As String
Dim KeyName As String
Dim KeyValue As String

    
    m_TimerInterval = CInt(1000 * 30)
    m_EmailScanInterval = Val(txtScan) * 2
    
    SectionName = "Options"
    KeyName = "ServerDelete"
    
    'Check if we delete from the server
    If chkOption(0).Value = vbChecked Then
        KeyValue = "False"
    Else
        KeyValue = "True"
    End If
    WriteProfile SectionName, KeyName, KeyValue
    
    SectionName = "Options"
    KeyName = "ServerPreviewMessages"
    If chkOption(1).Value = vbChecked Then
        KeyValue = "True"
    Else
        KeyValue = "False"
    End If
    WriteProfile SectionName, KeyName, KeyValue
    
    SectionName = "Options"
    KeyName = "DeliverSMTPMessagesDirect"
    If chkOption(3).Value = vbChecked Then
        KeyValue = "True"
    Else
        KeyValue = "False"
    End If
    WriteProfile SectionName, KeyName, KeyValue
    
    
    m_RetrieveAllMessages = False
    m_RetrieveNoMessages = False
    m_RetrievePGPMessages = False

    If optEmailScan(SCAN_PGP_ONLY).Value = True Then m_RetrievePGPMessages = True
    If optEmailScan(SCAN_ALL).Value = True Then m_RetrieveAllMessages = True
    If optEmailScan(SCAN_NONE).Value = True Then m_RetrieveNoMessages = True
    
    btnRetrievalOptionsApply.Enabled = False
    
End Sub

Private Sub Check1_Click()

End Sub



Private Sub chkDefaultPort_Click()
If cmbConnectOption.ListIndex = CONNECT_POP3 Then
    txtMailPort = 110
Else
    txtMailPort = 143
End If
End Sub

Private Sub chkOption_Click(Index As Integer)
Dim SHelpString As String
btnRetrievalOptionsApply.Enabled = True
Select Case Index
    Case 0
        lblHelp.Visible = True
        SHelpString = "Note:  If you choose this option your messages will not be erased from the server.  Good for viewing the messages and then using another e-mail client to download again later etc."
        lblHelp.Caption = SHelpString
    Case 1
        lblHelp.Visible = True
        SHelpString = "Note:  If you choose this option you can view each message before downloading.  This is useful for e-mails with large attachements or for SPAM e-mail."
        lblHelp.Caption = SHelpString
    Case 2
        
    Case 3
        lblHelp.Visible = True
        SHelpString = "Note:  If you choose this option you may need to fill in the address or name of your local DNS server on the 'Mail Server' Tab." & vbCrLf & vbCrLf
        SHelpString = SHelpString & "Also, if you are sending an e-mail with a large number of recipients, then this may be slow as each unique e-mail address will require a separate connection to the destination mail server."
        lblHelp.Caption = SHelpString
        
End Select
End Sub



Private Sub cmbConnectOption_Click()
Select Case cmbConnectOption.ListIndex
    Case CONNECT_POP3
        txtMailPort = 110
    Case CONNECT_IMAP4
        txtMailPort = 143
        End Select

End Sub

Private Sub cmdMailServerApply_Click()
Dim SectionName As String
    Dim KeyName As String
    Dim KeyValue As String
    Dim tmpstr As String
    '---------------------------------------------
    'get data and save options to init file
    '---------------------------------------------
    SectionName = "Options"
        
    MailConnector.ConnectUsing = cmbConnectOption.ListIndex
    If MailConnector.ConnectUsing = CONNECT_POP3 Then
        KeyValue = "CONNECT_POP3"
    Else
        KeyValue = "CONNECT_IMAP4"
    End If
    KeyName = "ConnectUsing"
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.POPPort = txtMailPort
    KeyName = "MailPort"
    KeyValue = MailConnector.POPPort
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.NNTPServerName = Text1.Text
    KeyName = "NNTPServerName"
    KeyValue = MailConnector.NNTPServerName
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.SMTPPort = txtSMTPPort.Text
    KeyName = "SMTPPort"
    KeyValue = MailConnector.SMTPPort
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.DNSServerName = txtDNSServerName.Text
    KeyName = "DNSServerName"
    KeyValue = MailConnector.DNSServerName
    WriteProfile SectionName, KeyName, KeyValue
        
    MailConnector.EmailAddress = Text2.Text
    KeyName = "EmailAddress"
    KeyValue = MailConnector.EmailAddress
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.ReplyEmailAddress = txtReplyTo
    KeyName = "ReplyEmailAddress"
    KeyValue = MailConnector.ReplyEmailAddress
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.RealName = Text3.Text
    KeyName = "RealName"
    KeyValue = MailConnector.RealName
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.SMTPServerName = txtSMTPServerName.Text
    KeyName = "SMTPServerName"
    KeyValue = MailConnector.SMTPServerName
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.AccountName = Text5.Text
    KeyName = "AccountName"
    KeyValue = MailConnector.AccountName
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.AuthenticationRequired = chkAuthenticationRequired.Value
    KeyName = "AuthenticationRequired"
    KeyValue = MailConnector.AuthenticationRequired
    WriteProfile SectionName, KeyName, KeyValue
    
    MailConnector.MailServerName = txtMailServerName
    KeyName = "MailServerName"
    KeyValue = MailConnector.MailServerName
    WriteProfile SectionName, KeyName, KeyValue
   
    MailConnector.AccountPassword = Text8.Text
    KeyName = "AccountPassword"
   ' KeyValue = Text8.Text
    SavePasswordToDatabase Text8.Text
    WriteProfile SectionName, KeyName, "***********"  'KeyValue
    
    cmdMailServerApply.Enabled = False
End Sub

Private Sub Command1_Click()
'Note this has to be me.hide as the variables are passed to frmMain.
Me.Hide
End Sub

Private Sub Command2_Click()

End Sub


Private Sub DNSHelpButton_Click()

End Sub

Private Sub Form_Activate()
If m_RetrievePGPMessages Then optEmailScan(0).Enabled = True
If m_RetrieveAllMessages Then optEmailScan(1).Enabled = True
If m_RetrieveNoMessages Then optEmailScan(2).Enabled = True
'Call SSTab1_Click(pOptionsTab)
SSTab1.Tab = pOptionsTab
Call SSTab1_Click(pOptionsTab)
End Sub

Private Sub Form_Load()
'SSTab1.Tab = 0
'Call SSTab1_Click(pOptionsTab)
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Public Property Get ScanInterval() As Integer
ScanInterval = m_EmailScanInterval
End Property
Public Property Get TimerSetting() As Integer
TimerSetting = m_TimerInterval
End Property
Public Property Let ScanInterval(ByVal ScanInterval As Integer)
 m_EmailScanInterval = ScanInterval
End Property
Public Property Let TimerSetting(ByVal TimerSetting As Integer)
 m_TimerInterval = TimerSetting
End Property

Private Sub Form_Unload(Cancel As Integer)
Set frmMailServerOptions = Nothing
End Sub
Public Property Get EmailScanOption() As Integer
If m_RetrievePGPMessages Then EmailScanOption = SCAN_PGP_ONLY
If m_RetrieveAllMessages Then EmailScanOption = SCAN_ALL
If m_RetrieveNoMessages Then EmailScanOption = SCAN_NONE

End Property

Public Property Let EmailScanOption(ByVal ScanOption As Integer)
 m_RetrievePGPMessages = False
 m_RetrieveAllMessages = False
 m_RetrieveNoMessages = False
 
Select Case ScanOption
        Case 0
            optEmailScan(SCAN_PGP_ONLY).Value = True
            m_RetrievePGPMessages = True
        Case 1
            optEmailScan(SCAN_ALL).Value = True
            m_RetrieveAllMessages = True
        Case 2
            optEmailScan(SCAN_NONE).Value = True
            m_RetrieveNoMessages = True
    End Select
End Property

Private Sub lstMailList_Click()
btnRemove.Enabled = True
btnAddMembers.Enabled = True
End Sub

Private Sub lstMailList_DblClick()
Call btnAddMembers_Click
End Sub

Private Sub lstMailList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As String
 If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      PopupMenu mnuMailGroupPopUpMenu   ' Display the File menu as a
                        ' pop-up menu.
End If


End Sub


Private Sub MailServerOptionHelp_Click(Index As Integer)
Dim SHelpString As String
Select Case Index
    Case 1
        SHelpString = "May be blank.  Use this field when you want to send your messages/e-mail direct to the destination server." & vbCrLf & vbCrLf
        SHelpString = SHelpString & "Your DNS Server may not be required if Private Idaho can find a local DNS Server on your site.  If PI can't find a DNS Server it will use the value in this field." & vbCrLf & vbCrLf
        SHelpString = SHelpString & "The DNS Server is needed so that the IP address of the recipient's mail server can be determined.  The value can be either entered as an address, e.g. dns.yourdomain.com or as an IP address, e.g. 12.12.123.32."
        MsgBox SHelpString, vbApplicationModal, "Using the DNS Field"
    Case 0
        SHelpString = "Can be Blank.  " & vbCrLf & vbCrLf
        SHelpString = "Use this field when you want to post to NewsGroups etc.  This is normally provided by your ISP or some onther provider.  It can be an address like newsserver.somewhere.com, or as an IP address, e.g. 12.12.123.32."
        MsgBox SHelpString, vbApplicationModal, "Using the NNTP Field"
        
    Case 2
        SHelpString = "This field holds the address of the mail server that you connect to to send e-mails, if different from the 'Mail Server Name' below. "
        SHelpString = SHelpString & vbCrLf & vbCrLf & "However this field can can be blank, "
        SHelpString = SHelpString & "if you send e-mails direct to the destination server - then you don't need to fill in this field."
        MsgBox SHelpString, vbApplicationModal, "Using the SMTP Field"
    Case 3
         SHelpString = "This can be used for sending and receiving e-mails. However it can be blank if the following conditions are met:  " & vbCrLf & vbCrLf
        SHelpString = SHelpString & "1. You use the option to send e-mails directly to the destination server, and " & vbCrLf
        SHelpString = SHelpString & "2. You don't use POP3 etc to receive your e-mails."
        MsgBox SHelpString, vbApplicationModal, "Using the MailServer Field"
    
End Select

End Sub

Private Sub mnuDeleteMailGroup_Click()
If Not lstMailList.ListIndex = -1 Then
    DeleteListTable lstMailList.List(lstMailList.ListIndex)
    LoadMailGroupList lstMailList
Else
    MsgBox "You must select a group from the listed", vbApplicationModal + vbExclamation
End If
End Sub

Private Sub mnuRenameMailGroup_Click()
If Not lstMailList.ListIndex = -1 Then
    RenameListTable lstMailList.List(lstMailList.ListIndex)
    LoadMailGroupList lstMailList
Else
    MsgBox "You must select a group from the listed", vbApplicationModal + vbExclamation
End If
End Sub

Private Sub optEmailScan_Click(Index As Integer)
btnRetrievalOptionsApply.Enabled = True
Exit Sub
'If Value = True Then
    m_RetrievePGPMessages = False
    m_RetrieveAllMessages = False
     m_RetrieveNoMessages = False
    Select Case Index
        Case 0
            m_RetrievePGPMessages = True
            optEmailScan(0).Enabled = True
            'chkOption(1).Enabled = True
        Case 1
            m_RetrieveAllMessages = True
            optEmailScan(1).Enabled = True
            'chkOption(1).Enabled = True
          '  m_RetrievePlainMessages = True
        Case 2
            m_RetrieveNoMessages = True
            optEmailScan(2).Enabled = False
            'chkOption(1).Enabled = False
            
    End Select
'End If

btnRetrievalOptionsApply.Enabled = True
End Sub

Private Sub optEmailScan_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'btnRetrievalOptionsApply.Enabled = True
End Sub

Private Sub optEmailScan1_Click(Index As Integer, Value As Integer)
Exit Sub
If Value = True Then
    m_RetrievePGPMessages = False
    m_RetrieveAllMessages = False
     m_RetrieveNoMessages = False
    Select Case Index
        Case 0
            m_RetrievePGPMessages = True
            optEmailScan(0).Enabled = True
            'chkOption(1).Enabled = True
        Case 1
            m_RetrieveAllMessages = True
            optEmailScan(1).Enabled = True
            'chkOption(1).Enabled = True
          '  m_RetrievePlainMessages = True
        Case 2
            m_RetrieveNoMessages = True
            optEmailScan(2).Enabled = False
            'chkOption(1).Enabled = False
            
    End Select
End If

btnRetrievalOptionsApply.Enabled = True
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim i As Integer
Dim rs As Recordset
DoEvents
    
Select Case SSTab1.Tab

    Case 0
    '---------------------------------------------
    'fill text boxes
    '---------------------------------------------
    If MailConnector.ConnectUsing = CONNECT_POP3 Then
        cmbConnectOption.ListIndex = CONNECT_POP3
    Else
        cmbConnectOption.ListIndex = CONNECT_IMAP4
    End If
    Text1.Text = MailConnector.NNTPServerName
    txtDNSServerName.Text = MailConnector.DNSServerName
    Text2.Text = MailConnector.EmailAddress
    Text3.Text = MailConnector.RealName
    txtSMTPServerName = MailConnector.SMTPServerName
    Text5.Text = MailConnector.AccountName
    txtMailServerName = MailConnector.MailServerName
    If MailConnector.ReplyEmailAddress = "" Then
        MailConnector.ReplyEmailAddress = MailConnector.EmailAddress
    End If
    txtReplyTo = MailConnector.ReplyEmailAddress
    If MailConnector.POPPort = 0 Then
        chkDefaultPort.Value = vbChecked
    Else
        txtMailPort = MailConnector.POPPort
    End If
    Text8.Text = MailConnector.AccountPassword
    txtSMTPPort = IIf(MailConnector.SMTPPort = 0, 25, MailConnector.SMTPPort)
    chkAuthenticationRequired = IIf(MailConnector.AuthenticationRequired, 1, 0)
    Case 1
        Dim SectionName As String
        Dim sRes As String
        txtScan = CInt(m_EmailScanInterval / 2)
        SectionName = "Options"
        sRes = ReadProfile(SectionName, "ServerDelete")
        If sRes = "True" Then
            chkOption(0).Value = vbUnchecked
        Else
            chkOption(0).Value = vbChecked
        End If

        sRes = ReadProfile(SectionName, "ServerPreviewMessages")
        If sRes = "True" Then
            chkOption(1).Value = vbChecked
        Else
            chkOption(1).Value = vbUnchecked
        End If
        
        sRes = ReadProfile(SectionName, "DeliverSMTPMessagesDirect")
        If sRes = "True" Then
            chkOption(3).Value = vbChecked
        Else
            chkOption(3).Value = vbUnchecked
        End If
        
        
        If m_RetrievePGPMessages Then optEmailScan(0).Enabled = True
        If m_RetrieveAllMessages Then optEmailScan(1).Enabled = True
        If m_RetrieveNoMessages Then optEmailScan(2).Enabled = True
   Case 2
        LoadMailGroupList lstMailList
    End Select
        
DoEvents
End Sub

Private Sub Text1_Change()
btnRetrievalOptionsApply.Enabled = True
End Sub

Private Sub Text2_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub Text3_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub Text4_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub Text5_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub Text6_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub Text8_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub txtDNSServerName_Change()
cmdMailServerApply.Enabled = True
End Sub

Private Sub txtMailPort_KeyDown(KeyCode As Integer, Shift As Integer)
cmdMailServerApply.Enabled = True
End Sub

Private Sub txtScan_Change()
btnRetrievalOptionsApply.Enabled = True
End Sub


