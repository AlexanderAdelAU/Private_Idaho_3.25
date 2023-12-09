VERSION 5.00
Object = "{33333233-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "SMTP32.OCX"
Object = "{33333183-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "HTTP32.OCX"
Object = "{33333143-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "NETCOD32.OCX"
Object = "{33333243-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "POP32.OCX"
Object = "{33333253-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "NNTP32.OCX"
Object = "{33333113-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "IPPORT32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{33333153-F789-11CE-86F8-0020AFD8C6DB}#2.0#0"; "IPINFO32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Private Idaho 2.8c"
   ClientHeight    =   6060
   ClientLeft      =   405
   ClientTop       =   825
   ClientWidth     =   8835
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
   Icon            =   "EUPGP.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   419.668
   ScaleMode       =   0  'User
   ScaleWidth      =   2404.082
   Begin POPLib.POP POP1 
      Left            =   5580
      OleObjectBlob   =   "EUPGP.frx":030A
      Top             =   1305
   End
   Begin SMTPLib.SMTP SMTP1 
      Left            =   5085
      OleObjectBlob   =   "EUPGP.frx":0336
      Top             =   1305
   End
   Begin HTTPLib.HTTP HTTP1 
      Left            =   4140
      OleObjectBlob   =   "EUPGP.frx":0368
      Top             =   1305
   End
   Begin NetCodeLib.NetCode NetCode1 
      Left            =   4635
      OleObjectBlob   =   "EUPGP.frx":03BC
      Top             =   1305
   End
   Begin NNTPLib.NNTP NNTP1 
      Left            =   3630
      OleObjectBlob   =   "EUPGP.frx":03EC
      Top             =   1305
   End
   Begin IPPortLib.IPPort IPPort1 
      Left            =   3870
      OleObjectBlob   =   "EUPGP.frx":0418
      Top             =   2625
   End
   Begin IPInfoLib.IPInfo IPInfo1 
      Left            =   4980
      OleObjectBlob   =   "EUPGP.frx":0459
      Top             =   2655
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
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
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1650
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   2250
      TabIndex        =   6
      Top             =   1545
      Width           =   6300
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Attachment:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   225
      TabIndex        =   4
      Top             =   1335
      Width           =   1650
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   1305
      TabIndex        =   15
      Top             =   2430
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2055
      Width           =   885
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   225
      TabIndex        =   20
      Top             =   2430
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1665
      TabIndex        =   0
      Top             =   30
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2025
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1665
      TabIndex        =   3
      Top             =   990
      Width           =   6855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1665
      TabIndex        =   2
      Top             =   675
      Width           =   6855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1665
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3045
      Width           =   8370
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4365
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DialogTitle     =   "CommonDialog1"
      FontName        =   "Arial"
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   5355
      TabIndex        =   24
      Top             =   2835
      Width           =   3210
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "To:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   225
      MousePointer    =   5  'Size
      TabIndex        =   23
      Top             =   90
      Width           =   315
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "mod"
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
      Height          =   300
      Left            =   5355
      TabIndex        =   22
      Top             =   2565
      Width           =   3210
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   2265
      TabIndex        =   14
      Top             =   2535
      Width           =   3135
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "up time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   7065
      TabIndex        =   19
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "latency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5805
      TabIndex        =   18
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remailer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2250
      TabIndex        =   17
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Message:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   210
      TabIndex        =   13
      Top             =   2835
      Width           =   885
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "BCC:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   225
      TabIndex        =   12
      Top             =   990
      Width           =   525
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CC:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   225
      TabIndex        =   11
      Top             =   675
      Width           =   435
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subject:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   225
      TabIndex        =   10
      Top             =   375
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click Here"
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   540
      MousePointer    =   5  'Size
      TabIndex        =   9
      Top             =   90
      Width           =   1170
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD NYM PASSWORD AND ENCRYPT"
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
      Height          =   255
      Left            =   3915
      MousePointer    =   5  'Size
      TabIndex        =   16
      Top             =   1350
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu FileImport 
         Caption         =   "&Import message..."
      End
      Begin VB.Menu FileExport 
         Caption         =   "Sa&ve message..."
      End
      Begin VB.Menu FilePrintM 
         Caption         =   "Print message"
      End
      Begin VB.Menu FilePage 
         Caption         =   "Page setup..."
         Visible         =   0   'False
      End
      Begin VB.Menu FileNull3 
         Caption         =   "-"
      End
      Begin VB.Menu FileAddress 
         Caption         =   "&Address book..."
      End
      Begin VB.Menu FileOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu FileMode 
         Caption         =   "&Expert"
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save settings"
      End
      Begin VB.Menu FileNull 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu FileConnect 
         Caption         =   "&Connect"
         Visible         =   0   'False
      End
      Begin VB.Menu FileDisconnect 
         Caption         =   "&Disconnect"
         Visible         =   0   'False
      End
      Begin VB.Menu FileNull2 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu EditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu EditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu EditSep 
         Caption         =   "-"
      End
      Begin VB.Menu EditClrAll 
         Caption         =   "Clea&r all"
         Shortcut        =   ^R
      End
      Begin VB.Menu EditClrMsg 
         Caption         =   "C&lear message"
         Shortcut        =   ^L
      End
      Begin VB.Menu EditCopyMsg 
         Caption         =   "C&opy message"
         Shortcut        =   ^O
      End
      Begin VB.Menu EditPasteMsg 
         Caption         =   "Paste &message"
         Shortcut        =   ^M
      End
      Begin VB.Menu EditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditSetFont 
         Caption         =   "&Font"
         Shortcut        =   ^F
      End
      Begin VB.Menu EditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EditHeader 
         Caption         =   "Insert &header"
         Shortcut        =   ^H
      End
      Begin VB.Menu EditSig 
         Caption         =   "Insert s&ignature"
         Shortcut        =   ^I
      End
      Begin VB.Menu EditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EditOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu PGPMenu 
      Caption         =   "&PGP"
      Begin VB.Menu PGPEncrypt 
         Caption         =   "&Encrypt message"
         Shortcut        =   ^E
      End
      Begin VB.Menu PGPEnSign 
         Caption         =   "Encrypt and &sign message"
      End
      Begin VB.Menu PGPClearSign 
         Caption         =   "&Clear sign message"
         Shortcut        =   ^G
      End
      Begin VB.Menu PGPDecrypt 
         Caption         =   "&Decrypt or verify message"
         Shortcut        =   ^D
      End
      Begin VB.Menu PGPSep 
         Caption         =   "-"
      End
      Begin VB.Menu PGPMultiple 
         Caption         =   "Use multiple keys"
      End
      Begin VB.Menu PGPSelf 
         Caption         =   "Encrypt to self"
      End
      Begin VB.Menu PGPEyes 
         Caption         =   "Eyes only"
      End
      Begin VB.Menu PGPConvent 
         Caption         =   "Conventional encrypt"
      End
      Begin VB.Menu PGPObscurity 
         Caption         =   "Obscurity"
      End
      Begin VB.Menu PGPWrap 
         Caption         =   "Word wrap on encrypt/sign"
      End
      Begin VB.Menu PGPSepx 
         Caption         =   "-"
      End
      Begin VB.Menu PGPFile 
         Caption         =   "File operations"
      End
      Begin VB.Menu PGPSep1 
         Caption         =   "-"
      End
      Begin VB.Menu PGPMin 
         Caption         =   "Run PGP minimized"
      End
      Begin VB.Menu PGPClose 
         Caption         =   "Close PGP window on exit"
      End
      Begin VB.Menu PGPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu PGPVerify 
         Caption         =   "&Verify PGP Distribution"
      End
      Begin VB.Menu PGPSep3 
         Caption         =   "-"
      End
      Begin VB.Menu PGPOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu KeyMenu 
      Caption         =   "&Keys"
      Begin VB.Menu KeyCreate 
         Caption         =   "&Create key pair"
      End
      Begin VB.Menu KeyCertify 
         Caption         =   "&Sign a key"
      End
      Begin VB.Menu KeyEditTrust 
         Caption         =   "Change level of &trust in a key"
      End
      Begin VB.Menu keySep 
         Caption         =   "-"
      End
      Begin VB.Menu PGPDeleteKey 
         Caption         =   "&Delete key..."
      End
      Begin VB.Menu PGPAddKey 
         Caption         =   "&Add key from message"
      End
      Begin VB.Menu KeySep1 
         Caption         =   "-"
      End
      Begin VB.Menu PGPInsertKey 
         Caption         =   "&Insert key in message..."
      End
      Begin VB.Menu PGPAdd 
         Caption         =   "&Update PUBKEYS.OUT"
      End
      Begin VB.Menu keySep3 
         Caption         =   "-"
      End
      Begin VB.Menu KeySubmit 
         Caption         =   "Submit &key to server"
      End
      Begin VB.Menu PGPGetKey 
         Caption         =   "&Get key from server"
      End
      Begin VB.Menu keySep4 
         Caption         =   "-"
      End
      Begin VB.Menu KeyOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu RemailerMenu 
      Caption         =   "&Remailers"
      Begin VB.Menu RemailersCP 
         Caption         =   "Cypherpunk"
         Checked         =   -1  'True
      End
      Begin VB.Menu RemailersMix 
         Caption         =   "Mixmaster"
      End
      Begin VB.Menu RemailersSoda 
         Caption         =   "Eric"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu RemailersPenet 
         Caption         =   "Anon.penet.fi"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu RemailersNull 
         Caption         =   "-"
      End
      Begin VB.Menu OptionsAdv 
         Caption         =   "Advanced C&ypherpunk"
         Begin VB.Menu RemailersPGP 
            Caption         =   "Encrypt to remailer(s)"
         End
         Begin VB.Menu RemailersHash 
            Caption         =   "Replace subject"
         End
         Begin VB.Menu RemailersLatent 
            Caption         =   "Use latent time"
         End
         Begin VB.Menu RemailersCut 
            Caption         =   "Use cutmarks"
         End
         Begin VB.Menu RemailersPost 
            Caption         =   "Post to USENET"
         End
      End
      Begin VB.Menu RemailersNull2 
         Caption         =   "-"
      End
      Begin VB.Menu RemailerUpdate 
         Caption         =   "&Update remailer info"
         Shortcut        =   ^U
      End
      Begin VB.Menu RemailerKeys 
         Caption         =   "&Get remailer keys"
      End
   End
   Begin VB.Menu UseNetMenu 
      Caption         =   "&Newsgroups"
      Begin VB.Menu USENETNone 
         Caption         =   "None"
         Checked         =   -1  'True
      End
      Begin VB.Menu USENETGate 
         Caption         =   "E-mail gateway"
      End
      Begin VB.Menu UseNetSoda 
         Caption         =   "Eric"
         Visible         =   0   'False
      End
      Begin VB.Menu USENETFi 
         Caption         =   "Anon.penet.fi"
         Visible         =   0   'False
      End
      Begin VB.Menu USENETheader 
         Caption         =   "Remailer newsgroup header"
      End
      Begin VB.Menu NewsNull1 
         Caption         =   "-"
      End
      Begin VB.Menu GetNews 
         Caption         =   "&Newsgroup manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu NewsNull2 
         Caption         =   "-"
      End
      Begin VB.Menu GetFinger 
         Caption         =   "&Finger operations"
         Enabled         =   0   'False
      End
      Begin VB.Menu PI_Test_Click 
         Caption         =   "Special Test"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu DoConnect 
         Caption         =   "Connect"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu TransferMenu 
      Caption         =   "&Message"
      Begin VB.Menu TransferSend 
         Caption         =   "&Send"
         Shortcut        =   ^S
      End
      Begin VB.Menu TransferReply 
         Caption         =   "&Insert reply markers"
      End
      Begin VB.Menu ReamilerAppend 
         Caption         =   "&Append info"
      End
      Begin VB.Menu TransferAES 
         Caption         =   "A&ppend on send"
      End
      Begin VB.Menu TransferXHeaders 
         Caption         =   "&X-Headers..."
      End
      Begin VB.Menu MsgSep 
         Caption         =   "-"
      End
      Begin VB.Menu TransferEu 
         Caption         =   "Trans&fer to app"
         Shortcut        =   ^T
      End
      Begin VB.Menu TransferApps 
         Caption         =   "Transf&er to"
         Visible         =   0   'False
         Begin VB.Menu TransferApp1 
            Caption         =   "Application 1"
         End
         Begin VB.Menu TransferApp2 
            Caption         =   "Application 2"
            Visible         =   0   'False
         End
         Begin VB.Menu TransferApp3 
            Caption         =   "Application 3"
            Visible         =   0   'False
         End
         Begin VB.Menu TransferApp4 
            Caption         =   "Application 4"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu TransferAT 
         Caption         =   "Append and &transfer"
      End
      Begin VB.Menu MsgSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EmailScan 
         Caption         =   "S&can for PGP messages"
         Shortcut        =   ^N
      End
      Begin VB.Menu EmailShow 
         Caption         =   "Show PGP &messages..."
         Shortcut        =   ^P
      End
      Begin VB.Menu MsgSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EmailWrap 
         Caption         =   "Word &wrap message"
         Shortcut        =   ^W
      End
      Begin VB.Menu MsgSep3 
         Caption         =   "-"
      End
      Begin VB.Menu TransferOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu NymMenu 
      Caption         =   "N&ym"
      Begin VB.Menu TransferNym 
         Caption         =   "&Create nym..."
      End
      Begin VB.Menu NymPasswd 
         Caption         =   "C&hange nym password..."
      End
      Begin VB.Menu NymReplyChange 
         Caption         =   "Change nym &reply block..."
      End
      Begin VB.Menu NymDelete 
         Caption         =   "&Delete nym..."
      End
      Begin VB.Menu NymShow 
         Caption         =   "&Show nym server stats..."
      End
      Begin VB.Menu NymNull1 
         Caption         =   "-"
      End
      Begin VB.Menu NymConfigure 
         Caption         =   "Ed&it existing alias.net nym..."
      End
      Begin VB.Menu TransferPrepare 
         Caption         =   "&Prepare nym message... "
         Shortcut        =   ^Y
      End
      Begin VB.Menu TransferEncrypt 
         Caption         =   "&Encrypt nym message"
      End
   End
   Begin VB.Menu WebMenu 
      Caption         =   "&Web"
      Begin VB.Menu WebJump 
         Caption         =   "&Anonymous jump to URL"
         Shortcut        =   ^J
      End
      Begin VB.Menu Web1 
         Caption         =   "Web1"
         Index           =   1
      End
      Begin VB.Menu WebSep 
         Caption         =   "-"
      End
      Begin VB.Menu WebOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu HelpInfo 
         Caption         =   "&Information..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu SendSysInfo 
         Caption         =   "PI &Diagnostics"
      End
      Begin VB.Menu HelpStep 
         Caption         =   "&Step me through..."
         Begin VB.Menu StepEncrypt 
            Caption         =   "Encrypting a message"
         End
         Begin VB.Menu StepDecrypt 
            Caption         =   "Decrypting a message"
         End
         Begin VB.Menu StepSign 
            Caption         =   "Signing a message"
         End
         Begin VB.Menu StepSend 
            Caption         =   "Sending a message"
         End
         Begin VB.Menu StepAttach 
            Caption         =   "Sending an attachment"
         End
         Begin VB.Menu StepSep1 
            Caption         =   "-"
         End
         Begin VB.Menu StepSendKey 
            Caption         =   "Giving my PGP key to someone"
         End
         Begin VB.Menu StepGetKey 
            Caption         =   "Getting/sending MIT server keys"
         End
         Begin VB.Menu StepAddKey 
            Caption         =   "Adding a PGP key to my public key ring"
         End
         Begin VB.Menu StepDelete 
            Caption         =   "Deleting a PGP key from my key ring"
         End
         Begin VB.Menu StepCreateKey 
            Caption         =   "Creating a new PGP key pair"
         End
         Begin VB.Menu StepSep2 
            Caption         =   "-"
         End
         Begin VB.Menu StepRemailer 
            Caption         =   "Sending an anonymous message"
         End
         Begin VB.Menu StepUSENET 
            Caption         =   "Posting an anonymous USENET article"
         End
         Begin VB.Menu StepUpdateInfo 
            Caption         =   "Updating remailer information"
         End
         Begin VB.Menu StepSep3 
            Caption         =   "-"
         End
         Begin VB.Menu StepNym 
            Caption         =   "Creating a nym"
         End
         Begin VB.Menu StepNymSend 
            Caption         =   "Sending a nym message"
         End
         Begin VB.Menu StepNymPass 
            Caption         =   "Changing a nym password"
         End
         Begin VB.Menu StepNymReply 
            Caption         =   "Changing a nym reply block"
         End
         Begin VB.Menu StepNymDelete 
            Caption         =   "Deleting a nym"
         End
         Begin VB.Menu StepSep4 
            Caption         =   "-"
         End
         Begin VB.Menu StepWeb 
            Caption         =   "Anonymously accessing a Web page"
         End
         Begin VB.Menu StepInfo 
            Caption         =   "Getting Internet information"
         End
      End
      Begin VB.Menu HelpSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu HelpSend 
         Caption         =   "&Send feedback"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu HelpSys 
         Caption         =   "Add system info"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Checkbox handler for adding an attachment.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added capability to use user's
'               preferred path for files.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
      
    On Error GoTo AttachError

    If Check1.Value = 0 Then
        Label15.Caption = ""
    Else
        '---------------------------------------------
        'get the file to be attached to the message
        '---------------------------------------------
        CommonDialog1.DialogTitle = "Open file to attach"
        CommonDialog1.Flags = &H2& + &H4&
        CommonDialog1.Filter = "All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = gPathName
        CommonDialog1.Action = 1
        Label15.Caption = LCase(CommonDialog1.FileName)
        ChDrive Mid$(App.Path, 1, 3)
        ChDir App.Path
    End If
    Exit Sub

AttachError:
    If Err = 32755 Then
        Check1.Value = False
    End If
    Exit Sub
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

Private Sub Command2_Click()
    Text5.SetFocus
    If TransferAES.Checked Then
        AppendInfo
    End If
    If CheckConnection Then
        SendMailMessage
    End If
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub DoConnect_Click()
    CheckConnection
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditClrAll_Click()
    Combo2.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub


'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditClrMsg_Click()
    Text5.Text = ""
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditCopy_Click()
    On Error GoTo CopyError
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
    Exit Sub

CopyError:
    MsgBox Err.Description + "EditCopy. Can't copy from combo."
    Exit Sub
    Resume
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditCopyMsg_Click()
    Dim tempString
    Dim length
    Dim start
    If Len(Text5) > 0 Then
        Clipboard.Clear
        start = Text5.SelStart
        length = Text5.SelLength
        tempString = Mid(Text5, start, length)
    '    Clipboard.SetText Text5.Text
        Clipboard.SetText tempString
    End If
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditCut_Click()
    On Error GoTo CutError
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
    Screen.ActiveControl.SelText = ""
    Exit Sub

CutError:
    MsgBox Err.Description + "EditCut.  Can't cut from combo."
    Exit Sub
    Resume
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditHeader_Click()
    Text5.SelText = gHeader
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditOptions_Click()
    Form17.Show 1
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditPaste_Click()
    On Error GoTo PasteError
    Screen.ActiveControl.SelText = Clipboard.GetText()
    Exit Sub

PasteError:
    MsgBox Err.Description + "EditPaste.  Can't paste to combo."
    Exit Sub
    Resume
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditPasteMsg_Click()
    Text5.Text = Clipboard.GetText()
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1997 BC
'Author:        BC
'Description:   Capability for CTRL-A select all
'               in the message area.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/15/97
'Description:   Created.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Public Sub EditSelectAll_Click()
    Text5.SetFocus
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5)
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EditSig_Click()
    Text5.SetFocus
    Text5.SelStart = Len(Form1.Text5.Text)
    Text5.SelText = CRLF + CRLF + gSig
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EmailScan_Click()
    Dim Temps
    
    On Error GoTo ScanError
    
    If CheckConnection Then
        Label13.Caption = "Connecting to mail server for scan"
        gServerState = POPSTATE
        Form1.POP1.WinsockLoaded = True
        Form1.POP1.MaxLineLength = 30000
        POPLogin
        ConnectPOP
        If Not gPOPPass = "HSLAMXYDELKSHALP" Then
           'password okay
            gNumMessages = Form1.POP1.MessageCount
            If gPOPPass <> "@_@" Then
                gPOPPass = ""
                gFoundMessages = ""
                gFoundMessages = ScanTopHeaders("-----BEGIN PGP MESSAGE-----")
                If gFoundMessages <> "" Then
                    ParseFoundMessages (gFoundMessages)
                End If
            End If
            '---------------------------------------------
            'need this to enter the update state
            'we be done, so...
            '---------------------------------------------
            Form1.POP1.Action = 2
            Form1.POP1.WinsockLoaded = False
            gServerState = 0
        Else
           'password okay
            Form1.POP1.Action = 2
            Form1.POP1.WinsockLoaded = False
        End If
    End If
    Label13 = ""
    Exit Sub
ScanError:
    Label13 = ""
    Select Case Err
        Case 20113
            '---------------------------------------------
            'Last action tied up network, so disconnect
            '---------------------------------------------
            Form1.POP1.Action = 2
            Form1.POP1.WinsockLoaded = False
            Exit Sub
        Case 25066
            '---------------------------------------------
            '[10065] No route to host.
            '---------------------------------------------
'            MsgBox "This is an ISP connection error.  Apparently, you have become disconnected from your Internet Service Provider.  Please connect to your ISP and retry."
'            Exit Sub
        Case 26005
            '---------------------------------------------
            '[11004] Valid name, no data record (check DNS setup)
            '---------------------------------------------
'            #If win16 Then
'                MsgBox "FATAL INTERNET CONNECTION ERROR.  You tried to perform an internet operation without connecting to your Internet Service Provider first.  The only way to clear this error is to REBOOT your machine.  Please try again later after you reboot and connect to your ISP first."
'            #Else
'                Resume
'            #End If
'            Exit Sub
        Case Else
'            MsgBox "This is an ISP connection error.  Please click on Help Send Feedback and send the attached message through your regular e-mail program.  This will enable you to obtain technical support for your problem."
'            Exit Sub
    End Select
    gServerState = 0
    Resume
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EmailShow_Click()
    Form70.Show 1
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub EmailWrap_Click()
    Text5.Text = InsertCRLFs()
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileAddress_Click()
    Form23.Show 1
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileConnect_Click()
    '   IPPort1.WinsockLoaded = True
    FileConnect.Enabled = False
    FileDisconnect.Enabled = True
    Command2.Enabled = True
    TransferSend.Enabled = True
    EmailScan.Enabled = True
    '  Label13.Caption = "e-mail status"
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileDisconnect_Click()
    '   IPPort1.WinsockLoaded = False
    FileDisconnect.Enabled = False
    FileConnect.Enabled = True
    Command2.Enabled = False
    TransferSend.Enabled = False
    EmailScan.Enabled = False
    ' Label13.Caption = "Winsock not connected"
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileExit_Click()
    Unload Me
    End
End Sub

Private Sub FileExport_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu entry for saving text.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC @@
'Change Date:   04/18/97
'Description:   Repaired bug that does not set PI
'               back to its working drive if the
'               user saves a file on a different
'               drive.
'---------------------------------------------
'Developer:     BC @@
'Change Date:   04/18/97
'Description:   Added capability to point to
'               a preferred path for saving
'               files.
'---------------------------------------------
'Developer:     BC  @@
'Change Date:   04/18/97
'Description:   Added option to automatically
'               use the .txt extension on a
'               saved file if the user did not
'               specify an extension.
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    On Error GoTo ExportError
    
    '---------------------------------------------
    'prepare the file save as dialog
    '---------------------------------------------
    CommonDialog1.DialogTitle = "Save message as"
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = gPathName
    CommonDialog1.FileName = "*.txt"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Action = 2
    FileNum = FreeFile
    Open CommonDialog1.FileName For Output As FileNum
    Print #FileNum, Text5.Text
    Close #FileNum
    '---------------------------------------------
    'switch back to normal settings
    '---------------------------------------------
    ChDrive Mid$(App.Path, 1, 3)
    ChDir App.Path
ExportError:
    Exit Sub
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileImport_Click()
    Dim FileNum As Integer
    On Error GoTo ImportError
    CommonDialog1.DialogTitle = "Open message text file"
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|Asc Files (*.asc)|*.asc|All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = gPathName
    CommonDialog1.FileName = "*.asc"
    CommonDialog1.DefaultExt = "asc"
    CommonDialog1.Action = 1
    FileNum = FreeFile
    Open CommonDialog1.FileName For Input As FileNum
    If LOF(FileNum) > 30000 Then
        MsgBox "File is too big to import as a message.", 16, "Import Error"
    Else
        While Not EOF(FileNum)
            Line Input #FileNum, textline$
            msgtext$ = msgtext$ + textline$ + Chr$(13) + Chr$(10)
        Wend
        If Len(Form1.Text5.Text) > 0 Then
            '---------------------------------------------
            'test first to see if message area is empty
            '---------------------------------------------
            temp = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Send Feedback")
            If temp = vbYes Then
                Text5.Text = msgtext$
            Else
                Exit Sub
            End If
        Else
            Text5.Text = msgtext$
        End If
    End If
    Close #FileNum
    ChDir App.Path
    Exit Sub

ImportError:
    x$ = Err.Description(Err)
    Exit Sub
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileMode_Click()
    If FileMode.Caption = "User" Then
        ExpertMode
        FileMode.Caption = "Expert"
    Else
        NoviceMode
        FileMode.Caption = "User"
    End If
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileOptions_Click()
    Form7.Show 1
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FilePage_Click()
    MsgBox "Not implemented yet."
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FilePrintM_Click()
    Dim Buffer As String
    ndx& = fGetLineCount&()
    For N& = 1 To ndx&
        Buffer = fGetLine(N& - 1)
        Printer.Print Buffer
    Next N&
    Printer.EndDoc
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub FileSave_Click()
    SaveSettings
End Sub

Private Sub Form_Load()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Main loading routine for PI.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added ExecCmd function for 32-bit
'               version.
'---------------------------------------------

'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Dim FileNum As Integer
    Dim where As Integer
    Dim first As Integer
    Dim tmp As Integer
    Dim Stimer As Single
    Dim SectionName As String
    On Error GoTo LoadError
    '---------------------------------------------
    'init stuff
    '---------------------------------------------
    'Set PI version here
    HTTP1.UserAgent = PiStr + " 2.8c"
    '---------------------------------------------
    first = 0
    Obscurity = 0
    RemailerType = Cypherpunk
    NewsgroupType = 0
    PIPIF = "pipgp.pif"
    MinState = 1
    AppINI = "PIDAHO.INI"
    CutStr = ""
    LatentStr = ""
    CRLF = vbCr & vbCrLf
    gSMTPLog = 0
    gNymState = 0
    gPrivInit = 0
    gPOPPass = ""
    gPassPhrase = ""
    PiStr = "Private Idaho"
    gExit = 0
    'need to decrypt the rings first
    'were the key rings previously encrypted?
    Dim k As String
    k = Environ$("PGPPATH") + "\SECRING.PID"
    '++++++++
'    MsgBox "Here 1"
    
    If iFileExists(k) Then
        gPassPhrase = "2"
        frmPOPPassword.Caption = "Key ring Passphrase"
        frmPOPPassword.Label2.Caption = "Enter the passphrase to decrypt your key rings:"
        frmPOPPassword.Show 1
        Screen.MousePointer = vbHourglass
        Dim foo As String
        Dim cmd As String
        Dim Result As Integer
        Dim Ecount As Integer
        foo = StripExt(k) + "PGP"
        cmd$ = App.Path + "\PIPGPX.PIF +batchmode " + k$ + " -o " + foo$ + " -z " + Chr$(34) + gPassPhrase + Chr$(34)
        CheckLen (cmd$)
        #If win16 Then
            Result = Shell(cmd$, 1)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        
        'delete the version with the pgp extension
        If iFileExists(foo$) Then
            Kill k$
        Else
            MsgBox "Couldn't decrypt SECRING.PGP.  Wrong password.  Run Private Idaho and try again.", 16
            Unload Me
        End If
        Screen.MousePointer = 0
    End If
    '++++++++
    'MsgBox "Here 2"
    k = Environ$("PGPPATH") + "\PUBRING.PID"
    If iFileExists(k$) Then
        Screen.MousePointer = 11
        foo$ = StripExt(k$) + "PGP"
        cmd$ = App.Path + "\PIPGPX.PIF +batchmode " + k$ + " -o " + foo$ + " -z " + Chr$(34) + gPassPhrase + Chr$(34)
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, 1)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        'delete the version with the pgp extension
        If iFileExists(foo$) Then
            Kill k$
        End If
        Screen.MousePointer = 0
    End If
    where = 1
    'secure mode means store the INI and NYM file encrypted
    '++++++++
    'MsgBox "Here 3"
    If Command$ <> "" Then
        If InStr(1, UCase$(Command$), "PIDAHO.INI") <> 0 Then
            gPassPhrase = "2"
            frmPOPPassword.Caption = "PGP Passphrase"
            frmPOPPassword.Label2.Caption = "Enter your PGP passphrase to decrypt configuration files:"
            frmPOPPassword.Show 1
            Screen.MousePointer = vbHourglass
            'decrypt the .INI file
            foo$ = StripExt(Command$) + "ASC"
            If iFileExists(foo$) Then
                cmd$ = App.Path + "\PIPGPX.PIF +batchmode " + foo$ + " -o " + Command$ + " -z " + Chr$(34) + gPassPhrase + Chr$(34)
                CheckLen (cmd$)
                #If win16 Then
                    Result% = Shell(cmd$, 2)
                    Do While GetModuleUsage(Result%) > 0
                        Ecount = DoEvents()
                    Loop
                #Else
                    ExecCmd (cmd$)
                #End If
                'success, so delete the old
                If iFileExists(Command$) Then
                    Kill foo$
                Else
                    MsgBox "Couldn't decrypt PIDAHO.INI.  Wrong password.  Run Private Idaho and try again.", 16
                    Screen.MousePointer = 0
                    Unload Me
                End If
            End If
        Else
            MsgBox "PIDAHO.INI not specified in the command line.", 16
            End
        End If
        'decrypt the nym file
        If iFileExists(App.Path + "\NYMS.ASC") Then
            cmd$ = App.Path + "\PIPGPX.PIF " + App.Path + "\NYMS.ASC" + " -o " + App.Path + "\NYMS.TXT -z " + Chr$(34) + gPassPhrase + Chr$(34)
            CheckLen (cmd$)
            #If win16 Then
                Result% = Shell(cmd$, 2)
                Do While GetModuleUsage(Result%) > 0
                    Ecount = DoEvents()
                Loop
            #Else
                ExecCmd (cmd$)
            #End If
            If iFileExists(App.Path + "\NYMS.TXT") Then
                Kill App.Path + "\NYMS.ASC"
            End If
        End If
        'decrypt the pubkeys.out file
        If iFileExists(App.Path + "\PUBKEYS.ASC") Then
            cmd$ = App.Path + "\PIPGPX.PIF " + App.Path + "\PUBKEYS.ASC" + " -o " + App.Path + "\PUBKEYS.OUT -z " + Chr$(34) + gPassPhrase + Chr$(34)
            CheckLen (cmd$)
            #If win16 Then
                Result% = Shell(cmd$, 2)
                Do While GetModuleUsage(Result%) > 0
                    Ecount = DoEvents()
                Loop
            #Else
                ExecCmd (cmd$)
            #End If
            If iFileExists(App.Path + "\PUBKEYS.OUT") Then
                Kill App.Path + "\PUBKEYS.ASC"
            End If
        End If
        'decrypt the address file
        If iFileExists(App.Path + "\ADDRESS.ASC") Then
            cmd$ = App.Path + "\PIPGPX.PIF " + App.Path + "\ADDRESS.ASC" + " -o " + App.Path + "\ADDRESS.TXT -z " + Chr$(34) + gPassPhrase + Chr$(34)
            CheckLen (cmd$)
            #If win16 Then
                Result% = Shell(cmd$, 2)
                Do While GetModuleUsage(Result%) > 0
                    Ecount = DoEvents()
                Loop
            #Else
                ExecCmd (cmd$)
            #End If
            If iFileExists(App.Path + "\ADDRESS.TXT") Then
                Kill App.Path + "\ADDRESS.ASC"
            End If
        End If
        Screen.MousePointer = 0
    End If
    'splash screen
    Form12.Show
    Stimer = Timer + 1  ' Sleep one second
    where = 2
    ' Get the form attributes
    SectionName = "Settings"
    'get the window settings
    '++++++++
    'MsgBox "Here 4"
    Form1.Top = Val(ReadProfile(SectionName, "Top"))
    Form1.Left = Val(ReadProfile(SectionName, "Left"))
    wrapNum = Val(ReadProfile(SectionName, "Wrap"))
    If wrapNum = 0 Then wrapNum = -1
    tmp = Val(ReadProfile(SectionName, "Width"))
    If tmp > Form1.Width Then Form1.Width = tmp
    tmp = Val(ReadProfile(SectionName, "Height"))
    If tmp > Form1.Height Then Form1.Height = tmp
    where = 3
    'get the PGP Path
    SectionName = "PGP Info"
    piPGPPath = ReadProfile(SectionName, "PGPPath")
    '++++++++
    'MsgBox "Here 5"
    Dim x As Integer
    If piPGPPath = "" Then
        piPGPPath = Environ$("PGPPATH")
        If piPGPPath = "" Then
            'they don't have PGP, so tell them they need it and bail...
            x = MsgBox("You don't seem to have PGP installed correctly (couldn't find the PGPPATH variable in AUTOEXEC.BAT).  To get the most from Private Idaho, you need to install PGP.  You can easily install it with PGP QuickStart (use your Web browser to go to: http://www.eskimo.com/~joelm/pi.html).  Would you still like to load Private Idaho?", 48 + 4, PiStr)
            If x = 7 Then
                End
            End If
        End If
        WriteProfile SectionName, "PGPPath", piPGPPath
    End If
    where = 4
    'see if we want the passphrase loaded into memory
    If gPassPhrase = "" Then
        'hasn't been set yet
        If ReadProfile(SectionName, "Passphrase") = "true" Then
            gPassPhrase = "1"
            frmPOPPassword.Caption = "PGP Passphrase"
            frmPOPPassword.Label2.Caption = "Enter your PGP passphrase:"
            frmPOPPassword.Show 1
        End If
    End If
    where = 5
    'get the PGP key
    '++++++++
    'MsgBox "Here 6"
    PGPKeyID = ReadProfile(SectionName, "KeyID")
    If PGPKeyID = "" Then
        Form12.Hide
        MsgBox "Private Idaho needs to be configured.  You won't need to go through these steps the next time you run Private Idaho.", 64, PiStr
        'assume this is the first time run and generate
        'the pubkeys.out file
        first = 1
        'then display the options dialog to specify key id
        gPGPTempFile = "pvtidaho"
        Form6.Command1.Enabled = False
        Form6.Show 1
        cmd$ = Environ$("COMSPEC") + " /C " + piPGPPath + "\PGP.EXE -KV >" + App.Path + "\PUBKEYS.OUT"
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
    End If
    'make sure the key will be explicitly passed
    'this places the key in quotes
    PGPKeyID = Chr$(34) + PGPKeyID + Chr$(34)
    'get the e-mailer app info
    SectionName = "Mailer Info"
    gEmailer = ReadProfile(SectionName, "Emailer")
    If gEmailer = "" Then
        Form5.Command2.Enabled = False
        Form5.Show 1
        Form5.Command2.Enabled = True
    End If
    '++++++++
    'MsgBox "Here 7"
    tranScript = ReadProfile(SectionName, "tranScript")
    SectionName = "Options"
    'see if OS/2 delay should be turned on
    OS2 = ReadProfile(SectionName, "OS2")
    where = 6
    'get the e-mail settings
    gEmailAddress = ReadProfile(SectionName, "EmailAddress")
    gRealName = ReadProfile(SectionName, "RealName")
    gServerName = ReadProfile(SectionName, "ServerName")
    gPOPSeverName = ReadProfile(SectionName, "POPServerName")
    gPOPName = ReadProfile(SectionName, "POPAccount")
    If gEmailAddress = "" Then
        If first = 1 Then
            Form7.Show 1
        Else
            WriteProfile SectionName, "EmailAddress", gEmailAddress
        End If
    End If
    'see if SMTP logging should be turned on
    foo$ = ReadProfile(SectionName, "SMTPLog")
    If foo$ <> "" Then
        gSMTPLog = 1
        gSMTPFile = FreeFile
        Open App.Path + "\netlog.txt" For Output As gSMTPFile
    End If
    'get PGP temp file info
    '++++++++
    'MsgBox "Here 8"
    SectionName = "PGP Info"
    gPGPTempFile = ReadProfile(SectionName, "PGPTempFile")
    'it should never be this state, but...
    If gPGPTempFile = "" Then
        gPGPTempFile = "pvtidaho"
        WriteProfile SectionName, "PGPTempFile", gPGPTempFile
    End If
    SectionName = "Web Info"
    gBrowserPath = ReadProfile(SectionName, "BrowserPath")
    If gBrowserPath = "" Then
        gBrowserPath = "c:\netscape\netscape.exe"
        WriteProfile SectionName, "BrowserPath", gBrowserPath
    End If
    gBrowserString = ReadProfile(SectionName, "BrowserString")
    If gBrowserString = "" Then
        gBrowserString = "Netscape - ["
        WriteProfile SectionName, "BrowserString", gBrowserString
    End If
    gURLStart = ReadProfile(SectionName, "URLStart")
    If gURLStart = "" Then
        gURLStart = "^l"
        WriteProfile SectionName, "URLStart", gURLStart
    End If
    gURLEnd = ReadProfile(SectionName, "URLEnd")
    If gURLEnd = "" Then
        gURLEnd = "~"
        WriteProfile SectionName, "URLEnd", gURLEnd
    End If
    c2WWWAnon = ReadProfile(SectionName, "c2WWWAnon")
    If c2WWWAnon = "" Then
        c2WWWAnon = "http://www.anonymizer.com:8080/"
        WriteProfile SectionName, "c2WWWAnon", c2WWWAnon
    End If
    '++++++++
    'MsgBox "Here 9"
    where = 7
    'set the transfer menus
    SectionName = "Options"
    foo$ = ReadProfile(SectionName, "App1Name")
    If foo$ <> "" Then
        TransferApps.Visible = True
        TransferApp1.Caption = foo$
        TransferApp1.Visible = True
    End If
    foo$ = ReadProfile(SectionName, "App2Name")
    If foo$ <> "" Then
        TransferApp2.Caption = foo$
        TransferApp2.Visible = True
    End If
    foo$ = ReadProfile(SectionName, "App3Name")
    If foo$ <> "" Then
        TransferApp3.Caption = foo$
        TransferApp3.Visible = True
    End If
    foo$ = ReadProfile(SectionName, "App4Name")
    If foo$ <> "" Then
        TransferApp4.Caption = foo$
        TransferApp4.Visible = True
    End If
    where = 8
    'get the address list
    '++++++++
    'MsgBox "Here 10"
    FileNum = FreeFile
    Dim Item As String
    If iFileExists(App.Path + "\ADDRESS.TXT") Then
        Open App.Path + "\ADDRESS.TXT" For Input As FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, Item$
            Combo2.AddItem Item$
        Loop
        Close FileNum
    End If
    'get the header/sig
    'if file exists then open
    FileNum = FreeFile
    '++++++++
    'MsgBox "Here 11"
    If iFileExists(App.Path + "\HEADER.TXT") Then
        Open App.Path + "\HEADER.TXT" For Input As FileNum
        gHeader = Input$(LOF(FileNum), FileNum)
        Close FileNum
    End If
    FileNum = FreeFile
    'if file exists then open
    If iFileExists(App.Path + "\SIG.TXT") Then
        Open App.Path + "\SIG.TXT" For Input As FileNum
        gSig = Input$(LOF(FileNum), FileNum)
        Close FileNum
    End If
    'set the newsgroup menu for cp state
    USENETGate.Visible = True
    USENETFi.Visible = False
    UseNetSoda.Visible = False
    'load the remailer array on startup
    InitializeRemailers (App.Path + "\remailer.htm")
    where = 9
    'restore settings stored in the ini
    '++++++++
    'MsgBox "Here 12"
    RestoreSettings
    where = 10
    'check the mode
    NoviceMode
    SectionName = "Settings"
    tmp = Val(ReadProfile(SectionName, "Mode"))
    If tmp = 0 Then
        ExpertMode
    Else
        NoviceMode
    End If
    where = 11
    ' Wait for our splash screen to show a while
    While Stimer > Timer
        DoEvents
    Wend
    where = 12
    'initialize the screen
    UpdateMainScreen
    'load the Web menu
    where = 13
    CreateWebMenu
    where = 14
    Unload Form12
    '++++++++
    'MsgBox "Here 13"
    where = 15
    If iFileExists(App.Path + "\HELLO.TXT") Then
        FirstIntro
    End If
    Exit Sub

LoadError:
    MsgBox Err.Description + " EUPGP Form Load " + Str$(where) + "-" + Str$(Err)
    Form12.Hide
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub Form_Resize()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Handler for resizing of Form1.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Disabled resizing of message box
'               now that user fonts are enabled.
'               May want to do something on this
'               again later.  @@
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    If Form1.WindowState = 0 Then
        If Form1.Width < 8880 Then Form1.Width = 8880
        If Form1.Height < 7035 Then Form1.Height = 7035
'        Form1.Text5.Width = Form1.Width - 600
'        Select Case wrapNum
'        Case -1
'            Form1.Text5.Width = Form1.Width - 705
'        Case 65
'            Form1.Text5.Width = 8175
'        Case 70
'            Form1.Text5.Width = 8700
'        Case 75
'            Form1.Text5.Width = 9375
'        End Select
'        Text5.Height = Form1.Height - 4005
'        Text5.Height = Form1.Height - 5000
    End If
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub Form_Unload(Cancel As Integer)
    gExit = 1
    SaveSettings
    If gSMTPLog = 1 Then Close gSMTPFile
    Set Form1 = Nothing
    End
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub GetNews_Click()
    SendNewsCommand
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub GetFinger_Click()
    If CheckConnection Then
        SendFinger
    End If
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub HelpAbout_Click()
    Form2.Show 1
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub HelpInfo_Click()
    ShowWebHelp
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub HelpSys_Click()
    HelpSys.Checked = Not HelpSys.Checked
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub HTTP1_EndTransfer()
    On Error GoTo HTTPEndTransferErr
    Select Case gWebState
    Case GETREMAILERUPDATE
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path + "\remailer.htm", gWebPage
        InitializeRemailers (App.Path + "\remailer.htm")
        FillRemailerList
    Case MIXUPDATE
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path + "\mixmaste.htm", gWebPage
        InitializeMixRemailers (App.Path + "\mixmaste.htm")
        FillRemailerList
    Case TYPE2UPDATE
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path + "\type2.lis", gWebPage
    Case PUBRINGUPDATE
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path + "\pubring.mix", gWebPage
    Case GETSERVERKEY
        gWebPage = ConvertStr(gWebPage)
        Text5.SelText = GetKeyBlock(gWebPage)
        Text5.SetFocus
    Case GETREMAILERKEYS
        gWebPage = ConvertStr(gWebPage)
        Text5.SelText = gWebPage
        Text5.SetFocus
    End Select
    Screen.MousePointer = 0
    Form1.HTTP1.WinsockLoaded = False
    gWebState = HTTPIDLE
    HideStatus
    Exit Sub

HTTPEndTransferErr:
    MsgBox Err.Description + "HTTPEndTransferErr"
    Exit Sub
End Sub

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
Private Sub HTTP1_Error(ErrorCode As Integer, Description As String)
    gWebState = IdleState
    Form1.HTTP1.WinsockLoaded = False
    MsgBox "HTTP error " + Format$(ErrorCode) + ".  " + Description, 16, "HTTP Error"
End Sub

Private Sub HTTP1_StartTransfer()
'Workfile:      EUGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Initialize for an internet transfer
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    '---------------------------------------------
    'gWebPage will hold the text contents of the downloaded web page
    '---------------------------------------------
    gWebPage = ""
    Screen.MousePointer = 11
End Sub

Private Sub HTTP1_Transfer(BytesTransferred As Long, Text As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Do an HTTP transfer.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    On Error GoTo HTTPTransferErr
    
    Form33.Label3.Caption = Format$(BytesTransferred)
    If BytesTransferred > 62000 Then
        MsgBox ("64k string size exceeded.  HTML will be truncated.")
    Else
        gWebPage = gWebPage + Text
    End If
    Exit Sub

HTTPTransferErr:
    MsgBox Err.Description + "HTTPTransfer"
    Exit Sub
End Sub

Private Sub IPInfo1_RequestComplete(RequestID As Long, StatusCode As Integer, Description As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    gBigTest1 = RequestID
    gBigTest2 = StatusCode
    gBigTest3 = Description
End Sub

Private Sub IPPort1_DataIn(Text As String, EOL As Integer)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    TLen = Len(Text)
    Done = False
    Prechopped = True
    MyText = ""
    pos = 1
    Do While Not Done
        x = Mid(Text, pos, 1) '
        If x = vbLf Then
            Text5.SelText = MyText + CRLF
            MyText = ""
            Prechopped = False
        ElseIf x = vbCrLf Then
            Text5.SelText = MyText + CRLF
            MyText = ""
            Prechopped = False
       Else
           MyText = MyText + x
        End If
        pos = pos + 1
        If pos > TLen Then
            Done = True
            If Prechopped Then
                Text5 = Text5 + Text + CRLF
            End If
        End If
    Loop
End Sub
Private Sub IPPort1_Disconnected(StatusCode As Integer, Description As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Label13 = "Status " & Description + CRLF
End Sub

Private Sub IPPort1_Error(ErrorCode As Integer, Description As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Text5 = ErrorCode & ": " & Description
End Sub

Private Sub IPPort1_ReadyToSend()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If Len(Text5.Text) > 0 Then
        '---------------------------------------------
        'test first to see if message area is empty
        '---------------------------------------------
        temp = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Send Feedback")
        If temp = vbYes Then
            Text5.Text = ""
        Else
            IPPort1.Connected = False
            Label13 = "Finger operation cancelled."
        End If
    End If
    If Len(Text5) = 0 Then
        Text5 = "Ready to transmit..."
    End If
End Sub

Private Sub KeyCertify_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1997 BC
'Author:        BC
'Description:   Menu option to sign a public key.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97.
'Description:   Created.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    On Error GoTo KeyCertifyError
    keyID = ""
    gCancelButton = False
    
    '---------------------------------------------
    'display the list of personal keys to sign with
    '---------------------------------------------
    z$ = Form26.Label1.Caption
    Form26.Label1.Caption = "Select a key to use for signing the public key."
    CheckMultipleKey
    
    '---------------------------------------------
    'display the public key ring
    '---------------------------------------------
    storStr1$ = Form8.Label2.Caption
    storStr2$ = Form8.Label1.Caption
    Form8.Label2.Caption = "Select a key to sign"
    Form8.Label1.Caption = "from the public key ring."
    Form8.Show 1
    Form8.Label2.Caption = storStr1$
    Form8.Label1.Caption = storStr2$
    Unload Form8
    
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd$ = ""
        cmd$ = App.Path + "\" + PIPIF + " -ks " + Chr$(34) + keyID + Chr$(34) + " -u " + Chr$(34) + PGPKeyID + Chr$(34)
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, 3)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        cmd$ = Environ$("COMSPEC") + " /C PGP -KV >" + App.Path + "\PUBKEYS.OUT"
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
    Exit Sub
KeyCertifyError:
    If Err = 32755 Then
        Exit Sub
        MsgBox "Could not decrypt the message.  Suggest you open a DOS session and type pgp " + gPGPTempFile + ".out at the the command prompt in your PGP directory to find out what happened."
    ElseIf Err = 53 Then 'no pgpTempFile, it means PGP command did not complete, just abort
    End If
    Exit Sub
    Resume Next
End Sub

Private Sub KeyCreate_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   ? @@
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added ExecCmd function for 32-bit
'               version.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    CreatePGPKeyPair
    cmd$ = Environ$("COMSPEC") + " /C PGP -KV >" + App.Path + "\PUBKEYS.OUT"
    CheckLen (cmd$)
    #If win16 Then
        Result% = Shell(cmd$, MinState)
        Do While GetModuleUsage(Result%) > 0
            Ecount = DoEvents()
        Loop
    #Else
        ExecCmd (cmd$)
    #End If
End Sub

Private Sub KeyEditTrust_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1997 BC
'Author:        BC
'Description:   Menu option to sign a public key.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97.
'Description:   Created.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    keyID = ""
    gCancelButton = False
    storStr1$ = Form8.Label2.Caption
    storStr2$ = Form8.Label1.Caption
    Form8.Label2.Caption = "Select a key to sign"
    Form8.Label1.Caption = "from the public key ring."
    Form8.Show 1
    Form8.Label2.Caption = storStr1$
    Form8.Label1.Caption = storStr2$
    Unload Form8
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd$ = ""
        cmd$ = App.Path + "\" + PIPIF + " -ke " + Chr$(34) + keyID + Chr$(34)
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, 3)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        cmd$ = Environ$("COMSPEC") + " /C PGP -KV >" + App.Path + "\PUBKEYS.OUT"
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
End Sub
Private Sub KeySign_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1997 BC
'Author:        BC
'Description:   Menu option to sign a personal key.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97.
'Description:   Created.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    keyID = ""
    gCancelButton = False
    storStr1$ = Form8.Label2.Caption
    storStr2$ = Form8.Label1.Caption
    Form8.Label2.Caption = "Select your key from"
    Form8.Label1.Caption = "the public key ring."
    Form8.Show 1
    Form8.Label2.Caption = storStr1$
    Form8.Label1.Caption = storStr2$
    Unload Form8
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd$ = ""
        cmd$ = App.Path + "\" + PIPIF + " -ks " + Chr$(34) + keyID + Chr$(34) + " -u " + Chr$(34) + keyID + Chr$(34)
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, 3)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        cmd$ = Environ$("COMSPEC") + " /C PGP -KV >" + App.Path + "\PUBKEYS.OUT"
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
End Sub

Private Sub KeyOptions_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    MultiType = 1
    Form25.Show 1
End Sub

Private Sub KeySubmit_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If CheckConnection Then
        'is there a key there?
        If InStr(1, Form1.Text5.Text, "PUBLIC KEY BLOCK") = 0 Then
            MsgBox "Please enter the PGP public key you'd like to submit to the key server in the Message box."
            Exit Sub
        End If
        Combo2.Text = "pgp-public-keys@keys.us.pgp.net"
        Text2.Text = "add"
        SendMailMessage
    End If
End Sub

Private Sub Label1_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Processes label1 click to pull up public key ring.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Upgraded to strip off quotes
'               added by another new process
'               in PI.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    Dim Pos1 As Integer
    Dim Pos2 As Integer
        
    '---------------------------------------------
    'clear any current name from the "to:" box
    '---------------------------------------------
    Combo2.Text = ""
    keyID = ""
    '---------------------------------------------
    'present the list of public keys
    '---------------------------------------------
    Form8.Label2 = ""
    Form8.Label1 = "Select a recipient from your public key ring"
    Form8.Show 1
    '---------------------------------------------
    'process the key, if any
    '---------------------------------------------
    If keyID <> "null" Then
        '---------------------------------------------
        'if user has selected a key, strip off quotes
        '---------------------------------------------
        Pos2 = Len(keyID)
        tempstr = ""
        For i = 1 To Pos2
            x = Mid(keyID, i, 1)
            If (x <> Chr$(34)) Then
                tempstr = tempstr + x
            End If
        Next
        Combo2.Text = LTrim(tempstr)
    End If
End Sub

Private Sub Label13_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Label13.Caption = ""
End Sub


Private Sub Label5_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form29.Show 1
End Sub

Private Sub List3_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If List3.ListIndex = 0 Then
        Check1.Enabled = True
    Else
        Check1.Value = 0
        Label15.Caption = ""
        Check1.Enabled = False
    End If
End Sub

Private Sub NNTP1_GroupList(Group As String, FirstArticle As Long, LastArticle As Long, CanPost As Boolean)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
     temp1 = Group
     temp2 = FirstArticle
     temp3 = LastArticle
End Sub


Private Sub IPPort1_Connected(StatusCode As Integer, Description As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If 0 = StatusCode Then    'OK
        Label13 = ""
    Else
        Label13 = "Connection failed: " & Description
    End If
End Sub


Private Sub NNTP1_GroupOverview(ArticleNumber As Long, Subject As String, From As String, ArticleDate As String, MessageID As String, References As String, ArticleSize As Long, ArticleLines As Long, OtherHeaders As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    'ArticleNumber contains the number of the article within the group.
    temp1 = ArticleNumber
    'Subject contains the subject of the article.
    temp2 = Subject
    'From contains the email address of the article author.
    temp3 = From
    'ArticleDate contains the date the article was posted.
    temp4 = ArticleDate
    'MessageID contains the unique message id for the article.
    Temp5 = MessageID
    'References contains the message ids for the articles this article refers to (separated by spaces).
    Temp6 = References
    'ArticleSize contains the size of the article in bytes.
    Temp7 = ArticleSize
    'ArticleLines contains the number of lines in the article.
    Temp8 = ArticleLines
    'OtherHeaders contains any other article headers that NewsServer chooses to display for the article.
    Temp9 = OtherHeaders
End Sub

Private Sub NNTP1_Header(Field As String, Value As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    'The Field parameter contains the name of the header (same case as it is delivered).
    temp1 = Field
    'The Value parameter contains the header contents.
    temp2 = Value
    'If the header line being retrieved is a continuation header line, then the Field parameter contains "" (empty string).
End Sub

Private Sub NNTP1_Transfer(BytesTransferred As Long, Text As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    'The Text parameter contains the portion of article data being retrieved.
    temp1 = Text
    'The BytesTransferred parameter contains the number of bytes transferred since the beginning of the article, including header bytes.
    temp2 = BytesTransferred
End Sub

Private Sub NymConfigure_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If Not iFileExists(App.Path + "\NYMS.TXT") Then
        MsgBox "You need to create an alias.net nym before you can configure it.", vbApplicationModal + vbCritical, App.Title
    Else
        gNymConfigStr = ""
        x = MinState
        gNymState = NYMCONFIG
        MinState = 1
        Form27.Show 1
        gNymState = NYMIDLE
        MinState = x
    End If
End Sub

Private Sub NymDelete_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    gNymState = NYMDEL
    Form27.Show 1
    gNymState = NYMIDLE
End Sub

Private Sub NymPasswd_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    gNymState = NYMPASSCHANGE
    Form27.Show 1
    gNymState = NYMIDLE
End Sub

Private Sub NymReplyChange_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If RemailersCP.Checked = False Then
        MsgBox "Only Cypherpunk type remailers can be used for reply blocks.", 16, PiStr
    Else
        gNymState = NYMRPLYCHANGE
        Form27.Show 1
        gNymState = NYMIDLE
    End If
End Sub

Private Sub NymShow_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    gShowNymStatus = 1
    Form31.Show
    gShowNymStatus = 0
End Sub

Public Sub PGPAdd_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Updates the PUBKEYS.OUT file.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added comments.

'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    Dim MyTemp As Long
    
    cmd$ = Environ$("COMSPEC") + " /C " + piPGPPath + "\PGP.EXE -KV >" + App.Path + "\PUBKEYS.OUT"
    
    CheckLen (cmd$)
    
    #If win16 Then
        Result% = Shell(cmd$, MinState)
        Do While GetModuleUsage(Result%) > 0
            Ecount = DoEvents()
        Loop
    #Else
        ExecCmd (cmd$)
    #End If

End Sub
Private Sub PGPAddKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   ? @@
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added ExecCmd function for 32-bit
'               version.
'---------------------------------------------
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added request to clear message
'               area after processing add key.
'               version.
'---------------------------------------------
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added check for empty message.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If Len(Text5) = 0 Then
        Text5 = "No key found ..."
        Exit Sub
    End If
    ClipText$ = Form1.Text5.Text
    FileNum = FreeFile
    If InStr(1, gPGPTempFile, ":") = 0 Then
        Open piPGPPath + "\" + gPGPTempFile + ".out" For Output As FileNum
    Else
        Open gPGPTempFile + ".out" For Output As FileNum
    End If
    Print #FileNum, ClipText$
    Close #FileNum
    where = 3
    If InStr(1, gPGPTempFile, ":") = 0 Then
        cmd$ = App.Path + "\" + PIPIF + " -ka " + piPGPPath + "\" + gPGPTempFile + ".out"
    Else
        cmd$ = App.Path + "\" + PIPIF + " -ka " + gPGPTempFile + ".out"
    End If
    CheckLen (cmd$)
    #If win16 Then
        Result% = Shell(cmd$, MinState)
        Do While GetModuleUsage(Result%) > 0
            Ecount = DoEvents()
        Loop
    #Else
        ExecCmd (cmd$)
    #End If
    If OS2 = "true" Then
        OS2Pause
    End If
    Close #FileNum
    If InStr(1, gPGPTempFile, ":") = 0 Then
        Kill piPGPPath + "\" + gPGPTempFile + ".out"
    Else
        Kill gPGPTempFile + ".out"
    End If
    cmd$ = Environ$("COMSPEC") + " /C PGP -KV >" + App.Path + "\PUBKEYS.OUT"
    CheckLen (cmd$)
    #If win16 Then
        Result% = Shell(cmd$, MinState)
        Do While GetModuleUsage(Result%) > 0
            Ecount = DoEvents()
        Loop
    #Else
        ExecCmd (cmd$)
    #End If
    If Len(Text5.Text) > 0 Then
        '---------------------------------------------
        'test to see if message area is empty
        '---------------------------------------------
        temp = MsgBox("Clear the message area?", vbYesNo, "Add Key Complete")
        If temp = vbYes Then
            Text5.Text = ""
        End If
    End If
End Sub
Private Sub PGPClose_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPClose.Checked = Not PGPClose.Checked
    If PGPClose.Checked = True Then
        PIPIF = "pipgpx.pif"
    Else
        PIPIF = "pipgp.pif"
    End If
End Sub

Private Sub PGPConvent_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPConvent.Checked = Not PGPConvent.Checked
End Sub

Private Sub PGPDecrypt_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   ? @@
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added ExecCmd function for 32-bit
'               version.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Dim FileNum As Integer
    On Error GoTo DecryptError
    
    '---------------------------------------------
    'set the PIF files for the decryption process
    '---------------------------------------------
    If Form1.PGPClose.Checked = True Then
        PIPIF = "pgpcmdx.pif"
    Else
        PIPIF = "pgpcmd.pif"
    End If
    If Not PGPFile.Checked Then
        If Obscurity = 1 Then
            Form1.Text5.SelStart = 0
            Form1.Text5.SelText = "-----BEGIN PGP MESSAGE-----" + CRLF + "Version: unknown" + CRLF + CRLF
            Form1.Text5.SelStart = Len(Form1.Text5.Text)
            Form1.Text5.SelText = "-----END PGP MESSAGE-----" + CRLF
        End If
        If Text5.Text = "" Then Exit Sub
        ClipText$ = Text5.Text
        FileNum = FreeFile
        If InStr(1, gPGPTempFile, ":") = 0 Then
            Open piPGPPath + "\" + gPGPTempFile + ".out" For Output As FileNum
        Else
            Open gPGPTempFile + ".out" For Output As FileNum
        End If
        Print #FileNum, ClipText$
        Close #FileNum
        FileNum = FreeFile
        Open piPGPPath + "\" + "pgpcmd.bat" For Output As FileNum
        cmd$ = "pgp " + piPGPPath + "\" + gPGPTempFile + ".out "
        Print #FileNum, cmd$
        Close #FileNum
        #If win16 Then
            Result% = Shell(App.Path + "\" + PIPIF, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$) '@@
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        FileNum = FreeFile
        If InStr(1, gPGPTempFile, ":") = 0 Then
            Open piPGPPath + "\" + gPGPTempFile For Input As FileNum
        Else
            Open gPGPTempFile For Input As FileNum
        End If
        While Not EOF(FileNum)
            Line Input #FileNum, textline$
            cyphertext$ = cyphertext$ + textline$ + Chr$(13) + Chr$(10)
        Wend
        Close #FileNum
        If gPOPState = POPDECRYPT Then
            cyphertext$ = "Subject: " + gMessageRecord.Subject + CRLF + CRLF + cyphertext$
            cyphertext$ = "Sent: " + gMessageRecord.SentDate + CRLF + cyphertext$
            cyphertext$ = "From: " + gMessageRecord.Sender + CRLF + cyphertext$
            gPOPState = 0
        End If
        Text5.Text = cyphertext$
        If InStr(1, gPGPTempFile, ":") = 0 Then
            Kill piPGPPath + "\" + gPGPTempFile
        Else
            Kill gPGPTempFile
        End If
        If InStr(1, gPGPTempFile, ":") = 0 Then
            WipeFile (piPGPPath + "\" + gPGPTempFile)
        Else
            WipeFile (gPGPTempFile)
        End If
    Else
        CommonDialog1.DialogTitle = "Open file to decrypt"
        CommonDialog1.Flags = &H2& + &H4&
        CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc|All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = gPathName
        CommonDialog1.Action = 1
        theName$ = CommonDialog1.FileName
        '---------------------------------------------
        'option to decrypt file to message area
        '---------------------------------------------
        temp = MsgBox("Would you like to put the output in the message window?", vbYesNo, "File Decrypt")
        If temp = vbYes Then
            '---------------------------------------------
            'check for text in the msg area
            '---------------------------------------------
            If Len(Text5) > 0 Then
            temp = MsgBox("Okay to clear the message area first?", vbYesNo, "Get Key From Server")
            If temp = vbYes Then
                Text5.Text = ""
                Else
                    Exit Sub
                End If
            End If
            FileNum = FreeFile
            '---------------------------------------------
            'fill the PGPcmd.bat file with the PGP command
            'I went to this approach to help with debugging.  @@
            '---------------------------------------------
            Open piPGPPath + "\" + "pgpcmd.bat" For Output As FileNum
            cmd$ = "pgp " + theName$ + " -o " + piPGPPath + "\" + gPGPTempFile
            Print #FileNum, cmd$
            Close #FileNum
            
            '---------------------------------------------
            'set the PIF files for the decryption process
            '---------------------------------------------
            If Form1.PGPClose.Checked = True Then
                PIPIF = "pgpcmdx.pif"
            Else
                PIPIF = "pgpcmd.pif"
            End If
            
            '---------------------------------------------
            'run the PGP decryption command
            '---------------------------------------------
            #If win16 Then
                Result% = Shell(App.Path + "\" + PIPIF, 1)
                Do While GetModuleUsage(Result%) > 0
                    Ecount = DoEvents()
                Loop
            #Else
                ExecCmd (App.Path + "\" + PIPIF)
            #End If
            If OS2 = "true" Then
                OS2Pause
            End If
            
            '---------------------------------------------
            'done, so reset the PIF files
            '---------------------------------------------
            If Form1.PGPClose.Checked = True Then
                PIPIF = "pipgpx.pif"
            Else
                PIPIF = "pipgp.pif"
            End If
            
            '---------------------------------------------
            'load the text into the message box
            '---------------------------------------------
            FileNum = FreeFile
            If InStr(1, gPGPTempFile, ":") = 0 Then
                Open piPGPPath + "\" + gPGPTempFile For Input As FileNum
            Else
                Open gPGPTempFile For Input As FileNum
            End If
            While Not EOF(FileNum)
                Line Input #FileNum, textline$
                cyphertext$ = cyphertext$ + textline$ + Chr$(13) + Chr$(10)
            Wend
            Close #FileNum
            Text5.Text = cyphertext$
            If InStr(1, gPGPTempFile, ":") = 0 Then
                Kill piPGPPath + "\" + gPGPTempFile
            Else
                Kill gPGPTempFile
            End If
            If InStr(1, gPGPTempFile, ":") = 0 Then
                WipeFile (piPGPPath + "\" + gPGPTempFile)
            Else
                WipeFile (gPGPTempFile)
            End If
        Else
            CommonDialog1.DialogTitle = "Save decrypted file as:"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = gPathName
            CommonDialog1.Action = 2
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            FileNum = FreeFile
            Open piPGPPath + "\" + "pgpcmd.bat" For Output As FileNum
            cmd$ = "pgp " + theName$ + " -o " + CommonDialog1.FileName
            Print #FileNum, cmd$
            Close #FileNum
            CheckLen (cmd$)
            #If win16 Then
                Result% = Shell(App.Path + "\" + PIPIF, MinState)
                Do While GetModuleUsage(Result%) > 0
                    Ecount = DoEvents()
                Loop
            #Else
                ExecCmd (cmd$)
            #End If
            If OS2 = "true" Then
                OS2Pause
            End If
        End If
    End If
    
    If Form1.PGPClose.Checked = True Then
        PIPIF = "pipgpx.pif"
    Else
        PIPIF = "pipgp.pif"
    End If
    
    Exit Sub

DecryptError:
    If Err = 6 Then 'overflow
        Resume Next
    ElseIf Err = 32755 Then
        Exit Sub
        MsgBox "Could not decrypt the message.  Suggest you open a DOS session and type pgp " + gPGPTempFile + ".out at the the command prompt in your PGP directory to find out what happened."
    ElseIf Err = 53 Then 'no pgpTempFile, it means PGP command did not complete, just abort
    End If
    
    If Form1.PGPClose.Checked = True Then
        PIPIF = "pipgpx.pif"
    Else
        PIPIF = "pipgp.pif"
    End If
    
    Exit Sub
    Resume Next
End Sub

Private Sub PGPDeleteKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   ? @@
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/02/97
'Description:   Added ExecCmd function for 32-bit
'               version.
'---------------------------------------------
'---------------------------------------------
'Developer:     BC
'Change Date:   05/13/97
'Description:   Added handler for cancel delete.
'---------------------------------------------
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    
    gCancelButton = False
    keyID = ""
    storStr1$ = Form8.Label2.Caption
    storStr2$ = Form8.Label1.Caption
    Form8.Label2.Caption = "Select which user's key to delete"
    Form8.Label1.Caption = "from your public ring."
    Form8.Show 1
    Form8.Label2.Caption = storStr1$
    Form8.Label1.Caption = storStr2$
    Unload Form8
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd$ = App.Path + "\" + PIPIF + " -kr " + " +force " + Chr$(34) + keyID + Chr$(34) + ""
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        cmd$ = Environ$("COMSPEC") + " /C PGP -KV >" + App.Path + "\PUBKEYS.OUT"
        CheckLen (cmd$)
        #If win16 Then
            Result% = Shell(cmd$, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
End Sub
Private Sub PGPEncrypt_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu entry for downloading new remailer data.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Extensively rewrote this subroutine.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added capability to look in
'               preferred path first for the
'               file(s)
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added code to switch back to
'               the application path after
'               opening the file
'---------------------------------------------
'Developer:     BC
'Change Date:   05/25/97
'Description:   Added option of saving message
'               to file, then encrypting the file
'---------------------------------------------
'Developer:     BC
'Change Date:   05/25/97
'Description:   Added test for filename in case
'               user is saving message to file then
'               encrypting
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    
    Dim PGPCmdString As String
    
    On Error GoTo FileEnSError

    
    '---------------------------------------------
    'verify there is something to encrypt
    '---------------------------------------------
    If (Len(Text5) = 0) And Not PGPFile.Checked Then
        Form1.Text5 = "Nothing to Encrypt!"
        Form1.Text5.SelStart = 0
        Form1.Text5.SelLength = Len(Form1.Text5)
        Exit Sub
    End If
    
    gCancelButton = False
    PGPCmdString = ""
    PGPKeyID = ""
    theName$ = ""
    
    '---------------------------------------------
    'handle case of saving a message to file, then encrypting
    '---------------------------------------------
    If Len(Text5) > 0 And PGPFile.Checked Then
        temp = MsgBox("Would you like to encrypt the message area to a file?", vbYesNo, "File Encrypt")
        If temp = vbYes Then
            CommonDialog1.DialogTitle = "Specify file for saving message"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = gPathName
            CommonDialog1.Action = 1
            theName$ = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            FileNum = FreeFile
            Open CommonDialog1.FileName For Output As FileNum
                Print #FileNum, Text5.Text
            Close #FileNum
        End If
    End If
    
    '---------------------------------------------
    'load the contents of the "to:" box
    '---------------------------------------------
    foo$ = Combo2.Text
    
    '---------------------------------------------
    'is this a eyes, conventional, other encrypt?
    '---------------------------------------------
    If Not PGPFile.Checked Then
        If PGPEyes.Checked = True Then
            PGPCmdString = "m"
        End If
        If PGPConvent.Checked = True Then
            EncryptMessage foo$, "-catw" + PGPCmdString
        Else
            EncryptMessage foo$, "-eatw" + PGPCmdString
        End If
    Else
    '---------------------------------------------
    'no, it is a file encrypt
    '---------------------------------------------
        If Len(theName$) = 0 Then
            '---------------------------------------------
            'user has not already specified a file name from above
            '---------------------------------------------
            CommonDialog1.DialogTitle = "Open file to encrypt"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = gPathName
            CommonDialog1.Action = 1
            theName$ = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
        End If
        If PGPEyes.Checked = True Then
            PGPCmdString = "m"
        End If
        If PGPConvent.Checked Then
            EncryptFile foo$, "-cat" + PGPCmdString, theName$
        Else
            EncryptFile foo$, "-eat" + PGPCmdString, theName$
        End If
    End If
FileEnSError:
        
    Exit Sub
End Sub

Private Sub PGPEnSign_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Handle menu item for encrypt and sign.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Extensively rewrote this subroutine.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added capability to look in
'               preferred path first for the
'               file(s)
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added code to switch back to
'               the application path after
'               opening the file
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added check for cancel button.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added check for empty message.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/25/97
'Description:   Added option of saving message
'               to file, then encrypting the file
'---------------------------------------------
'Developer:     BC
'Change Date:   05/25/97
'Description:   Added test for filename in case
'               user is saving message to file then
'               encrypting
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    
    Dim PGPCmdString As String
    
    On Error GoTo FileEnSError
    
    '---------------------------------------------
    'verify there is something to encrypt not a file
    '---------------------------------------------
    If (Len(Form1.Text5) = 0) And Not PGPFile.Checked Then
        Form1.Text5 = "Nothing to Encrypt!"
        Form1.Text5.SelStart = 0
        Form1.Text5.SelLength = Len(Form1.Text5)
        Exit Sub
    End If
    
    gCancelButton = False
    PGPCmdString = ""
    PGPKeyID = ""
    theName$ = ""
    
    '---------------------------------------------
    'handle case of saving a message to file, then encrypting
    '---------------------------------------------
    If Len(Text5) > 0 And PGPFile.Checked Then
        temp = MsgBox("Would you like to encrypt the message area to a file?", vbYesNo, "File Encrypt")
        If temp = vbYes Then
            CommonDialog1.DialogTitle = "Specify file for saving message"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = gPathName
            CommonDialog1.Action = 1
            theName$ = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            FileNum = FreeFile
            Open CommonDialog1.FileName For Output As FileNum
                Print #FileNum, Text5.Text
            Close #FileNum
        End If
    End If
    
    z$ = Form26.Label1.Caption
    Form26.Label1.Caption = "Select the key to sign the message with."
    '---------------------------------------------
    'display the list of personal keys to sign with
    'if key is selected, it is in global PGPKeyID
    '---------------------------------------------
    CheckMultipleKey
    
    If Not gCancelButton Then
        '---------------------------------------------
        'user has not canceled with form26
        '---------------------------------------------
        If Form26.Label1.Caption = "" Then
            Form26.Label1.Caption = z$
            Exit Sub
        End If
    
        '---------------------------------------------
        'get the recipient from the to: box, if any
        '---------------------------------------------
        foo$ = Combo2.Text
    
        '---------------------------------------------
        'is this an eyes, conventional, other encrypt?
        '---------------------------------------------
        If Not PGPFile.Checked Then
            If PGPEyes.Checked = True Then
                PGPCmdString = "m"
            End If
            If PGPConvent.Checked = True Then
                EncryptMessage foo$, "-satcw" + PGPCmdString
            Else
                EncryptMessage foo$, "-satew" + PGPCmdString
            End If
        Else
        '---------------------------------------------
        'no, it is a file encrypt
        '---------------------------------------------
            If Len(theName$) = 0 Then
                '---------------------------------------------
                'user has not already specified a file name from above
                '---------------------------------------------
                CommonDialog1.DialogTitle = "Open file to encrypt"
                CommonDialog1.Flags = &H2& + &H4&
                CommonDialog1.Filter = "All Files (*.*)|*.*"
                CommonDialog1.FilterIndex = 1
                CommonDialog1.CancelError = True
                CommonDialog1.InitDir = gPathName
                CommonDialog1.Action = 1
                theName$ = CommonDialog1.FileName
                ChDrive Mid$(App.Path, 1, 3)
                ChDir App.Path
            End If
            If PGPEyes.Checked = True Then
                PGPCmdString = "m"
            End If
            If PGPConvent.Checked Then
                EncryptFile foo$, "-scat" + PGPCmdString, theName$
            Else
                EncryptFile foo$, "-seat" + PGPCmdString, theName$
            End If
        End If
    Else
        '---------------------------------------------
        'user canceled at form26
        '---------------------------------------------
        Exit Sub
    End If
FileEnSError:
        Exit Sub
End Sub

Private Sub PGPEyes_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPEyes.Checked = Not PGPEyes.Checked
End Sub

Private Sub PGPFile_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu handler for PGP->file as input.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/25/97
'Description:   Revised for new PGP file operations.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    If Not PGPFile.Checked Then
        '---------------------------------------------
        'is the file option is not checked, this is the menu
        '---------------------------------------------
        PGPEncrypt.Caption = "&Encrypt file..."
        PGPDecrypt.Caption = "&Decrypt or verify file..."
        PGPEnSign.Caption = "Encrypt and &sign file..."
        PGPClearSign.Caption = "&Clear sign file..."
    Else
        '---------------------------------------------
        'is the file option is checked, this is the menu
        '---------------------------------------------
        PGPEncrypt.Caption = "&Encrypt message"
        PGPDecrypt.Caption = "&Decrypt or verify message"
        PGPEnSign.Caption = "Encrypt and &sign message"
        PGPClearSign.Caption = "&Clear sign message"
    End If
    '---------------------------------------------
    'toggle the state
    '---------------------------------------------
    PGPFile.Checked = Not PGPFile.Checked
End Sub
Private Sub PGPGetKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    On Error GoTo GetKeyError
    If CheckConnection Then
        If Len(Text5) > 0 Then
            '---------------------------------------------
            'test first to see if message area is empty
            '---------------------------------------------
            temp = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Get Key From Server")
            If temp = vbYes Then
                Text5.Text = ""
            Else
                Exit Sub
            End If
        End If
        If Len(Combo2) = 0 Then
            '---------------------------------------------
            'the user did not specify a recipient in the "to:" box
            '---------------------------------------------
            Text5 = "No user specified in the To: box." + CRLF
            Text5 = Text5 + "Please enter a valid e-mail address," + CRLF
            Text5 = Text5 + "or click on the right arrow button in the" + CRLF
            Text5 = Text5 + "To: box to choose a name from your address book."
            Exit Sub
        End If
        gServerState = HTTPSTATE
        gWebState = GETSERVERKEY
        Form1.HTTP1.WinsockLoaded = True
        If gGetKeyURL = "" Then
            Dim SectionName As String
            SectionName = "Net Info"
            gGetKeyURL = ReadProfile(SectionName, "GetKeyURL")
            If gGetKeyURL = "" Then
                gGetKeyURL = "http://swissnet.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
                WriteProfile SectionName, "GetKeyURL", gGetKeyURL
            End If
        End If
        ShowStatus ("Requesting key from server.  Please wait.  Hint: Use 'Add key from message' when done.")
        GetWebURL (gGetKeyURL + Form1.Combo2.Text)
        gServerState = 0
    End If
    Exit Sub
GetKeyError:
    HideStatus
    MsgBox Err.Description + "PGPGetKey"
    gServerState = 0
    Exit Sub
End Sub

Private Sub PGPInsertKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    InsertKey (1)
End Sub

Private Sub PGPMin_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPMin.Checked = Not PGPMin.Checked
    If PGPMin.Checked = True Then
        MinState = 2
    Else
        MinState = 1
    End If
End Sub

Private Sub PGPMultiple_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPMultiple.Checked = Not PGPMultiple.Checked
End Sub

Private Sub PGPObscurity_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If PGPObscurity.Checked = True Then
        PGPObscurity.Checked = False
        RemailersPGP.Enabled = True
        Obscurity = 0
    Else
        PGPObscurity.Checked = True
        RemailersPGP.Checked = False
        RemailersPGP.Enabled = False
        Obscurity = 1
    End If
End Sub

Private Sub PGPOptions_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form6.Show 1
End Sub

Private Sub PGPSelf_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPSelf.Checked = Not PGPSelf.Checked
End Sub

Private Sub PGPClearSign_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu handler for clear-signing the message
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Changed PGPSign -> PGPClearSign,
'               also in menu structure.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added option to temporarily turn off
'               encrypting with multiple keys.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added option to temporarily turn off
'               encrypting to self.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added option to temporarily turn off
'               obscurity.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added check for cancel button.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added check for empty message.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
        
    Dim PGPMutlipleToggle As Boolean
    Dim PGPEncryptSelfToggle As Boolean
    Dim PGPObscurityToggle As Boolean
    
    On Error GoTo ClearSignError
    
    '---------------------------------------------
    'first verify there is something to encrypt
    '---------------------------------------------
    If (Len(Text5) = 0) And Not PGPFile.Checked Then
        Form1.Text5 = "Nothing to Encrypt!"
        Form1.Text5.SelStart = 0
        Form1.Text5.SelLength = Len(Form1.Text5)
        Exit Sub
    End If
    
    PGPCmdString = ""
    PGPKeyID = ""
    gCancelButton = False
    
    z$ = Form26.Label1.Caption
    If PGPFile.Checked Then
        Form26.Label1.Caption = "Select the key to sign the file with."
    Else
        Form26.Label1.Caption = "Select the key to sign the message with."
    End If
    '---------------------------------------------
    'display the list of personal keys to sign with
    'if key is selected, it is in global PGPKeyID
    '---------------------------------------------
    CheckMultipleKey
    
    If Not gCancelButton Then
        '---------------------------------------------
        'user has not canceled at form26
        '---------------------------------------------
        'temporarily disable encrypting with multiple keys
        '---------------------------------------------
        PGPMultipleToggle = False
        If PGPMultiple.Checked Then
            PGPMultipleToggle = True
            PGPMultiple.Checked = False
        End If
        
        '---------------------------------------------
        'temporarily disable encrypting to self
        '---------------------------------------------
        PGPEncryptSelfToggle = False
        If PGPSelf.Checked Then
            PGPEncryptSelfToggle = True
            PGPSelf.Checked = False
        End If
        
        '---------------------------------------------
        'temporarily disable obscurity
        '---------------------------------------------
        PGPObscurityToggle = False
        If PGPObscurity.Checked Then
            PGPObscurityToggle = True
            PGPObscurity.Checked = False
        End If
        
        If Form26.Label1.Caption = "" Then
            Form26.Label1.Caption = z$
            Exit Sub
        End If
        
        foo$ = Combo2.Text
        
        '---------------------------------------------
        'is this a basic message encrypt
        '---------------------------------------------
        If Not PGPFile.Checked Then
            If PGPConvent.Checked Then
                EncryptMessage foo$, "-stac"
            Else
                EncryptMessage foo$, "-sta"
            End If
        Else
        '---------------------------------------------
        'no, it is a file signing
        '---------------------------------------------
            CommonDialog1.DialogTitle = "Select file to clear sign"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "Txt Files (*.txt)|*.txt"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = gPathName
            CommonDialog1.Action = 1
            theName$ = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            If PGPConvent.Checked Then
                EncryptFile foo$, "-stac", theName$
            Else
                EncryptFile foo$, "-sta", theName$
            End If
        End If
    
        '---------------------------------------------
        're-enable encrypting with multiple keys
        '---------------------------------------------
        If PGPMultipleToggle Then
            PGPMultiple.Checked = True
        End If
    
        '---------------------------------------------
        're-enable encrypting to self
        '---------------------------------------------
        If PGPEncryptSelfToggle Then
            PGPSelf.Checked = True
        End If
    
        '---------------------------------------------
        're-enable obscurity
        '---------------------------------------------
        If PGPObscurityToggle Then
            PGPObscurity.Checked = True
        End If
        Exit Sub
    Else
        '---------------------------------------------
        'user canceled at form26
        '---------------------------------------------
        Exit Sub
    End If
    
ClearSignError:
    If Err = 32755 Then
        Exit Sub
    End If
    If where = 1 Then
        errStr$ = "Couldn't create input file."
    End If
    If where = 2 Then
        errStr$ = "Error in shelling out to PGP."
    End If
    If where = 3 Then
        errStr$ = "Couldn't create output file."
    End If
    If where = 4 Then
        errStr$ = "Error in writing to output file."
    End If
    If where = 5 Then
        errStr$ = "Error in deleting temp file."
    End If
    MsgBox Err.Description + "PGPClearSign. " + errStr$, 48, PiStr
    Exit Sub
    Resume Next
End Sub

Private Sub PGPVerify_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1997 BC
'Author:        BC
'Description:   Helps to verify integrity of PGP.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Created.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
        
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim pp As String
    Dim pix As String
    
    On Error GoTo PGPVerifyError
    
    If Len(Text5.Text) > 0 Then
        '---------------------------------------------
        'test first to see if message area is empty
        '---------------------------------------------
        temp = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Update Remailer")
        If temp = vbYes Then
            Text5.Text = ""
        Else
            Exit Sub
        End If
    End If
    
    Text5.SelText = "Running 3-Step PGP verification." + CRLF + CRLF
    Text5.SelText = "This will check the signature on your pgp262i.zip file" + CRLF
    Text5.SelText = "to see if it was signed by Jeffrey Schiller at MIT." + CRLF
    
    temp1 = Environ$("PGPPATH") + "\PGP262I.ASC"
    temp2 = Environ$("PGPPATH") + "\PGP262I.ZIP"
    temp3 = App.Path + "\JKEY.ASC"
    temp4 = Environ$("PGPPATH") + "pubring.pgp"
    
    If iFileExists(temp1) And iFileExists(temp2) And iFileExists(temp3) Then
    
        '---------------------------------------------
        'set the PIF files for this encryption process
        '---------------------------------------------
        If Form1.PGPClose.Checked = True Then
            PIPIF = "pgpcmdx.pif"
        Else
            PIPIF = "pgpcmd.pif"
        End If

        pp = piPGPPath
        
        'step 1:  run pgp -ka +batchmode jkey.asc command
        FileNum = FreeFile
        Open piPGPPath + "\" + "pgpcmd.bat" For Output As FileNum
        cmd$ = "pgp -ka +batchmode " + App.Path + "\jkey.asc"
        Print #FileNum, cmd$
        Close #FileNum
        #If win16 Then
            Result% = Shell(App.Path + "\" + PIPIF, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        
        'step2:  run pgp pgp262i.asc pgp262i.zip > zcghtemp command
        FileNum = FreeFile
        Open piPGPPath + "\" + "pgpcmd.bat" For Output As FileNum
        cmd$ = "pgp " + temp1 + " " + temp2 + "  > " + piPGPPath + "\zcghtemp"
        Print #FileNum, cmd$
        Close #FileNum
        #If win16 Then
            Result% = Shell(App.Path + "\" + PIPIF, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
        
        'step3:  run pgp -kvc "Jeffrey I. Schiller" >> zcghtemp command
        FileNum = FreeFile
        Open piPGPPath + "\" + "pgpcmd.bat" For Output As FileNum
        cmd$ = "pgp -kvc " + Chr$(34) + "Jeffrey I. Schiller" + Chr$(34) + " " + pp + "\pubring.pgp >> " + piPGPPath + "\zcghtemp"
        Print #FileNum, cmd$
        Close #FileNum
        #If win16 Then
            Result% = Shell(App.Path + "\" + PIPIF, MinState)
            Do While GetModuleUsage(Result%) > 0
                Ecount = DoEvents()
            Loop
        #Else
            ExecCmd (cmd$)
        #End If
        If OS2 = "true" Then
            OS2Pause
        End If
           
        'step4:  process the zcghtemp file
        FileNum = FreeFile
        Open piPGPPath + "\" + "zcghtemp" For Input As FileNum
        cyphertext$ = ""
        Okay = False
        While Not EOF(FileNum)
            Line Input #FileNum, textline$
            If InStr(1, textline$, "DD DC 88 AA 92 DC DD D5  BA 0A 6B 59 C1 65 AD 01", 0) Then
                Text5.Text = ""
                Text5.SelText = "Your pgp262i.zip file has a valid signature on it." + CRLF
                Text5.SelText = "The file was signed by Jeffrey I. Schiller at MIT." + CRLF
                Text5.SelText = "This is not a 100% certification." + CRLF
                Text5.SelText = "Reason:" + CRLF
                Text5.SelText = "1.  A hacker could have tampered with the PGP source code" + CRLF
                Text5.SelText = "    so that the pgp command outputs that this signature" + CRLF
                Text5.SelText = "    is okay, when in reality, it is not Jeffrey's signature." + CRLF
                Text5.SelText = "Possible Defenses:" + CRLF
                Text5.SelText = "1.  Examine the PGP source code and build your own executable (best defense)." + CRLF
                Okay = True
                Exit Sub
            End If
        Wend
        Close FileNum
        If Not Okay Then
            Text5.Text = ""
            Text5.Text = "A problem with your PGP installation has been found" + CRLF
            Text5.SelText = "The signature on the pgp262i.zip file is not valid." + CRLF
            Text5.SelText = "Use the send feedback command to get help with this." + CRLF
        End If
    Else
        Text5.Text = ""
        If Not iFileExists(temp1) Then
            Text5.Text = "You do not have a copy of pgp262i.asc in your pgp directory." + CRLF
        End If
        If Not iFileExists(temp2) Then
            Text5.SelText = "You do not have a copy of pgp262i.zip in your pgp directory." + CRLF
        End If
        If Not iFileExists(temp3) Then
            Text5.SelText = "You do not have a copy of jkey.asc in your Private Idaho directory." + CRLF
        End If
        If Not iFileExists(temp4) Then
            Text5.SelText = "You do not have the pubring.pgp file in your pgp directory." + CRLF
        End If
        Text5.SelText = "PGP verification cannot complete without the required file(s)." + CRLF
    End If
    '---------------------------------------------
    'done, so reset the PIF files
    '---------------------------------------------
    If Form1.PGPClose.Checked = True Then
        PIPIF = "pipgpx.pif"
    Else
        PIPIF = "pipgp.pif"
    End If
Exit Sub
    
PGPVerifyError:
    Resume Next

End Sub

Private Sub PGPWrap_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PGPWrap.Checked = Not PGPWrap.Checked
End Sub

Private Sub PI_Test_Click_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
   ' PI_Test
End Sub

Private Sub POP1_Header(Field As String, Value As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    gMessage = gMessage + Field + ": " + Value + CRLF
End Sub

Private Sub POP1_PITrail(Direction As Integer, Message As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If gSMTPLog = 1 Then
        Print #gSMTPFile, "POP: " + Format$(Direction) + ": " + Message
    End If
End Sub


Private Sub POP1_Transfer(BytesTransferred As Long, Text As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If Len(gMessage) > 62000 Then
        MsgBox ("64k string size exceeded.  Message will be truncated.")
    Else
        gMessage = gMessage + Text & Chr$(13) & Chr$(10)
    End If
End Sub

Private Sub ReamilerAppend_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    AppendInfo
    USENETFi.Checked = False
    UseNetSoda.Checked = False
    USENETGate.Checked = False
    USENETNone.Checked = True
    NewsgroupType = 0
End Sub

Private Sub RemailerKeys_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/26/97
'Description:   Commented out check for text in Text5.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    On Error GoTo RemailerKeysError
    
'    If Len(Text5.Text) > 0 Then
'        '---------------------------------------------
'        'test first to see if message area is empty
'        '---------------------------------------------
'        temp = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Fetch Remailer Keys")
'        If temp = vbYes Then
'            Text5.Text = ""
'        Else
'            Exit Sub
'        End If
'    End If
    
    If CheckConnection Then
        gWebState = GETREMAILERKEYS
        Form1.HTTP1.WinsockLoaded = True
        ShowStatus ("Getting current remailer PGP keys.  Please wait.  Hint: Use 'Add key from message' when done.")
        If gPGPKeysURL = "" Then
            Dim SectionName As String
            SectionName = "Net Info"
            gPGPKeysURL = ReadProfile(SectionName, "PGPKeysURL")
            If gPGPKeysURL = "" Then
                gPGPKeysURL = "http://kiwi.cs.berkeley.edu/pgpkeys"
                WriteProfile SectionName, "PGPKeysURL", gPGPKeysURL
            End If
        End If
        GetWebURL (gPGPKeysURL)
    End If
    Exit Sub
RemailerKeysError:
    HideStatus
    MsgBox Err.Description + "RemailerKeys"
    Exit Sub
End Sub
Private Sub RemailersCP_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu entry for selecting the cypherpunk remailer list.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Fixed a bug where the remailer
'               list alignments are off if you
'               switch from mixmaster to
'               cypherpunk on the remailer menu.
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    '---------------------------------------------
    'restore settings
    '---------------------------------------------
    '---------------------------------------------
    'clear the remailer arrays
    '---------------------------------------------
    ClearRemailerSort
    gSortRemailer.cpunk = 1
    gSortRemailer.eric = 0
    '---------------------------------------------
    'Load the matched remailer array with cypherpunk style
    '---------------------------------------------
    SortRemailersByType
    '---------------------------------------------
    'prepare to load the remailer list
    '---------------------------------------------
    RemailersMix.Checked = 0
    RemailersPenet.Checked = 0
    RemailersSoda.Checked = 0
    '---------------------------------------------
    'load the remailer list
    '---------------------------------------------
    FillRemailerList
    RemailersCP.Checked = True
    RemailerType = Cypherpunk
    USENETGate.Visible = True
    USENETFi.Visible = False
    UseNetSoda.Visible = False
    USENETNone.Checked = True
    NewsgroupType = 0
    Form1.Label6.Caption = "Remailer update as of " + Format$(FileDateTime(App.Path + "\remailer.htm"), "mm/dd/yy hh:mm")
    If RemailersCP.Checked = True Then
        OptionsAdv.Enabled = True
    End If
End Sub

Private Sub RemailersCut_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    RemailersCut.Checked = Not RemailersCut.Checked
    If RemailersCut.Checked Then
        gSortRemailer.cut = 1
    Else
        gSortRemailer.cut = 0
    End If
    SortRemailers
    FillRemailerList
End Sub

Private Sub RemailersHash_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    RemailersHash.Checked = Not RemailersHash.Checked
    If RemailersHash.Checked Then
        gSortRemailer.hash = 1
    Else
        gSortRemailer.hash = 0
    End If
    SortRemailers
    FillRemailerList
End Sub

Private Sub RemailersLatent_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    RemailersLatent.Checked = Not RemailersLatent.Checked
    If RemailersLatent.Checked Then
        gSortRemailer.latent = 1
    Else
        gSortRemailer.latent = 0
    End If
    SortRemailers
    FillRemailerList
End Sub

Private Sub RemailersMix_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    RemailersSoda.Checked = 0
    RemailersCP.Checked = 0
    RemailersPenet.Checked = 0
    RemailersMix.Checked = True
    RemailerType = mix
    If RemailersMix.Checked = True Then
        OptionsAdv.Enabled = False
    End If
    ClearRemailerSort
    gSortRemailer.eric = 0
    gSortRemailer.cpunk = 0
    gSortRemailer.mix = 1
    InitializeMixRemailers (App.Path + "\mixmaste.htm")
    FillRemailerList
End Sub

Private Sub RemailersPenet_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    ClearRemailerSort
    gSortRemailer.penet = 1
    SortRemailersByType
    FillRemailerList
    RemailersCP.Checked = 0
    RemailersSoda.Checked = 0
    RemailersMix.Checked = 0
    RemailersPenet.Checked = True
    RemailerType = penet
    USENETGate.Visible = False
    UseNetSoda.Visible = False
    USENETFi.Visible = True
    USENETNone.Checked = True
    NewsgroupType = 0
    If RemailersPenet.Checked = True Then
        OptionsAdv.Enabled = False
    End If
    Form1.Label6.Caption = "Remailer update as of " + Format$(FileDateTime(App.Path + "\remailer.htm"), "mm/dd/yy hh:mm")
End Sub

Private Sub RemailersPGP_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu handler for encrypting to remailers.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================
'=============================================
'Module Level Variable Declaration Section
'=============================================

    RemailersPGP.Checked = Not RemailersPGP.Checked
    If RemailersPGP.Checked Then
        gSortRemailer.encrypt = 1
    Else
        gSortRemailer.encrypt = 0
    End If
    SortRemailers
    FillRemailerList
End Sub

Private Sub RemailersPost_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    RemailersPost.Checked = Not RemailersPost.Checked
    If RemailersPost.Checked Then
        gSortRemailer.post = 1
    Else
        gSortRemailer.post = 0
    End If
    SortRemailers
    FillRemailerList
End Sub

Private Sub RemailersSoda_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    ClearRemailerSort
    gSortRemailer.eric = 1
    SortRemailersByType
    FillRemailerList
    RemailersCP.Checked = 0
    RemailersPenet.Checked = 0
    RemailersMix.Checked = 0
    RemailersSoda.Checked = True
    RemailerType = Soda
    If RemailersSoda.Checked = True Then
        OptionsAdv.Enabled = False
    End If
    Form1.Label6.Caption = "Remailer update as of " + Format$(FileDateTime(App.Path + "\remailer.htm"), "mm/dd/yy hh:mm")
End Sub

Private Sub RemailerUpdate_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   Menu entry for downloading new remailer data.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added option to clear text area.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/12/97
'Description:   Added CheckConnection.
'---------------------------------------------
'Developer:     BC
'Change Date:   05/26/97
'Description:   Commented out check for text in Text5.
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    
    On Error GoTo RemailerError
    
'    If Len(Text5.Text) > 0 Then
'        '---------------------------------------------
'        'test first to see if message area is empty
'        '---------------------------------------------
'        temp = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Update Remailer")
'        If temp = vbYes Then
'            Text5.Text = ""
'        Else
'            Exit Sub
'        End If
'    End If
    
    If CheckConnection Then
        gServerState = HTTPSTATE
        '---------------------------------------------
        'fetch the URL for obtaining remailer data
        '---------------------------------------------
        If gRemailerURL = "" Then
            Dim SectionName As String
            SectionName = "Net Info"
            gMixListURL = ReadProfile(SectionName, "MixListURL")
            If Len(gMixListURL) = 0 Then
                gMixListURL = "http://kiwi.cs.berkeley.edu/mixmaster-list.html"
                WriteProfile SectionName, "MixListURL", gMixListURL
            End If
            gMixType2URL = ReadProfile(SectionName, "MixType2URL")
            If Len(gMixType2URL) = 0 Then
                gMixType2URL = "http://kiwi.cs.berkeley.edu/type2.list"
                WriteProfile SectionName, "MixType2URL", gMixType2URL
            End If
            gMixPubRingURL = ReadProfile(SectionName, "MixPubRingURL")
            If Len(gMixPubRingURL) = 0 Then
                gMixPubRingURL = "http://kiwi.cs.berkeley.edu/pubring.mix"
                WriteProfile SectionName, "MixPubRingURL", gMixPubRingURL
            End If
            gRemailerURL = ReadProfile(SectionName, "RemailerURL")
            If Len(gRemailerURL) = 0 Then
                gRemailerURL = "http://www.cs.berkeley.edu/~raph/remailer-list.html"
                WriteProfile SectionName, "RemailerURL", gRemailerURL
            End If
        End If
        '---------------------------------------------
        'mixmaster option selected on menu
        '---------------------------------------------
        If RemailersMix.Checked Then
            gWebState = MIXUPDATE
            Form1.HTTP1.WinsockLoaded = True
            ShowStatus ("Downloading and updating Mixmaster remailer information from Raph's pinging service.  Please wait.")
            GetWebURL (gMixListURL)
            '---------------------------------------------
            'stall, then get the type2
            '---------------------------------------------
            Do Until Form1.HTTP1.WinsockLoaded = False
            Loop
            gWebState = TYPE2UPDATE
            Form1.HTTP1.WinsockLoaded = True
            ShowStatus ("Downloading and updating the Mixmaster Type2.lis file from Raph's pinging service.  Please wait.")
            GetWebURL (gMixType2URL)
            '---------------------------------------------
            'stall, then get the pubring
            '---------------------------------------------
            Do Until Form1.HTTP1.WinsockLoaded = False
            Loop
            gWebState = PUBRINGUPDATE
            Form1.HTTP1.WinsockLoaded = True
            ShowStatus ("Downloading and updating the Mixmaster Pubring.mix file from Raph's pinging service.  Please wait.")
            GetWebURL (gMixPubRingURL)
        Else
            '---------------------------------------------
            'mixmaster is not checked on the menu
            '---------------------------------------------
            gWebState = GETREMAILERUPDATE
            '---------------------------------------------
            'run a winsock
            '---------------------------------------------
            Form1.HTTP1.WinsockLoaded = True
            '---------------------------------------------
            'put up the remailer downloading status dialog
            '---------------------------------------------
            ShowStatus ("Downloading and updating remailer information from Raph's pinging service.  Please wait.")
            '---------------------------------------------
            'run a winsock
            '---------------------------------------------
            GetWebURL (gRemailerURL)
        End If
        '---------------------------------------------
        'finally, do an update as though we are restarting
        'this is a temporary work around for a bug in the
        'update remailer info structure @@
        '---------------------------------------------
        InitializeRemailers (App.Path + "\remailer.htm")
        '---------------------------------------------
        'set the newsgroup menu for cp state
        '---------------------------------------------
        USENETGate.Visible = True
        USENETFi.Visible = False
        UseNetSoda.Visible = False
        RestoreSettings
    End If
    Exit Sub
RemailerError:
    HideStatus
    MsgBox Err.Description + "RemailerUpdate"
    gServerState = 0
    Exit Sub
End Sub


Private Sub SendSysInfo_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    PrepareFeedback
End Sub




Private Sub SMTP1_Error(ErrorCode As Integer, Description As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form1.Label13.Caption = "Error" + Str$(ErrorCode) + ". Mail not sent."
End Sub


Private Sub SMTP1_PITrail(Direction As Integer, Message As String)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form1.Label13.Caption = Message
    If gSMTPLog = 1 Then
        Print #gSMTPFile, "SMTP: " + Format$(Direction) + ": " + Message
    End If
End Sub
Private Sub SMTP1_StartTransfer()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form1.Label13.Caption = "Sending mail."
End Sub

Private Sub SMTP1_Transfer(BytesTransfered As Long)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form33.Label3.Caption = Format$(BytesTransfered)
End Sub

Private Sub StepAddKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Adding a public key"
    Form32.Text1.Text = "Before you can send someone an encrypted message, you need a copy of their public key.  To add a copy of your key:" + CRLF + CRLF + "1. - Copy the public key (either from an e-mail message or key server) and paste it into the Message text box." + CRLF + CRLF + "2. - From the Keys menu, choose the 'Add key from message' command." + CRLF + CRLF + "3. - PGP will run in the DOS window.  If you are running Windows 95, click the icon in the taskbar.  Certify the key." + CRLF + CRLF + "The key is added to your public key ring."
    Form32.Show
End Sub

Private Sub StepAttach_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Sending an attachment"
    Form32.Text1.Text = "You can attach a file to a message sent from Private Idaho.  To send a message with an attachment:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + CRLF + CRLF + "3. - Make sure you have a connection to the Internet if you are using the 16-bit version." + CRLF + CRLF + "4. - Check the Attachment checkbox and specify the file to attach." + CRLF + CRLF + "5. - In the drop-down box, specify if you'd like the file encrypted, you may also want to encrypt the text message before you send it." + CRLF + CRLF + "6. - Select the Send button (or choose 'Send' from the Message menu)." + CRLF + CRLF + "Note: The file is Base64 encoded to a MIME compliant attachment." + CRLF + CRLF + "Private Idaho currently doesn't support sending attachments through remailers."
    Form32.Show
End Sub

Private Sub StepCreateKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Creating a PGP key pair"
    Form32.Text1.Text = "If you'd like to create a PGP secret and public key to use with a nym:" + CRLF + CRLF + "1. - From the Keys menu, choose the 'Create key pair' command." + CRLF + CRLF + "2. - PGP will run in the DOS window.  Follow the steps for creating a key." + CRLF + CRLF + "Hint: Use a key size of 1024 bits or higher."
    Form32.Show
End Sub

Private Sub StepDecrypt_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Decrypting a message"
    Form32.Text1.Text = "To decrypt a PGP message you've received:" + CRLF + CRLF + "1. - Copy the message from your e-mail application." + CRLF + CRLF + "2. - Paste it into Private Idaho's Message text box." + CRLF + CRLF + "3. - In the PGP menu, choose the 'Decrypt message' command." + CRLF + CRLF + "4. - PGP will run in the DOS window.  Enter your PGP passphrase." + CRLF + CRLF + "The encrypted text in the Message box is replaced by the decrypted text."
    Form32.Show
End Sub

Private Sub StepDelete_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Deleting a public key"
    Form32.Text1.Text = "To remove a key from your public key ring:" + CRLF + CRLF + "1. - From the Keys menu, choose the 'Delete key' command." + CRLF + CRLF + "2. - Select the key to remove and click OK." + CRLF + CRLF + "3. - PGP will run in the DOS window.  If you are running Windows 95, click the icon in the taskbar.  Verify you want to remove the key." + CRLF + CRLF + "The key is removed from your public key ring."
    Form32.Show
End Sub

Private Sub StepEncrypt_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Encrypting a message"
    Form32.Text1.Text = "To PGP encrypt a message:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - In the To: text box, enter the PGP user ID of person you'll be sending the message to.  This is usually an e-mail address.  You must have already added the person's key to your public key ring." + CRLF + CRLF + "Hint: You can use the Address book in the File menu to store common user IDs and addresses.  You can access these by clicking on the button on the right side of the To: box.  Another way is to click to the left of the To: field where it says Click Here.  This will bring up a copy of the recipients on your public key ring." + CRLF + CRLF + "3. - From the PGP menu, choose the 'Encrypt message' command." + CRLF + CRLF + "PGP will run in the DOS window.  If you are running Windows 95, you will need to double click on the PGP icon in the task bar." + CRLF + CRLF + "The original text in the Message box is replaced by the encrypted text."
    Form32.Show
End Sub

Private Sub StepGetKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Sending/getting MIT server keys"
    Form32.Text1.Text = "MIT has an experimental PGP key server where you can publish and retrieve keys.  Private Idaho can directly access the server without going to the Web page.  To request a copy of someone's key:" + CRLF + CRLF + "1. - Make sure you have an Internet connection (only applies to the 16-bit version)." + CRLF + CRLF + "2. - In the To: text box, enter the e-mail address of the person's key you want to retrieve." + CRLF + CRLF + "3. - From the Keys menu, choose the 'Get key from server' command." + CRLF + CRLF + "4. - Private Idaho will connect with the server.  If the key is in the data base, it will be displayed in the Message text box.  You can then add the key to your public key ring with the 'Add key from message' command." + CRLF + CRLF + "To add a key to the server, use the 'Insert key in message' command in the Keys menu to insert your key.  Then choose the 'Submit key to server' command."
    Form32.Show
End Sub

Private Sub StepInfo_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Internet privacy info"
    Form32.Text1.Text = "Private Idaho has links to a variety of Internet privacy sources.  To access them:" + CRLF + CRLF + "1. - Ensure Private Idaho can communicate with your Web browser.  The default is Netscape Navigator.  If you're using another browser, choose the Options command in the Web menu." + CRLF + CRLF + "2. - You should be connected to the Internet with the browser running and not minimized." + CRLF + CRLF + "3. - From the Web menu, choose the information you'd like to access." + CRLF + CRLF + "Your browser will display the associated Web page."
    Form32.Show
End Sub

Private Sub StepNym_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Creating a nym"
    Form32.Text1.Text = "A nym (as in ano'nym'ous) is an alias used for private communications.  Once you've created a nym account, you can send messages through it to people.  Unlike anonymous remailers, they can reply back to you without knowing your identity.  Various free servers are available for setting up nym accounts.  These are much more secure than using anon.penet.fi.  To create a nym:" + CRLF + CRLF + "1. - From the Nym menu, choose the 'Create nym' command." + CRLF + CRLF + "This command steps you through the entire nym creation process with a series of easy to follow dialog boxes." + CRLF + CRLF + "Refer to the on-line help for additional information."
    Form32.Show
End Sub

Private Sub StepNymDelete_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Deleting a nym"
    Form32.Text1.Text = "To delete a nym:" + CRLF + CRLF + "1. - From the Nym menu, choose the 'Delete nym' command." + CRLF + CRLF + "2. - Select the nym account to delete and click OK." + CRLF + CRLF + "3. - Send the prepared message to the nym server.  Your nym will be deleted."
    Form32.Show
End Sub

Private Sub StepNymPass_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Changing a nym password"
    Form32.Text1.Text = "If you want to change your nym password:" + CRLF + CRLF + "1. - From the Nym menu, choose the 'Change nym password' command." + CRLF + CRLF + "2. - Select the nym to change and click OK." + CRLF + CRLF + "3. - In the Message text box enter your old and new passwords." + CRLF + CRLF + "4. - From the Nym menu, choose the 'Encrypt nym message' command." + CRLF + CRLF + "5. Send the message."
    Form32.Show
End Sub

Private Sub StepNymReply_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Changing a nym reply block"
    Form32.Text1.Text = "If you want to change the routing of your nym messages:" + CRLF + CRLF + "1. - Enter the final destination e-mail address in the To: text box." + CRLF + CRLF + "2. - Specify the new remailer or chain in the Remailer drop-down list." + CRLF + CRLF + "3. - From the Nym menu, choose the 'Change reply block' command." + CRLF + CRLF + "4. - Select the nym you want to change." + CRLF + CRLF + "5. - Enter your nym password in the Message text box after Password:.  Make sure the reply block is correct." + CRLF + CRLF + "6. - For alias type nyms, from the Nym menu, choose the 'Encrypt nym message' command." + CRLF + CRLF + "7. - Send the message."
    Form32.Show
End Sub

Private Sub StepNymSend_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Sending a nym message"
    Form32.Text1.Text = "Once you've created a nym account, you can send messages through it.  To do so:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + CRLF + CRLF + "3. - From the Nym menu, choose the 'Prepare nym message' command." + CRLF + CRLF + "4. - Select the nym account to use and click OK." + CRLF + CRLF + "5. - For alias type nyms, from the Nym menu, choose the 'Encrypt nym message' command." + CRLF + CRLF + "7. Send the message."
    Form32.Show
End Sub

Private Sub StepRemailer_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Anonymous messages"
    Form32.Text1.Text = "You can send some a message without revealing your identity by using an anonymous remailer.  To send an anonymous message:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + CRLF + CRLF + "3. - From the Remailers menu, choose the type of remailer to use." + CRLF + CRLF + "4. - Select the remailer to send the message through from the Remailer drop-down list.  Selecting 'chain' routes the message through several remailers." + CRLF + CRLF + "5. - From the Message menu, choose the 'Append info' command.  This formats the message for sending through a remailer.  If you selected 'chain,' a dialog box will prompt you for the remailers to use." + CRLF + CRLF + "6. - Send the message." + CRLF + CRLF + "Note: The Cypherpunk type remailers support a variety of advanced features.  Refer to the on-line help."
    Form32.Show
End Sub

Private Sub StepSend_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Sending a message from Private Idaho"
    Form32.Text1.Text = "If you have a connection to the Internet you can send a message directly from Private Idaho.  To send a message:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + CRLF + CRLF + "3. - Make sure you have a connection to the Internet." + CRLF + CRLF + "4. - Select the Send button (or choose 'Send' from the Message menu)." + CRLF + CRLF + "Note: Before sending a message, you need to provide information about your e-mail server.  From the File menu, choose the 'Options' command." + CRLF + CRLF + "If you can't send e-mail directly from Private Idaho, you can easily transfer the message to an e-mail application such as Eudora or Pegasus.  Refer to the on-line help."
    Form32.Show
End Sub

Private Sub StepSendKey_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Sending your public key"
    Form32.Text1.Text = "Before someone can send you an encrypted message, they need a copy of your public key.  To send a copy of your key:" + CRLF + CRLF + "1. - From the Keys menu, choose the 'Insert key in message' command." + CRLF + CRLF + "2. - Select your key from the user ID dialog box and click OK." + CRLF + CRLF + "3. - PGP will run and fetch the key from your public key ring." + CRLF + CRLF + "4. - The public key is inserted in the Message text box.  You can now send the key to someone you want to privately correspond with."
    Form32.Show
End Sub

Private Sub StepSign_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Signing a message"
    Form32.Text1.Text = "There are two ways to sign a message.  One way is to clear-sign a message.  This method leaves the text alone, but wraps it with a signature.  Signing the message must be the last step as the signature depends on the contents of the message.  The second way is to sign an encrypted file.  Either way, the intent is to let the recipient know that you are the one who created the message because it requires your secret key and password to compute the signature." + CRLF + CRLF + "To sign a message:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - From the PGP menu, choose the 'Clear sign message' or 'Encrypt and Sign' command." + CRLF + CRLF + "3. - PGP will run in the DOS window.  Enter your passphrase." + CRLF + CRLF + "A signature is attached to the message."
    Form32.Show
End Sub

Private Sub StepUpdateInfo_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Updating remailer info"
    Form32.Text1.Text = "To update remailer info:" + CRLF + CRLF + "1. - Make sure you have a Net connection." + CRLF + CRLF + "2. - From the Remailers menu, choose the 'Update remailer info' command." + CRLF + CRLF + "Private Idaho will connect to Raph's Web page and download the current remailer data and update the remailer list box.  If Cypherpunk is checked in the Remailer menu, Cypherpunk info is updated.  If Mixmaster is checked, Mixmaster info is updated."
    Form32.Show
End Sub

Private Sub StepUSENET_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "USENET articles"
    zap$ = "You can post a USENET article without revealing your identity by using a remailer.  To post an anonymous article:" + CRLF + CRLF + "1. - Compose the message in the Message text box." + CRLF + CRLF + "2. - In the To: text box, enter the name of the USENET newsgroup (i.e. alt.security.pgp)." + CRLF + CRLF + "3. - From the Remailers menu, choose 'Cypherpunks.'  Then choose 'Post to USENET' in the 'Advanced Cypherpunk' menu item." + CRLF + CRLF + "4. - Select the remailer to send the message through from the Remailer drop-down list." + CRLF + CRLF + "5. - From the Message menu, choose the 'Append info' command.  This formats the message for sending through the remailer." + CRLF + CRLF + "6. - Send the message." + CRLF + CRLF + "Note: Only a few remailers support direct posting to USENET newsgroups.  You can also use e-mail gateways to send a message through any remailer.  To do so, enter the newsgroup on the To: line, then from the Newsgroup menu choose the 'E-mail gateway'"
    Zip$ = " command and select the gateway you'd like to use.  Select the remailer(s) to route the message through." + CRLF + CRLF + "Posting anonymous USENET articles is not as reliable as sending anonymous e-mail.  Sometimes messages do not get posted.  Refer to the on-line help for additional information."
    Form32.Text1.Text = zap$ + Zip$
    Form32.Show
End Sub

Private Sub StepWeb_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form32.Caption = "Anonymous Web page access"
    Form32.Text1.Text = "It's very easy for someone to log Web pages you visit.  Community Connexion (c2.org), has a free Web anonymizer service.  To use it from Private Idaho:" + CRLF + CRLF + "1. - Type the Web page URL you want to anonymously visit in the Message text box and select (highlight) it." + CRLF + CRLF + "2. Ensure Private Idaho can communicate with your Web browser.  The default is Netscape Navigator.  If you're using another browser, choose the Options command in the Web menu." + CRLF + CRLF + "3. - You should be connected to the Internet with the browser running and not minimized." + CRLF + CRLF + "4. - From the Web menu, choose the 'Anonymous jump to URL' command." + CRLF + CRLF + "Your browser will anonymously access the Web page." + CRLF + CRLF + "Hint: Enter frequently accessed Web pages in Private Idaho's address book."
    Form32.Show
End Sub


Private Sub TransferAES_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    TransferAES.Checked = Not TransferAES.Checked
End Sub

Private Sub TransferApp1_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App1Wind")
    AppScript = ReadProfile(SectionName, "App1Script")
    TransferInfo WindApp$, AppScript
    'normal call
    '    AppActivate AppWind
    'added for NT support
    '    x = FindWindow(0, AppWind)
    '    SetWindowPos x, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    '    SendScript (AppScript)
End Sub

Private Sub TransferApp2_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App2Wind")
    AppScript = ReadProfile(SectionName, "App2Script")
    TransferInfo WindApp$, AppScript
    'normal call
    '  AppActivate AppWind
    'added for NT support
    '  x = FindWindow(0, AppWind)
    '  SetWindowPos x, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    '  SendScript (AppScript)
End Sub

Private Sub TransferApp3_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App3Wind")
    AppScript = ReadProfile(SectionName, "App3Script")
    TransferInfo WindApp$, AppScript
    'normal call
    '  AppActivate AppWind
    'added for NT support
    '  x = FindWindow(0, AppWind)
    '  SetWindowPos x, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    '  SendScript (AppScript)
End Sub

Private Sub TransferApp4_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App4Wind")
    AppScript = ReadProfile(SectionName, "App4Script")
    TransferInfo WindApp$, AppScript
    'normal call
    '   AppActivate AppWind
    'added for NT support
    '   x = FindWindow(0, AppWind)
    '   SetWindowPos x, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    '   SendScript (AppScript)
End Sub

Private Sub TransferAS_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    AppendInfo
End Sub

Private Sub TransferAT_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    AppendInfo
    TransferInfo gEmailer, tranScript
End Sub

Private Sub TransferEncrypt_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    foo$ = Combo2.Text
    EncryptMessage foo$, "-eatw"
End Sub

Private Sub TransferEu_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    TransferInfo gEmailer, tranScript
End Sub

Private Sub TransferNym_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form4.Show
End Sub

Private Sub TransferOptions_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    SetEmailers
    Form5.Show 1
End Sub

Private Sub TransferPrepare_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    gNymState = NYMPREPARE
    Form27.Show 1
    gNymState = NYMIDLE
End Sub

Private Sub TransferReply_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If Form1.Text5.Text <> "" Then
        fie$ = ">"
        foo$ = InsertCRLFs()
        For x = 1 To Len(foo$)
            tmpChar$ = Mid$(foo$, x, 1)
            fie$ = fie$ + tmpChar$
            If tmpChar$ = Chr$(10) Then
                fie$ = fie$ + ">"
            End If
        Next
        Form1.Text5.Text = fie$
        Combo2.Text = gMessageRecord.Sender
    End If
End Sub

Private Sub TransferSend_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If TransferAES.Checked Then
        AppendInfo
    End If
    If CheckConnection Then
        SendMailMessage
    End If
End Sub

Private Sub TransferXHeaders_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    TransferXHeaders.Checked = Not TransferXHeaders.Checked
    If TransferXHeaders.Checked Then
        Form34.Show
    End If
End Sub

Private Sub USENETFi_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    USENETNone.Checked = False
    USENETGate.Checked = False
    UseNetSoda.Checked = False
    USENETFi.Checked = True
    NewsgroupType = penet
End Sub

Private Sub USENETGate_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If InStr(1, Combo2.Text, "@") <> 0 Or Combo2.Text = "" Then
        MsgBox "The To: line must contain a valid newsgroup name (such as alt.security.pgp), to send through an e-mail to USENET gateway."
        USENETGate.Checked = False
    Else
        USENETNone.Checked = False
        USENETFi.Checked = False
        UseNetSoda.Checked = False
        USENETGate.Checked = True
        NewsgroupType = USENET
        Form9.Show 1
    End If
End Sub

Private Sub USENETheader_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    USENETheader.Checked = Not USENETheader.Checked
End Sub

Private Sub USENETNone_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    USENETFi.Checked = False
    USENETGate.Checked = False
    UseNetSoda.Checked = False
    USENETNone.Checked = True
    NewsgroupType = 0
End Sub

Private Sub UseNetSoda_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    USENETNone.Checked = False
    USENETGate.Checked = False
    USENETFi.Checked = False
    UseNetSoda.Checked = True
    NewsgroupType = Soda
End Sub

Private Sub Web1_Click(Index As Integer)
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If CheckConnection Then
        If iFileExists(App.Path + "\URLS.TXT") Then
            FileNum = FreeFile
            Open App.Path + "\URLS.TXT" For Input As FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, Item$
                If InStr(1, Item$, Web1(Index).Caption) Then
                    GetURL (Mid$(Item$, InStr(1, Item$, ",") + 1, Len(Item$) - InStr(1, Item$, ",")))
                    Close #FileNum
                    Exit Sub
                End If
            Loop
            Close #FileNum
        Else
            MsgBox "URLS.TXT not present in the same directory as " + PiStr + ".  URLs won't be shown.", 48, PiStr
        End If
    End If
End Sub

Private Sub WebJump_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    If CheckConnection Then
        AnonURLJump
    End If
End Sub

Private Sub WebOptions_Click()
'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/14/97
'Description:   Added comments.
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================
    Form15.Show 1
End Sub
Private Sub EditSetFont_Click()

'Workfile:      EUPGP.FRM (Form1)
'Copyright:     Copyright © 1997 BC
'Author:        BC
'Description:   Sets font styles for text and combo boxes on Form1.
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   04/18/97
'Description:   Added comments.
    
'=============================================
'Module Level Constant Declaration Section
'=============================================

'=============================================
'Module Level Variable Declaration Section
'=============================================

    '---------------------------------------------
    'display the common font dialog
    '---------------------------------------------
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    
    '---------------------------------------------
    'set the font for the message area
    '---------------------------------------------
    Text5.FontName = CommonDialog1.FontName
    Text5.FontBold = CommonDialog1.FontBold
    Text5.FontItalic = CommonDialog1.FontItalic
    Text5.FontStrikethru = CommonDialog1.FontStrikethru
    Text5.FontSize = CommonDialog1.FontSize
    
    '---------------------------------------------
    'set the font for the to: box
    '---------------------------------------------
    Combo2.FontName = CommonDialog1.FontName
    Combo2.FontBold = CommonDialog1.FontBold
    Combo2.FontItalic = CommonDialog1.FontItalic
    Combo2.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the subject: box
    '---------------------------------------------
    Text2.FontName = CommonDialog1.FontName
    Text2.FontBold = CommonDialog1.FontBold
    Text2.FontItalic = CommonDialog1.FontItalic
    Text2.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the cc: box
    '---------------------------------------------
    Text3.FontName = CommonDialog1.FontName
    Text3.FontBold = CommonDialog1.FontBold
    Text3.FontItalic = CommonDialog1.FontItalic
    Text3.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the bcc: box
    '---------------------------------------------
    Text4.FontName = CommonDialog1.FontName
    Text4.FontBold = CommonDialog1.FontBold
    Text4.FontItalic = CommonDialog1.FontItalic
    Text4.FontStrikethru = CommonDialog1.FontStrikethru

End Sub
