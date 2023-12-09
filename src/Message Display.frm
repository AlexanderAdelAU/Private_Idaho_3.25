VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "Threed20.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{33337113-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ipport40.ocx"
Object = "{33337143-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "netcod40.ocx"
Object = "{33337153-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ipinfo40.ocx"
Object = "{33337183-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "http40.ocx"
Object = "{33337233-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "smtp40.ocx"
Object = "{33337253-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "nntp40.ocx"
Object = "{33337283-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "mime40.ocx"
Object = "{C3153B59-A95E-11D1-A5B6-006097DC9787}#2.0#0"; "FILEIC~1.OCX"
Begin VB.Form frmMessageDisplay 
   AutoRedraw      =   -1  'True
   Caption         =   "Private Idaho 32  Version 4"
   ClientHeight    =   7095
   ClientLeft      =   405
   ClientTop       =   825
   ClientWidth     =   11340
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
   Icon            =   "Message Display.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   11340
   Begin VB.CommandButton btnTo 
      Caption         =   "Cc..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   23
      Top             =   900
      Width           =   555
   End
   Begin VB.CommandButton btnTo 
      Caption         =   "To..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   22
      Top             =   570
      Width           =   555
   End
   Begin VB.TextBox txtCC 
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
      Left            =   750
      TabIndex        =   21
      Top             =   870
      Width           =   8145
   End
   Begin VB.TextBox txtSubject 
      Height          =   300
      Left            =   750
      TabIndex        =   20
      Top             =   1200
      Width           =   8205
   End
   Begin VB.TextBox txtTo 
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
      Left            =   750
      TabIndex        =   19
      Top             =   510
      Width           =   8145
   End
   Begin FileIconImageList.FileIconsImageList FileIconsImageList1 
      Height          =   405
      Left            =   10620
      TabIndex        =   18
      Top             =   840
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   714
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   6660
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   767
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6703
            MinWidth        =   6703
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Name of Attachment"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Message Display.frx":030A
            Text            =   "Attachment"
            TextSave        =   "Attachment"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10500
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "CommonDialog1"
      FontName        =   "Arial"
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   820
      _Version        =   131074
      PictureBackgroundStyle=   2
      PictureBackground=   "Message Display.frx":041C
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   345
         Left            =   10260
         TabIndex        =   17
         Top             =   60
         Width           =   525
      End
      Begin VB.ComboBox cmbRemailerSelect 
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
         ItemData        =   "Message Display.frx":10D46
         Left            =   5610
         List            =   "Message Display.frx":10D48
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   60
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Send Options: "
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
         TabIndex        =   15
         Top             =   120
         Width           =   1035
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   20
         Left            =   8880
         TabIndex        =   13
         ToolTipText     =   "Send the message"
         Top             =   0
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":10D4A
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   21
         Left            =   9480
         TabIndex        =   12
         ToolTipText     =   "Reply to sender"
         Top             =   26
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":11250
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   10
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Reply to sender"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":113B8
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   9
         Left            =   3300
         TabIndex        =   10
         ToolTipText     =   "Encrypt the message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":114F6
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   7
         Left            =   2940
         TabIndex        =   9
         ToolTipText     =   "Add an attachment"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":11642
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   6
         Left            =   2580
         TabIndex        =   8
         ToolTipText     =   "Prepare Remailer Message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":1178A
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Send the message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":118B2
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   5
         Left            =   330
         TabIndex        =   6
         ToolTipText     =   "Save message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":11DB8
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   4
         Left            =   1110
         TabIndex        =   5
         ToolTipText     =   "Open Message"
         Top             =   45
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":1210A
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   3
         Left            =   780
         TabIndex        =   4
         ToolTipText     =   "New Message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":1245C
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         ToolTipText     =   "Paste"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":127AE
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   1
         Left            =   1830
         TabIndex        =   2
         ToolTipText     =   "Copy"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":12B00
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   8
         Left            =   1500
         TabIndex        =   1
         ToolTipText     =   "Cut"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "Message Display.frx":12E52
         ButtonStyle     =   3
      End
   End
   Begin RichTextLib.RichTextBox MessageArea 
      DragIcon        =   "Message Display.frx":131A4
      Height          =   1980
      Left            =   30
      TabIndex        =   24
      Top             =   1770
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3493
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"Message Display.frx":134AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView lvwAttachments 
      Height          =   975
      Left            =   30
      TabIndex        =   25
      Top             =   4530
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1720
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subject:"
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
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Top             =   1260
      Width           =   660
   End
   Begin MIMELibCtl.MIME MIME1 
      Left            =   10590
      Top             =   1440
      Boundary        =   ""
      ContentType     =   ""
      ContentTypeAttr =   ""
      Message         =   ""
      MessageHeaders  =   ""
   End
   Begin NNTPLibCtl.NNTP NNTP1 
      Left            =   10620
      Top             =   4920
      CurrentArticle  =   ""
      CurrentGroup    =   ""
      NewsServer      =   ""
      Password        =   ""
      User            =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin HTTPLibCtl.HTTP HTTP1 
      Left            =   10620
      Top             =   4320
      Accept          =   ""
      LocalFile       =   ""
      Password        =   ""
      ProxyPort       =   80
      ProxyServer     =   ""
      URL             =   ""
      User            =   ""
      UserAgent       =   "devSoft's HTTP Control"
      WinsockLoaded   =   -1  'True
   End
   Begin IPINFOLibCtl.IPInfo IPInfo1 
      Left            =   10620
      Top             =   3720
      PendingRequests =   15
      ServiceName     =   ""
      ServicePort     =   0
      ServiceProtocol =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin SMTPLibCtl.SMTP SMTP1 
      Left            =   10500
      Top             =   3120
      BCc             =   ""
      Cc              =   ""
      Date            =   ""
      From            =   ""
      MailServer      =   ""
      MessageText     =   ""
      ReplyTo         =   ""
      Subject         =   ""
      To              =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin NETCODELibCtl.NetCode NetCode1 
      Left            =   10620
      Top             =   2520
      DecodedData     =   ""
      EncodedData     =   ""
      FileName        =   ""
      Format          =   0
      IntelliCode     =   -1  'True
      MaxFileSize     =   0
      Mode            =   "0755"
      Overwrite       =   0   'False
      ProgressStep    =   1
   End
   Begin IPPORTLibCtl.IPPort IPPort1 
      Left            =   10620
      Top             =   1920
      EOL             =   ""
      InBufferSize    =   2048
      KeepAlive       =   0   'False
      Linger          =   -1  'True
      LocalPort       =   0
      MaxLineLength   =   2048
      OutBufferSize   =   2048
      RemoteHost      =   ""
      RemotePort      =   0
      WinsockLoaded   =   -1  'True
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu mFile_Save 
         Caption         =   "Save"
      End
      Begin VB.Menu FileExport 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu FileNul2 
         Caption         =   "-"
      End
      Begin VB.Menu mEncodeFile 
         Caption         =   "E&ncode a File into Message Area"
      End
      Begin VB.Menu mDecodeFile 
         Caption         =   "D&ecode File or Attachment"
      End
      Begin VB.Menu mFile_Import 
         Caption         =   "&Import File or Message"
      End
      Begin VB.Menu FileNull 
         Caption         =   "-"
      End
      Begin VB.Menu FileAddress 
         Caption         =   "&Address book..."
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save settings"
      End
      Begin VB.Menu FileNull2 
         Caption         =   "-"
      End
      Begin VB.Menu PrintSetup 
         Caption         =   "Print Setup..."
      End
      Begin VB.Menu FilePrintM 
         Caption         =   "Print Message"
      End
      Begin VB.Menu FilePage 
         Caption         =   "Page setup..."
         Visible         =   0   'False
      End
      Begin VB.Menu FileNull3 
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
      Begin VB.Menu FileNull4 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu EditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu EditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "&Paste"
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
   Begin VB.Menu mPGP 
      Caption         =   "&PGP"
      Begin VB.Menu PGPVersion_PGP5x 
         Caption         =   "Use PGP 5.x"
      End
      Begin VB.Menu PGPVersion_PGP26x 
         Caption         =   "Use PGP 2.6.x"
      End
      Begin VB.Menu mPGPsep 
         Caption         =   "-"
      End
      Begin VB.Menu KeyMenu 
         Caption         =   "&Keys"
         WindowList      =   -1  'True
         Begin VB.Menu KeyCreate 
            Caption         =   "&Create key pair"
         End
         Begin VB.Menu KeyCertify 
            Caption         =   "&Sign or Certify a Key"
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
            Caption         =   "&Add key/keys from message"
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
         Begin VB.Menu mKeyRingIDs 
            Caption         =   "&View Keys on Keyring"
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
         Begin VB.Menu SelectKeyServer 
            Caption         =   "Select a Key Server"
         End
         Begin VB.Menu keySep4 
            Caption         =   "-"
         End
         Begin VB.Menu KeyOptions 
            Caption         =   "&Options..."
         End
      End
      Begin VB.Menu mSeparator 
         Caption         =   "-"
      End
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
      Begin VB.Menu pgpSEP 
         Caption         =   "Encryption Options"
         Begin VB.Menu PGPMultiple 
            Caption         =   "Use multiple keys"
         End
         Begin VB.Menu PGPSelf 
            Caption         =   "Encrypt to self"
         End
         Begin VB.Menu PGPEyes 
            Caption         =   "Eyes only"
            Visible         =   0   'False
         End
         Begin VB.Menu PGPConvent 
            Caption         =   "Conventional encrypt"
         End
         Begin VB.Menu PGPObscurity 
            Caption         =   "Obscurity"
            Visible         =   0   'False
         End
         Begin VB.Menu PGPWrap 
            Caption         =   "Word wrap on encrypt/sign"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu PGPFile 
            Caption         =   "File operations"
         End
         Begin VB.Menu mAttachmentEncryptionOptions 
            Caption         =   "Attachment Encryption Options"
            Begin VB.Menu mDontEncryptAttachment 
               Caption         =   "No Encryption"
            End
            Begin VB.Menu mEncryptAttachmentWithKey 
               Caption         =   "Encrypt with Key"
            End
            Begin VB.Menu mConventionallyEncryptAttachment 
               Caption         =   "Conventionally Encrypt"
            End
         End
      End
      Begin VB.Menu PGPMin 
         Caption         =   "Run PGP minimized"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu PGPVerify 
         Caption         =   "&Verify PGP Distribution"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu PGPSep3 
         Caption         =   "-"
      End
      Begin VB.Menu PGPOptions 
         Caption         =   "PGP &Options..."
      End
   End
   Begin VB.Menu mNewsgroups 
      Caption         =   "&Newsgroups"
      Begin VB.Menu USENETGate 
         Caption         =   "mail2news"
         Begin VB.Menu Prepare_Usenet_Nym 
            Caption         =   "Prepare Mail2News Message using Nym"
         End
         Begin VB.Menu Prepare_usenet_standard 
            Caption         =   "Prepare Mail2News Message"
         End
         Begin VB.Menu mSendNewsGroupMessage 
            Caption         =   "Send Newsgroup Message"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu GetNews 
         Caption         =   "&Newsgroup manager"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu NewsNull2 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mFingerOps 
      Caption         =   "&Finger Operations"
      Begin VB.Menu GetFinger 
         Caption         =   "Finger"
      End
      Begin VB.Menu PI_Test_Click 
         Caption         =   "Special Test"
         Visible         =   0   'False
      End
      Begin VB.Menu DoConnect 
         Caption         =   "Connect"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mMessage 
      Caption         =   "&Message"
      Begin VB.Menu TransferSend 
         Caption         =   "&Send"
         Shortcut        =   ^S
      End
      Begin VB.Menu TransferReply 
         Caption         =   "&Insert reply markers"
      End
      Begin VB.Menu mTransferXHeaders 
         Caption         =   "&Mail-Headers..."
         Begin VB.Menu mAddMailHeaders 
            Caption         =   "VIew Mail Headers"
         End
         Begin VB.Menu mEnableMailHeaders 
            Caption         =   "Enable Mail Headers"
         End
      End
      Begin VB.Menu MsgSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu TransferEu 
         Caption         =   "Trans&fer to app"
         Shortcut        =   ^T
         Visible         =   0   'False
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
         Caption         =   "Prepare Remail Message"
         Visible         =   0   'False
      End
      Begin VB.Menu MsgSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu EmailWrap 
         Caption         =   "Word &wrap message"
         Enabled         =   0   'False
         Shortcut        =   ^W
         Visible         =   0   'False
      End
      Begin VB.Menu MsgSep3 
         Caption         =   "-"
      End
      Begin VB.Menu TransferOptions 
         Caption         =   "Email Application &Options..."
      End
   End
   Begin VB.Menu mNym 
      Caption         =   "N&ym"
      Begin VB.Menu TransferPrepare 
         Caption         =   "&Prepare nym message... "
         Shortcut        =   ^Y
      End
      Begin VB.Menu TransferEncrypt 
         Caption         =   "&Encrypt nym message"
         Visible         =   0   'False
      End
      Begin VB.Menu mFile_DecryptNymMessage 
         Caption         =   "Decrypt a Nym Message"
         Visible         =   0   'False
      End
      Begin VB.Menu NymNull1 
         Caption         =   "-"
      End
      Begin VB.Menu TransferNym 
         Caption         =   "&Create nym..."
      End
      Begin VB.Menu NymReplyChange 
         Caption         =   "Change nym &reply block..."
      End
      Begin VB.Menu mShowNyms 
         Caption         =   "Show Nyms"
      End
      Begin VB.Menu NymShow 
         Caption         =   "&Show nym server stats..."
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu HelpInfo 
         Caption         =   "&Information..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mRegistration 
         Caption         =   "Enter registration details"
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
Attribute VB_Name = "frmMessageDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Win As New CWindow
'Public vb2spgpContext As New spgpParms
Private m_BusyCancel As Boolean
Private m_ControlKey As Boolean

Private Sub btnMixPath_Click()
Dim SectionName As String
SectionName = "Remailer Info" & ""
gMixPath = ReadProfile(SectionName, "MixmasterPath")
frmMixmasterHome.Show vbModal
If gMixPath = "" Or Not iFileExists(gMixPath & "\mixmaste.exe") Then
    gMixPath = App.Path
    MsgBox "Mixmaster was not found.  Path reset to the application directory", vbApplicationModal + vbQuestion, "Mixmaster"
    Exit Sub
End If
On Error Resume Next
If Not iFileExists(gMixPath & "\mixmaster.htm") Then FileCopy App.Path & "\mixmaster.htm", gMixPath & "\mixmaster.htm"
If Not iFileExists(App.Path & "\type2.lis") Then FileCopy App.Path & "\type2.lis", App.Path & "\type2.lis"
If Not iFileExists(gMixPath & "\pubring.mix") Then FileCopy App.Path & "\pubring.mix", gMixPath & "\pubring.mix"
End Sub

Private Sub btnTo_Click(Index As Integer)
ShowStatus ("")
GetRecipient
End Sub


Private Sub cmbRemailerSelect_Click()

Select Case cmbRemailerSelect.ListIndex
        Case REMAILER_NONE
            DontUseRemailer
            Unload frmRemailerList
            SSRibbon1(7).Enabled = True
            SSRibbon1(0).Picture = SSRibbon1(20).Picture
            ShowAttachment ("")
        Case REMAILER_CYPHERPUNK
            frmRemailerList.Caption = "Cypherpunk Remailer List"
            frmRemailerList.Show
            UseCypherPunk
            WriteProfile "Remailer Info", "EncryptionToRemailers", "True"
        Case REMAILER_MIX
            frmRemailerList.Show
            frmRemailerList.Caption = "Mixmaster Remailer List"
            UseMixmaster
            ShowAttachment ("")
    End Select
SaveSettings ' this will save the options
End Sub

Private Sub cmbRemailerSelect_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub









Private Sub Combo3_Change()

End Sub

Private Sub Command1_Click()
Dim a As String
SetMessageReadState
Exit Sub
a = StripExt("c:\j\junk.asc")
a = StripFileName("c:\j\junk.asc")
a = GetExt("c:\j\junk.asc")
Dim itmX As ListItem
lvwAttachments.Icons = frmMain.ImageList1
Set itmX = lvwAttachments.ListItems.Add(, , "Hello")
itmX.Icon = 2
End Sub

Private Sub DoConnect_Click()
    CheckConnection
End Sub

Private Sub EditClrAll_Click()
    txtTo.Text = ""
    txtSubject.Text = ""
    txtCC.Text = ""
    MessageArea.Text = ""
    'txtXHeader.Text = ""
End Sub


Private Sub EditClrMsg_Click()
    MessageArea.Text = ""
End Sub

Private Sub EditCopy_Click()
    EditPerform WM_COPY
End Sub

Private Sub EditCopyMsg_Click()
    If Len(MessageArea) > 0 Then
        Clipboard.Clear
        Clipboard.SetText MessageArea.Text 'SelStart 'Mid(MessageArea, 1, Length)
    End If
End Sub

Private Sub EditCut_Click()
  EditPerform WM_CUT
End Sub

Private Sub EditHeader_Click()
    MessageArea.SelText = gHeader
End Sub

Private Sub EditOptions_Click()
    Form17.Show 1
End Sub

Private Sub EditPaste_Click()
  EditPerform WM_PASTE
End Sub

Private Sub EditPasteMsg_Click()
    MessageArea.Text = Clipboard.GetText()
End Sub



Public Sub EditSelectAll_Click()
    MessageArea.SetFocus
    MessageArea.SelStart = 0
    MessageArea.SelLength = Len(MessageArea)
End Sub

Private Sub EditSig_Click()
    MessageArea.SetFocus
    MessageArea.SelStart = Len(MessageArea.Text)
    MessageArea.SelText = gCRLF + gCRLF + gSig
End Sub
Private Sub EmailWrap_Click()
'turn this off for the present
MessageArea.Text = InsertCRLFs()
End Sub
Private Sub FileAddress_Click()
    frmAddressBook.Show vbModal
End Sub

Private Sub FileConnect_Click()
    '   IPPort1.WinsockLoaded = True
    FileConnect.Enabled = False
    FileDisconnect.Enabled = True
    TransferSend.Enabled = True
   ' EmailScan.Enabled = True
    '  ShowStatus = "e-mail status"
End Sub


Private Sub FileDisconnect_Click()
    '   IPPort1.WinsockLoaded = False
    FileDisconnect.Enabled = False
    FileConnect.Enabled = True
    TransferSend.Enabled = False
    'EmailScan.Enabled = False
    ' ShowStatus = "Winsock not connected"
End Sub

Private Sub FileExit_Click()
    Unload Me
End Sub

Private Sub FileExport_Click()

Dim FileNum As Integer

    On Error GoTo ExportError
    
    '---------------------------------------------
    'prepare the file save as dialog
    '---------------------------------------------
    CommonDialog1.DialogTitle = "Save message as"
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.FileName = "*.txt"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Action = 2
    FileNum = FreeFile
    Open CommonDialog1.FileName For Output As FileNum
    Print #FileNum, MessageArea.Text
    Close #FileNum
    '---------------------------------------------
    'switch back to normal settings
    '---------------------------------------------
    ChDrive Mid$(App.Path, 1, 3)
    ChDir App.Path
ExportError:
    Exit Sub
End Sub






Private Sub FilePage_Click()
    MsgBox "Not implemented yet."
End Sub

Private Sub FilePrintM_Click()
    Dim Buffer As String
    Dim ndx As Integer
    Dim foo As Integer
    
    If Printers.Count = 0 Then
        MsgBox "No printers are installed.  Can't continue", vbCritical, "Fatal Error."
        Exit Sub
    End If
    CommonDialog1.ShowPrinter
    Printer.Print ""
    MessageArea.SelPrint (Printer.hDC)
    Printer.EndDoc
End Sub


Private Sub FileSave_Click()
    SaveSettings
End Sub





Private Sub Form_Click()
ShowStatus ("")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'see VBPJ dec 97
If KeyCode = vbKeyControl Then
    m_ControlKey = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
    m_ControlKey = False
End If
End Sub

Private Sub Form_Load()
    Dim FileNum As Integer
    Dim where As Integer
    Dim First As Integer
    Dim strBuffer As String
    Dim Tmp As Integer
    Dim Stimer As Single
    Dim SectionName As String
    Dim cmd As String
    Dim Ecount As Integer
    Dim iResult As Integer
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim lVersion As Long
    Dim i As Integer
    
    Dim App As New CApplication
    On Error GoTo LoadError
    
    Win.Center Me, Null
    
    '---------------------------------------------
    'init stuff
   ' SubClass (Me.hWnd) 'don't allow them to use the system exit
    '---------------------------------------------
    'Set PI version here
    '---------------------------------------------
    Me.Caption = "Private Idaho (" & App.Version & ") for Win9x/NT"
   
    Set lvwAttachments.Icons = FileIconsImageList1.Icons
    Set lvwAttachments.SmallIcons = FileIconsImageList1.SmallIcons

    'This is the default
    MailHeader(0).Id = ""
    MailHeader(1).Id = "Newsgroups: "
    MailHeader(2).Id = "Subject: "
    MailHeader(3).Id = "Message ID: "
    MailHeader(4).Id = "Reference: "
    MailHeader(5).Id = "X-Header: "
    StatusBar.Height = TextHeight("Test") * 1.8
    gObscurity = 0
    gRemailerType = REMAILER_NONE
    gNewsgroupType = 0
    gMinState = 1
    gCutStr = ""
    gLatentStr = ""
    gCRLF = vbCrLf
    gSMTPLog = 0
    gNymState = gintNYM_IDLE
    gPassPhrase = ""
    gPiStr = "Private Idaho"
    gExit = 0
    gPGPTempFile = "pvtidaho"
    gPGPResponse.Count = 0
    'First get PGP path
    'get the PGP key
    '--------------------------------------------
    'Check for PGP Version
    '--------------------------------------------
   ' lVersion = spgpVersion
    'If lVersion > 0 Then
      '  ShowStatus = "SPGP Version: " & lVersion & " found!"
   ' Else
   '     ShowStatus = "SPGP not found..!"
   ' End If
       
    
    SectionName = "PGP Info"
    gPGPVersion = ReadProfile(SectionName, "PGP Version")
    If gPGPVersion = PGP26x Then
        PGPVersion_PGP26x.Checked = True
    Else
        PGPVersion_PGP5x.Checked = True
    End If
    gPGPKeyID = ReadProfile(SectionName, "KeyID")
    gPGPPath = ReadProfile(SectionName, "PGPPath")
    If Not gPGPVersion = NoPGP Then
        If gPGPVersion = PGP26x Then
            If gPGPKeyID = NoPGP Or gPGPPath = NoPGP Then
                frmPGPHome.Show vbModal
                If Len(gPGPPath) = 0 Then
            'they don't have PGP, so tell them
                    DisablePGPMenuItems
                End If
            End If
        End If
    Else
       DisablePGPMenuItems
    End If
    
       
    'need to decrypt the rings first
    'were the key rings previously encrypted?
    'HTTP1.UserAgent = gPiStr + " 3.1"
GoTo load0
    tmpstr2 = gPGPPath & "\SECRING.PID"
    If iFileExists(tmpstr2) Then
        gPassPhrase = "2"
        frmPOPPassword.Caption = "Key ring Passphrase"
        frmPOPPassword.Label2.Caption = "Trying to decrypt 'SECRING.PID'. Enter the passphrase to decrypt your key rings:"
        frmPOPPassword.Show vbModal
        tmpstr = StripExt(tmpstr2) & "PGP"
        'cmd = App.Path + "\PIPGPX.PIF +batchmode " + tmpstr2 + " -o " + tmpstr + " -z " + Chr$(34) + gPassPhrase + Chr$(34)
        cmd = gPGPPath & "\PGP " & tmpstr2 & " -o "
        cmd = cmd & tmpstr + " -z " + Chr$(34) + gPassPhrase + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        'delete the version with the pgp extension
        If iFileExists(tmpstr) Then
            Kill tmpstr2
        Else
            MsgBox "Couldn't decrypt SECRING.PGP.  Wrong password.  Run Private Idaho and try again.", vbApplicationModal + vbExclamation, App.Title
            End
        End If
    End If
load0:

GoTo load1
    'tmpstr2 = gPGPPath + "\PUBRING.PID"
    If iFileExists(gPGPPath + "\PUBRING.PID") Then
        If gPassPhrase <> "2" Then
            frmPOPPassword.Caption = "Key ring Passphrase"
            frmPOPPassword.Label2.Caption = "Trying to decrypt 'PUBRING.PID'. Enter the passphrase to decrypt your key rings:"
            frmPOPPassword.Show vbModal
        End If
            
        'tmpstr = StripExt(tmpstr2) + "PGP"
        cmd = gPGPPath & "\PGP " & gPGPPath & "\PUBRING.PID" & " -o "
        cmd = cmd & StripExt(gPGPPath + "\PUBRING.PID") + "PGP" & " -z " + Chr$(34) + gPassPhrase + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        
        'delete the version with the pgp extension
        If iFileExists(StripExt(gPGPPath + "\PUBRING.PID") + "PGP") Then
            Kill tmpstr2
        End If
    End If
load1:
GoTo load2
    'secure mode means store the INI and nym file encrypted
    'decrypt the nym file
    If iFileExists(App.Path + "\NYMS.ASC") Then
        'cmd = App.Path & "\PIPGPX.PIF " & App.Path & "\NYMS.ASC" + " -o " + App.Path + "\NYMS.TXT -z " + Chr$(34) + gPassPhrase + Chr$(34)
        If gPassPhrase <> "2" Then
            frmPOPPassword.Caption = "Key ring Passphrase"
            frmPOPPassword.Label2.Caption = "Enter the passphrase to decrypt your key rings:"
            frmPOPPassword.Show vbModal
        End If
        cmd = gPGPPath & "\PGP " & App.Path
        cmd = cmd & "\NYMS.ASC -o " & App.Path + "\NYMS.TXT  -z "
        cmd = cmd & Chr$(34) & gPassPhrase + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        If iFileExists(App.Path + "\NYMS.TXT") Then
            Kill App.Path + "\NYMS.ASC"
        End If
    End If
load2:
GoTo load3
'decrypt the pubkeys.out file
    If iFileExists(App.Path + "\PUBKEYS.ASC") Then
           ' cmd = App.Path & "\PIPGPX.PIF " & App.Path & "\PUBKEYS.ASC" + " -o " + App.Path + "\PUBKEYS.OUT -z " + Chr$(34) + gPassPhrase + Chr$(34)
       If gPassPhrase <> "2" Then
            frmPOPPassword.Caption = "Key ring Passphrase"
            frmPOPPassword.Label2.Caption = "Enter the passphrase to decrypt your key rings:"
            frmPOPPassword.Show vbModal
        End If
       cmd = gPGPPath & "\PGP " & App.Path & "\PUBKEYS.ASC -o " & App.Path & "\PUBKEYS.OUT -z " + Chr$(34) + gPassPhrase + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        If iFileExists(App.Path + "\PUBKEYS.OUT") Then
            Kill App.Path + "\PUBKEYS.ASC"
        End If
    End If
        'decrypt the address file
    If iFileExists(App.Path + "\ADDRESS.ASC") Then
           ' cmd$ = App.Path + "\PIPGPX.PIF " + App.Path + "\ADDRESS.ASC" + " -o " + App.Path + "\ADDRESS.TXT -z " + Chr$(34) + gPassPhrase + Chr$(34)
        If gPassPhrase <> "2" Then
            frmPOPPassword.Caption = "Key ring Passphrase"
            frmPOPPassword.Label2.Caption = "Trying to decrypt file 'Address.asc' which is your private address book. Enter the passphrase to decrypt your key rings:"
            frmPOPPassword.Show vbModal
        End If
        cmd = gPGPPath & "\PGP " & App.Path & "\ADDRESS.ASC -o " & App.Path & "\ADDRESS.TXT -z " + Chr$(34) + gPassPhrase + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        If iFileExists(App.Path + "\ADDRESS.TXT") Then
            Kill App.Path + "\ADDRESS.ASC"
        End If
    End If
    MousePointer = vbArrow

load3:
GoTo load4
    'see if we want the passphrase loaded into memory
    If gPassPhrase = "" Then
        'hasn't been set yet
        If ReadProfile(SectionName, "Passphrase") = "true" Then
            gPassPhrase = "1"
            frmPOPPassword.Caption = "PGP Passphrase"
            frmPOPPassword.Label2.Caption = "Enter your PGP passphrase:"
            frmPOPPassword.Show vbModal
        End If
    End If
    where = 5
load4:
GoTo load5
    'get the PGP key
    SectionName = "PGP Info"
    gPGPKeyID = ReadProfile(SectionName, "KeyID")
    gPGPPath = ReadProfile(SectionName, "PGPPath")
    If gPGPKeyID = "" Or gPGPPath = "" Then
        Unload frmSplash
        MsgBox "Private Idaho needs to be configured.  You won't need to go through these steps the next time you run Private Idaho.", 64, gPiStr
        'assume this is the first time run and generate
        'the pubkeys.out file
        First = 1
        frmPGPOptions.Command1.Enabled = False
        frmPGPOptions.Show vbModal
        DoEvents
        UpdatePublicKeysFile
    End If
load5:
    
    'gPGPKeyID = Chr$(34) & gPGPKeyID & Chr$(34)
'get the e-mailer app info
    SectionName = "Mailer Info"
    gEmailer = ReadProfile(SectionName, "Emailer")
    If gEmailer = "" Then
        Form5.Command2.Enabled = False
        Form5.Show 1
        Form5.Command2.Enabled = True
    End If
    gtranScript = ReadProfile(SectionName, "gtranScript")
    
    If gEmailAddress = "" Then
        If First = 1 Then
            frmFileOptions.Show 1
        Else
            WriteProfile SectionName, "EmailAddress", gEmailAddress
        End If
    End If
    
    'see if SMTP logging should be turned on
    tmpstr = ReadProfile(SectionName, "SMTPLog")
    If tmpstr <> "" Then
        gSMTPLog = 1
        gSMTPFile = FreeFile
        Open App.Path + "\netlog.txt" For Output As gSMTPFile
    End If
    
    'get PGP tmpstr file info
    SectionName = "PGP Info"
    gPGPFile = ReadProfile(SectionName, "PGPfooFile")
    
    'it should never be this state, but...
    If gPGPFile = "" Then
        gPGPFile = "pvtidaho"
        WriteProfile SectionName, "PGPTempFile", gPGPFile
    End If
    
    '************************************************************
    'Section WEB INFO
    '************************************************************
GoTo load6
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
    gc2WWWAnon = ReadProfile(SectionName, "gc2WWWAnon")
    If gc2WWWAnon = "" Then
        gc2WWWAnon = "http://www.anonymizer.com:8080/"
        WriteProfile SectionName, "gc2WWWAnon", gc2WWWAnon
    End If


    
    'set the transfer menus
    SectionName = "Options"
    tmpstr = ReadProfile(SectionName, "App1Name")
    If tmpstr <> "" Then
        TransferApps.Visible = True
        TransferApp1.Caption = tmpstr
        TransferApp1.Visible = True
    End If
    tmpstr = ReadProfile(SectionName, "App2Name")
    If tmpstr <> "" Then
        TransferApp2.Caption = tmpstr
        TransferApp2.Visible = True
    End If
    tmpstr = ReadProfile(SectionName, "App3Name")
    If tmpstr <> "" Then
        TransferApp3.Caption = tmpstr
        TransferApp3.Visible = True
    End If
    tmpstr = ReadProfile(SectionName, "App4Name")
    If tmpstr <> "" Then
        TransferApp4.Caption = tmpstr
        TransferApp4.Visible = True
    End If
    where = 8
load6:
GoTo load7:
    'get the address list
    FileNum = FreeFile
    Dim Item As String
    If iFileExists(App.Path + "\ADDRESS.TXT") Then
        Open App.Path + "\ADDRESS.TXT" For Input As FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, Item
            txtTo.AddItem Item
        Loop
        Close FileNum
    End If
    
    'get the header/sig
    'if file exists then open
    FileNum = FreeFile
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
load7:
    'set the newsgroup menu for cp state
    'USENETGate.Visible = True
    'USENETFi.Visible = False
    'UseNetSoda.Visible = False
    'load the remailer array on startup
    'Mote allow the user to update once PI is running
    'Remailer.txt will likely be the most uptodate...
    'If iFileExists(App.Path + "\remailer.txt") Then
        'InitializeRemailers (App.Path + "\remailer.txt")
    'Else
        'If iFileExists(App.Path + "\remailer.htm") Then InitializeRemailers (App.Path + "\remailer.htm")
    'End If
        
    'restore settings stored in the ini
    RestoreSettings
    SetAttachmentEncryptionOptions
    If gPGPVersion = PGP26x Then
        If Not iFileExists(App.Path + "\PUBKEYS.OUT") Then
            frmUpdatePublicKeys.Show vbModal
        End If
    End If
    'Check if first installation
    If iFileExists(App.Path + "\HELLO.TXT") Then
        FirstInitialisation
        On Error Resume Next
        FileNum = FreeFile
        Open App.Path + "\HELLO.TXT" For Binary As FileNum
        strBuffer = String(LOF(FileNum), " ")
        Get FileNum, , strBuffer
        MessageArea.Text = strBuffer
        Close FileNum
        Kill App.Path + "\HELLO.TXT"
    End If
    
    
    MessageArea.SelBold = False
    ShowAttachment ("")
    '
    'Set up the remailer options
    '
    cmbRemailerSelect.AddItem "Don't use remailers", 0
    cmbRemailerSelect.AddItem "Use Type 1 remailers (Cypherpunk)", 1
    cmbRemailerSelect.AddItem "Use Type 2 remailers (Mixmaster)", 2
        
    If gRemailerType = REMAILER_CYPHERPUNK Then cmbRemailerSelect.ListIndex = 1
    If gRemailerType = REMAILER_MIX Then cmbRemailerSelect.ListIndex = 2
    If gRemailerType = REMAILER_NONE Then cmbRemailerSelect.ListIndex = 0
    
    
   ' House Keeping
    Set App = Nothing
    Set Win = Nothing
    MousePointer = vbDefault
    ShowStatus ("")
    Exit Sub

LoadError:
     MousePointer = vbArrow
    MsgBox Err.Description & " Main Form Load " & Str$(where) & "-" + Str$(Err)
    Unload frmSplash
    Err.Clear
End Sub


Private Sub Form_Resize()
'Dim BottomMargin As Integer
'Dim LeftMargin As Integer
'Static StatusTop As Long
'DoEvents
On Error Resume Next
  ' BottomMargin = 800
   'LeftMargin = 200
  ' DoEvents
   If WindowState <> 1 Then
        DoEvents
        If lvwAttachments.Visible Then
            lvwAttachments.Top = Height - MessageArea.Top - lvwAttachments.Height '- StatusBar.Height '- 1500
            lvwAttachments.Width = Width - MessageArea.Left - 200
        End If
        
        
        If lvwAttachments.Visible Then
            MessageArea.Height = lvwAttachments.Top - MessageArea.Top
            MessageArea.Width = Width - MessageArea.Left - 200
        Else
            MessageArea.Height = Height - MessageArea.Top - StatusBar.Height - 50  '- 2000
            MessageArea.Width = Width - MessageArea.Left - 200
        End If
        
        If lvwAttachments.Visible Then
            'MessageTab.Height = lvwAttachments.Top + lvwAttachments.Height + 50
            'MessageTab.Width = Width - MessageTab.Left - 100
        Else
        
            'MessageTab.Height = Height - MessageTab.Top + MessageArea.Top + MessageArea.Height + 50
           ' MessageTab.Width = Width - MessageTab.Left - 100
        End If
        
        
        StatusBar.Panels(4).Width = MessageArea.Width - StatusBar.Panels(4).Left
        
        SSPanel1.Width = Width - 8
        
        txtCC.Width = MessageArea.Width - txtCC.Left
        txtSubject.Width = MessageArea.Width - txtSubject.Left
               
        txtTo.Width = MessageArea.Width - txtTo.Left
        'frmPI.Enabled = True
    End If

End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnSubClass(Me.hwnd)
    Unload frmRemailerList
    Set frmPI = Nothing
    'InstanceNumber=InstanceNumber-1
End Sub

Private Sub GetNews_Click()
    SendNewsCommand
End Sub

Private Sub GetFinger_Click()
    If CheckConnection Then
        frmFingerCommand.Show  'SendFinger
    End If
End Sub

Private Sub HelpAbout_Click()
  frmAbout.Show vbModal
End Sub
Private Sub HelpInfo_Click()
    ShowWebHelp
End Sub


Private Sub HelpSys_Click()
    HelpSys.Checked = Not HelpSys.Checked
End Sub
Private Sub HTTP1_EndTransfer(Direction As Integer)
   If Direction = 0 Then Exit Sub '0 = client
    On Error GoTo HTTPEndTransferErr
    Select Case gWebState
    Case GETREMAILERUPDATE
        DoEvents
        'Convert to txt format by adding crlfs etc
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path & "\remailer.htm", gWebPage
        'Get remailer from .htm file into Remailer.txt
        frmRemailerList.InitializeRemailers (App.Path + "\remailer.htm")
        frmRemailerList.lblstatus = "Filling Remailer List..."
        'This will fill the matched remailer list
        frmRemailerList.SortRemailers
        frmRemailerList.FillRemailerList
    Case MIXUPDATE
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path + "\mixmaster.htm", gWebPage
        frmRemailerList.InitializeMixRemailers (App.Path & "\mixmaster.htm")
        frmRemailerList.lblstatus = "Filling Remailer List."
        frmRemailerList.FillRemailerList
    Case TYPE2UPDATE
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        frmRemailerList.lblstatus = "Writing data to Mixmaster Type2.lis file."
        WriteStrToFile App.Path & "\type2.lis", gWebPage
    Case PUBRINGUPDATE
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        ShowStatus ("Updating " & gMixPath & "\pubring.mix file.")
        WriteStrToFile App.Path & "\pubring.mix", gWebPage
    Case GETSERVERKEY
        gWebPage = ConvertStr(gWebPage)
        'MessageArea.SelText = gWebPage
        MessageArea.SelText = GetKeyBlock(gWebPage)
        MessageArea.SetFocus
    Case GETREMAILERKEYS
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        MessageArea.SelText = gWebPage
        MessageArea.SetFocus
    End Select
    Screen.MousePointer = 0
    frmPI.HTTP1.WinsockLoaded = False
    gWebState = HTTPIDLE
    HideStatus
    Exit Sub

HTTPEndTransferErr:
    MsgBox Err.Description & ": -> HTTPEndTransferErr"
    frmPI.HTTP1.WinsockLoaded = False
    frmPI.HTTP1.Action = a_Idle
    Err.Clear
End Sub

Private Sub HTTP1_Error(ErrorCode As Integer, Description As String)
    'gWebState = IdleState
    frmPI.HTTP1.Action = a_Idle
    frmPI.HTTP1.WinsockLoaded = False
    MsgBox "HTTP error " & Format$(ErrorCode) & ".  " _
            & Description, 16, "HTTP Error"
End Sub

Private Sub HTTP1_StartTransfer(Direction As Integer)
    If Direction = 0 Then Exit Sub '0 is client
    '---------------------------------------------
    'gWebPage will hold the text contents of the downloaded web page
    '---------------------------------------------
    gWebPage = ""
    frmRemailerList.lblstatus = "Download started..."
    Screen.MousePointer = vbHourglass
End Sub

Private Sub HTTP1_Transfer(Direction As Integer, BytesTransferred As Long, Text As String)

    On Error GoTo HTTPTransferErr
    If Direction = 0 Then Exit Sub '0 = client
    frmStatus.Label3.Caption = BytesTransferred
    gWebPage = gWebPage & Text
    Exit Sub

HTTPTransferErr:
    MsgBox Err.Description & " : - > HTTPTransfer", vbApplicationModal, App.Title
    frmPI.HTTP1.Action = a_Idle
    frmPI.HTTP1.WinsockLoaded = False
    Err.Clear
End Sub


Private Sub IPPort1_DataIn(Text As String, EOL As Boolean)
Dim Done As Boolean
Dim Prechopped As Boolean
Dim Pos As Integer
Dim foo As String
Dim MyText As String
Dim TLen As Integer
    TLen = Len(Text)
    Done = False
    Prechopped = True
    MyText = ""
    Pos = 1
    Do While Not Done
        foo = Mid(Text, Pos, 1) '
        If foo = vbLf Then
            MessageArea.SelText = MyText & vbCrLf
            MyText = ""
            Prechopped = False
        ElseIf foo = vbCrLf Then
            MessageArea.SelText = MyText & vbCrLf
            MyText = ""
            Prechopped = False
       Else
           MyText = MyText + foo
        End If
        Pos = Pos + 1
        If Pos > TLen Then
            Done = True
            If Prechopped Then
                MessageArea.Text = MessageArea.Text + Text + vbCrLf
            End If
        End If
    Loop
End Sub
Private Sub IPPort1_Disconnected(StatusCode As Integer, Description As String)

    ShowStatus ("Status " & Description)
End Sub

Private Sub IPPort1_Error(ErrorCode As Integer, Description As String)

    MessageArea.Text = Err.Number & ": " & Err.Description
End Sub

Private Sub IPPort1_ReadyToSend()
Dim foo As Integer
Dim Storestring1 As String
Dim Storestring2 As String

Dim cmd As String
Dim iResult As Integer

    On Error GoTo KeyCertifyError
    '[??]
    Exit Sub
    gKeyID = ""
    gCancelButton = False
    
    '---------------------------------------------
    'display the list of personal keys to sign with
    '---------------------------------------------
    'z$ = Form26.Label1.Caption
    Form26.Label1.Caption = "Select a key to use for signing the public key."
    CheckMultipleKey
    
    '---------------------------------------------
    'display the public key ring
    '---------------------------------------------
    Storestring1 = frmSelectUserID.Label2.Caption
    Storestring2 = frmSelectUserID.Label1.Caption
    frmSelectUserID.Label2.Caption = "Select a key to sign"
    frmSelectUserID.Label1.Caption = "from the public key ring."
    frmSelectUserID.Show 1
    frmSelectUserID.Label2.Caption = Storestring1
    frmSelectUserID.Label1.Caption = Storestring2
    Unload frmSelectUserID
    
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd = ""
        cmd = gPGPPath & "\PGP -ks " + Chr$(34) + gKeyID + Chr$(34) + " -u " + Chr$(34) + gPGPKeyID + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        UpdatePublicKeysFile
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
    End If
Exit Sub
    
KeyCertifyError:
    If Err.Number = 32755 Then
        MsgBox "Could not decrypt the message.  Suggest you open a DOS session and type pgp " + gPGPFile + ".out at the the command prompt in your PGP directory to find out what happened."
    Else
        Err.Number = 53  'no pgpfooFile, it means PGP command did not complete, just abort
    End If
    Err.Clear
    
End Sub

Private Sub IPInfo1_Click()

End Sub

Private Sub IPPort1_Click()

End Sub

Private Sub KeyCertify_Click()
Dim msg As String
If gPGPVersion = PGP26x Then
    frmSignKeys.Show vbModal
Else
    msg = "The Key Certification Process is not available from this version of PGP. " & vbCrLf
    msg = msg & "It is better that you use the utilites that come with PGP to create your keys." & vbCrLf
    MsgBox msg, vbExclamation, "Key Certification"
End If
End Sub
Private Sub KeyCreate_Click()
Dim msg As String
On Error GoTo BadKeys
If gPGPVersion = PGP26x Then
    CreatePGPKeyPair
    UpdatePublicKeysFile
Else
    msg = "Create Key-Pair is not available from this version of PGP. " & vbCrLf
    msg = msg & "It is better that you use the utilites that come with PGP to create your keys." & vbCrLf
    MsgBox msg, vbExclamation, "Create Key Pair Command"
End If
Exit Sub
BadKeys:
    MsgBox Err.Description & " in Newkeys)", vbCritical + vbApplicationModal, App.Title
    Err.Clear
End Sub

Private Sub KeyEditTrust_Click()

'---------------------------------------------
Dim Storestring1 As String
Dim Storestring2 As String
Dim cmd As String
'Dim Ecount As Integer
'Dim iResult As Integer
Dim msg As String
If gPGPVersion = PGP26x Then
   gKeyID = ""
    gCancelButton = False
    Storestring1 = frmSelectUserID.Label2.Caption
    Storestring2 = frmSelectUserID.Label1.Caption
    frmSelectUserID.Label2.Caption = "Select a key to sign"
    frmSelectUserID.Label1.Caption = "from the public key ring."
    frmSelectUserID.Show vbModal
    frmSelectUserID.Label2.Caption = Storestring1
    frmSelectUserID.Label1.Caption = Storestring2
    Unload frmSelectUserID
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd = gPGPPath & "\PGP -ke " + Chr$(34) + gKeyID + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        UpdatePublicKeysFile
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
    End If
Else
    msg = "The Edit Key Trust Process is not available from this version of PGP. " & vbCrLf
    msg = msg & "It is better that you use the utilites that come with PGP to create your keys." & vbCrLf
    MsgBox msg, vbExclamation, "Edit Key Trust"
End If
End Sub
Private Sub KeySign_Click()
    Dim cmd As String
    Dim Ecount As Integer
    Dim iResult As Integer
    Dim Storestring1 As String
    Dim Storestring2 As String
    
    gKeyID = ""
    gCancelButton = False
    Storestring1 = frmSelectUserID.Label2.Caption
    Storestring2 = frmSelectUserID.Label1.Caption
    frmSelectUserID.Label2.Caption = "Select your key from"
    frmSelectUserID.Label1.Caption = "the public key ring."
    frmSelectUserID.Show 1
    frmSelectUserID.Label2.Caption = Storestring1
    frmSelectUserID.Label1.Caption = Storestring2
    Unload frmSelectUserID
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd = gPGPPath & "\PGP -ks " + Chr$(34) + gKeyID + Chr$(34) + " -u " + Chr$(34) + gKeyID + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        UpdatePublicKeysFile
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
End Sub

Private Sub KeyOptions_Click()
    gMultiType = 1
    Form25.Show vbModal
End Sub

Private Sub KeySubmit_Click()
Dim SectionName As String
    If CheckConnection Then
        'is there a key there?
        If InStr(1, MessageArea.Text, "PUBLIC KEY BLOCK") = 0 Then
            MsgBox "Please enter the PGP public key you'd like to submit to the key server in the Message box."
            Beep
            Exit Sub
        End If
        If gSubKeyURL = "" Then
            SectionName = "Net Info"
            gSubKeyURL = ReadProfile(SectionName, "SubmitKeyURL")
            If gSubKeyURL = "" Then frmSelectKeyServer.Show vbModal
            If gSubKeyURL = "" Then
                gSubKeyURL = "pgp-public-keys@pgp.ai.mit.edu"
            End If
        End If
        txtTo.Text = gSubKeyURL
            
            'http://pgp5.ai.mit.edu:11371/pks/lookup?op=vindex&search=alex@itech.net.au
            'gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
            'WriteProfile SectionName, "GetKeyURL", gGetKeyURL
            'End If
        'End If
        
        'Tell the server to add the key.
        txtSubject.Text = "add"
        ShowStatus ("")
        SendToOutBox
        'SendMailMessage
    End If
End Sub

Private Sub ShowStatus_Click()
    ShowStatus ("")
End Sub



Private Sub lvwAttachments_DblClick()
Dim res As Long
Dim obj As Object
Dim Attachment As String
'demo = "d:\temp\my-wife0.jpg"
Attachment = App.Path & "\mailbox\attachments\" & lvwAttachments.SelectedItem.Text
res = ShellExecute(Me.hwnd, "open", Attachment, vbNullString, CurDir, SW_SHOW)
If res < 32 Then
    MsgBox "Error was encountered launching the application associated with this attachment.", vbApplicationModal + vbCritical, "Launch Attachment"
End If
End Sub

Private Sub lvwAttachments_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim i As String
'Text1.Text = Item
i = lvwAttachments.SelectedItem.Text
End Sub

Private Sub lvwAttachments_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then
    lvwAttachments.ListItems.Remove (lvwAttachments.SelectedItem.Index)
End If

End Sub

Private Sub mAddMailHeaders_Click()
frmMailHeader.Show
End Sub

Private Sub mConventionallyEncryptAttachment_Click()
mEncryptAttachmentWithKey.Checked = False
mDontEncryptAttachment.Checked = False
mConventionallyEncryptAttachment.Checked = True
End Sub

Private Sub mDecodeFile_Click()
Dim EncodedData As String
Dim FileNum As Integer
Dim FileName As String
Dim TextLine As String
Dim NumBytes As Long
Dim msg As String
Dim foo As Long
Dim i As Integer
Dim FileSize As Long
Dim LineCount As Long
On Error GoTo ImportError
If Len(frmPI.MessageArea.Text) = 0 Then
    MsgBox "The is nothing in the message area.", vbExclamation + vbApplicationModal, "Empty Message Area"
    DoEvents
    Exit Sub
End If
'FileName = "temp"
frmPI.ShowStatus ("Looking for beginning of file...")
foo = InStr(1, MessageArea, "name=")
i = 0
If Not foo = 0 Then
    foo = foo + Len("name=""")
    Do While i < 128
        'InStr(foo + i, MessageArea, vbCrLf) = 0
        TextLine = Mid(MessageArea, foo + i, 1)
        If TextLine = vbCr Or TextLine = vbLf Or TextLine = """" Then Exit Do
        FileName = FileName & TextLine
        i = i + 1
    Loop
    frmPI.ShowStatus ("Found file: " & FileName)
    DoEvents
    'FileName = Mid(MessageArea, foo + 1, i - 1) 'jump over " +1 and -1
End If
If Not InStr(1, MessageArea, "base64") = 0 Then
    frmPI.NetCode1.Format = f_BASE64
Else
    frmPI.NetCode1.Format = f_UUEncode
End If
    
frmPI.NetCode1.MaxFileSize = 0
frmPI.NetCode1.Overwrite = True
frmPI.ShowStatus ("Decoding file: " & FileName)
DoEvents
frmPI.NetCode1.FileName = App.Path & "\" & FileName
On Error GoTo ImportError
frmPI.NetCode1.EncodedData = frmPI.MessageArea.Text
frmPI.NetCode1.Action = 3 'Decode to file
frmPI.NetCode1.Action = 0
'filenum = FreeFile
'Open frmPI.NetCode1.FileName For Output As filenum
'Print #filenum, frmPI.NetCode1.DecodedData
'Close filenum
DoEvents
MousePointer = vbDefault

CommonDialog1.DialogTitle = "Save file"
'CommonDialog1.FilterIndex = 1
If Not FileName = "" Then
    CommonDialog1.Filter = GetExt(FileName)
    CommonDialog1.FileName = App.Path & "\" & FileName
Else
    CommonDialog1.Filter = GetExt(frmPI.NetCode1.FileName)
    CommonDialog1.FileName = frmPI.NetCode1.FileName
End If
'CommonDialog1.CancelError = True
'CommonDialog1.FileName = App.Path & "\" & FileName

CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave
FileNum = FreeFile
Open CommonDialog1.FileName For Output As FileNum
Print #FileNum, frmPI.NetCode1.DecodedData
Close FileNum
frmPI.ShowStatus ("Saved file at: " & CommonDialog1.FileName)
ChDir App.Path

Exit Sub
ImportError:
    Beep
    MsgBox Err.Description, vbApplicationModal, App.Title
    MousePointer = vbDefault
    Close FileNum
    ChDir App.Path
    Err.Clear
End Sub

Private Sub mDontEncryptAttachment_Click()
mEncryptAttachmentWithKey.Checked = False
mDontEncryptAttachment.Checked = True
mConventionallyEncryptAttachment.Checked = False
End Sub

Private Sub mEnableMailHeaders_Click()
mEnableMailHeaders.Checked = Not mEnableMailHeaders.Checked
If mEnableMailHeaders.Checked = True Then
    MailHeader(0).Id = "ExtendedHeaders"
Else
    MailHeader(0).Id = ""
End If
End Sub

Private Sub mEncodeFile_Click()
frmEncodeFile.Show vbModal
End Sub

Private Sub mEncryptAttachmentWithKey_Click()
mEncryptAttachmentWithKey.Checked = True
mDontEncryptAttachment.Checked = False
mConventionallyEncryptAttachment.Checked = False
End Sub

Private Sub MessageArea_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim j As Integer
Dim k As String
Dim lListItem As ListItem

'Set lvwAttachments.Icons = FileIconsImageList1.Icons
'Set lvwAttachments.SmallIcons = FileIconsImageList1.SmallIcons
If Data.GetFormat(vbCFFiles) Then
    Dim vFN
    For Each vFN In Data.Files
       ' k = StripFileName(CStr(vFN))
        j = FileIconsImageList1.GetFileIconNum(StripFileName(CStr(vFN)))
        Set lListItem = lvwAttachments.ListItems.Add(, , vFN, j, j)
        'lListItem.SubItems(1) = FileSize("", mExtractCab.CompressedFiles.Item(i).FileSize)
        'lListItem.SubItems(2) = GetFileType(mExtractCab.CompressedFiles.Item(i).FileCabName)
        'lListItem.SubItems(3) = FileTime("", mExtractCab.CompressedFiles.Item(i).FileDate, mExtractCab.CompressedFiles.Item(i).FileTime)
        ' lvwAttachments MessageArea = vFN
        ' MessageArea = Data.GetData(vbCFFiles)
    Next vFN
    
End If
'Effect = 9
End Sub


Private Sub MessageTab_DblClick()

End Sub

Private Sub mFile_DecryptNymMessage_Click()
gNymState = gNYM_DECRYPT
frmMultiNyms.Show vbModal
gNymState = gintNYM_IDLE
End Sub

Private Sub mFile_Import_Click()
    ImportFile
End Sub

Private Sub mFile_Save_Click()
    SaveMessage
End Sub



Private Sub IPPort1_Connected(StatusCode As Integer, Description As String)

    If 0 = StatusCode Then    'OK
        'ShowStatus("")
    Else
        MsgBox "Connection failed: " & Description, vbAbortRetryIgnore, App.Title
    End If
End Sub

Private Sub mKeyRingIDs_Click()
Dim BufferOut As String
Dim i As Long
Dim Count As Long
Dim BufferLen As Long
    
    BufferLen = spgpKeyRingCount() * 256
    BufferOut = String(BufferLen, Chr(0))
    i = spgpKeyRingID(BufferOut, BufferLen)
    Count = CountCRLF(BufferOut)
    Call ChopKeyProps(BufferOut, Count)
     
    vb2spgpContext.SelectPrivateKeys = False
    frmViewKeyRing.lblContext = "This is a list of all keys on your keyring."
    frmViewKeyRing.Show vbModal
    
End Sub

Private Sub mSendNewsGroupMessage_Click()
If MessageArea.Text = "" Then
    MsgBox "There is nothing to send."
    Exit Sub
End If
MailHeader(0).Id = "USENET"
SendPIMessage
MailHeader(0).Id = ""
mSendNewsGroupMessage.Enabled = False
End Sub

Private Sub mShowNyms_Click()
frmNymsList.Show
End Sub

Private Sub NNTP1_GroupOverview(ArticleNumber As Long, Subject As String, From As String, ArticleDate As String, MessageID As String, References As String, ArticleSize As Long, ArticleLines As Long, OtherHeaders As String)

   'ArticleNumber contains the number of the article within the group.
    'foo1 = ArticleNumber
    'Subject contains the subject of the article.
    'foo2 = Subject
    'From contains the email address of the article author.
    'foo3 = From
    'ArticleDate contains the date the article was posted.
    'foo4 = ArticleDate
    'MessageID contains the unique message id for the article.
    'foo5 = MessageID
    'References contains the message ids for the articles this article refers to (separated by spaces).
    'foo6 = References
    'ArticleSize contains the size of the article in bytes.
   ' foo7 = ArticleSize
    'ArticleLines contains the number of lines in the article.
   ' foo8 = ArticleLines
    'OtherHeaders contains any other article headers that NewsServer chooses to display for the article.
   ' foo9 = OtherHeaders
End Sub

Private Sub NNTP1_Header(Field As String, Value As String)

    'The Field parameter contains the name of the header (same case as it is delivered).
    'foo1 = Field
    'The Value parameter contains the header contents.
    'foo2 = Value
    'If the header line being retrieved is a continuation header line, then the Field parameter contains "" (empty string).
End Sub

Private Sub NNTP1_Transfer(BytesTransferred As Long, Text As String)

    'The Text parameter contains the portion of article data being retrieved.
   ' foo1 = Text
    'The BytesTransferred parameter contains the number of bytes transferred since the beginning of the article, including header bytes.
   ' foo2 = BytesTransferred
End Sub









Private Sub NymReplyChange_Click()
    If Not gRemailerType = REMAILER_CYPHERPUNK Then
        MsgBox "Only Cypherpunk type remailers can be used for reply blocks.", 16, gPiStr
    Else
        gNymState = gNYMRPLYCHANGE
        frmMultiNyms.Show vbModal
        gNymState = gintNYM_IDLE
    End If
End Sub

Private Sub NymShow_Click()
    gShowNymStatus = 1
    frmNymServerStats.Show
    gShowNymStatus = 0
End Sub

Public Sub PGPAdd_Click()
Dim msg As String
If gPGPVersion = PGP26x Then
    If Not UpdatePublicKeysFile Then
        'DisablePGPMenuItems
        'PGPVersion_PGP26x.Checked = False
        'WriteProfile "PGP Info", "PGP Version", gPGPVersion
        Beep
        ShowStatus ("PGP 2.6.x is not functioning.")
    End If
Else
    msg = "Note.  For PGP 5 and 6 there is no longer any need to create the file PUBKEYS.OUT file. " & vbCrLf
    MsgBox msg, vbExclamation, "PGP Version Information."
End If
End Sub
Private Sub PGPAddKey_Click()
Dim ClipText As String
Dim cmd As String
Dim iResult As Integer
Dim Ecount As Integer
Dim FileNum As Integer
Dim foo As String

Dim i As Long
Dim bufferin As String
Dim spgperr As String * 256
Dim KeyProps As String
On Error GoTo AddKeyError

If gPGPVersion = PGP5x Then

    bufferin = String(Len(MessageArea.Text) + 1, Chr(0))
    bufferin = MessageArea.Text & Chr(0)
    KeyProps = String((Len(bufferin) / 256) * 512 + 1, Chr(0))
    
    ' keyprops takes either key id(s) or user id(s)
    ' and returns the key's properties
    
    i = spgpKeyImport(bufferin, KeyProps, Len(KeyProps))
    If Not i = 0 Then
        Beep
        Call spgpGetErrorString(i, spgperr)
        ShowStatus ("Error occurred importing keys.  Error code: " & spgperr)
       'Err.Raise 1000, "spgpKeyImport", spgperr
    Else
        ShowStatus ("Keys imported successfully...")
    End If
    

    ' parse the returned property-string into a TKey_Data record
   ' Key = ParseKeyData(KeyProps)
   
Else
    If Len(MessageArea) = 0 Then
        MessageArea.Text = "No key found ..."
        Exit Sub
    End If
    ClipText = MessageArea.Text
    FileNum = FreeFile
    If InStr(1, gPGPFile, ":") = 0 Then
        Open gPGPPath & "\" + gPGPFile & ".out" For Output As FileNum
    Else
        Open gPGPFile + ".out" For Output As FileNum
    End If
    Print #FileNum, ClipText
    Close #FileNum
    If InStr(1, gPGPFile, ":") = 0 Then
        cmd = gPGPPath & "\PGP -ka " & gPGPPath & "\" + gPGPFile & ".out"
    Else
        cmd = gPGPPath & "\PGP -ka " & gPGPFile & ".out"
    End If
    CheckLen (cmd)
    ExecCmd (cmd)
        
    If InStr(1, gPGPFile, ":") = 0 Then
        Kill gPGPPath & "\" + gPGPFile & ".out"
    Else
        Kill gPGPFile & ".out"
    End If
    UpdatePublicKeysFile
    If Len(MessageArea.Text) > 0 Then
        '---------------------------------------------
        'test to see if message area is empty
        '---------------------------------------------
        foo = MsgBox("Clear the message area?", vbYesNo, "Add Key Complete")
        If foo = vbYes Then
            MessageArea.Text = ""
        End If
    End If
End If
Exit Sub
AddKeyError:
    MsgBox Err.Description & " (in PGPAddKey)", vbCritical + vbApplicationModal, App.Title
    Err.Clear
End Sub


Private Sub PGPConvent_Click()
    PGPConvent.Checked = Not PGPConvent.Checked
End Sub
Private Sub PGPDecrypt_Click()

    Dim ClipText As String
    Dim FileNum As Integer
    Dim cmd As String
    Dim iResult As Integer
    Dim Ecount As Integer
    Dim TextLine As String
    Dim Cyphertext As String
    Dim TheFileName As String
    Dim foo As Integer
    Dim Key As TKey_Data
        
    On Error GoTo DecryptError
    
    If Not PGPFile.Checked Then
        If gPGPVersion = PGP5x Then
            Select Case spgpAnalyseMessage(MessageArea.Text)
                Case "Unknown"
                    MsgBox "The message area does not contain an encrypted message", vbApplicationModal + vbCritical, "Analyse Message"
                Case "Signed"
                    ShowStatus ("The message has been signed....")
                    spgpDecryptMessage
                Case "Encrypted"
                    spgpDecryptMessage
            End Select
            Exit Sub
        End If
        If gObscurity = 1 Then
            MessageArea.SelStart = 0
            MessageArea.SelText = "-----BEGIN PGP MESSAGE-----" & vbCrLf + "Version: unknown" + gCRLF + gCRLF
            MessageArea.SelStart = Len(MessageArea.Text)
            MessageArea.SelText = "-----END PGP MESSAGE-----" & vbCrLf
        End If
        If MessageArea.Text = "" Then
            Beep
            Exit Sub
        End If
        
            ClipText = MessageArea.Text
            FileNum = FreeFile
            If InStr(1, gPGPFile, ":") = 0 Then
                Open gPGPPath + "\" + gPGPFile + ".out" For Output As FileNum
            Else
                Open gPGPFile + ".out" For Output As FileNum
            End If
            Print #FileNum, ClipText
            Close #FileNum
            cmd = gPGPPath & "\pgp " & gPGPPath & "\" & gPGPFile & ".out "
            ExecCmd (cmd)
        
            FileNum = FreeFile
            If InStr(1, gPGPFile, ":") = 0 Then
                Open gPGPPath + "\" + gPGPFile For Input As FileNum
            Else
                Open gPGPFile For Input As FileNum
            End If
            While Not EOF(FileNum)
                Line Input #FileNum, TextLine
                Cyphertext = Cyphertext & TextLine & vbCrLf
            Wend
            Close #FileNum
        
     
        MessageArea.Text = Cyphertext
        If InStr(1, gPGPFile, ":") = 0 Then
            Kill gPGPPath + "\" + gPGPFile
        Else
            Kill gPGPFile
        End If
        If InStr(1, gPGPFile, ":") = 0 Then
            WipeFile (gPGPPath + "\" + gPGPFile)
        Else
            WipeFile (gPGPFile)
        End If
    
    Else
        CommonDialog1.DialogTitle = "Open file to decrypt"
        CommonDialog1.Flags = &H2& + &H4&
        CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc|All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path
        CommonDialog1.Action = 1
        TheFileName = CommonDialog1.FileName
        '---------------------------------------------
        'option to decrypt file to message area
        '---------------------------------------------
        foo = MsgBox("Would you like to put the output in the message window?", vbYesNo, "File Decrypt")
        If foo = vbYes Then
            '---------------------------------------------
            'check for text in the msg area
            '---------------------------------------------
            If Len(MessageArea) > 0 Then
                foo = MsgBox("Okay to clear the message area first?", vbYesNo, "Get Key From Server")
                If foo = vbYes Then
                    MessageArea.Text = ""
                Else
                    Exit Sub
                End If
            End If
            cmd = gPGPPath & "\pgp " + TheFileName + " -o " + gPGPPath + "\" + gPGPFile
            
            '---------------------------------------------
            'run the PGP decryption command
            '---------------------------------------------
            ExecCmd (cmd)
            
            '---------------------------------------------
            'load the text into the message box
            '---------------------------------------------
            FileNum = FreeFile
            If InStr(1, gPGPFile, ":") = 0 Then
                Open gPGPPath + "\" + gPGPFile For Input As FileNum
            Else
                Open gPGPFile For Input As FileNum
            End If
            While Not EOF(FileNum)
                Line Input #FileNum, TextLine
                Cyphertext = Cyphertext + TextLine + Chr$(13) + Chr$(10)
            Wend
            Close #FileNum
            MessageArea.Text = Cyphertext
            If InStr(1, gPGPFile, ":") = 0 Then
                Kill gPGPPath + "\" + gPGPFile
            Else
                Kill gPGPFile
            End If
            If InStr(1, gPGPFile, ":") = 0 Then
                WipeFile (gPGPPath + "\" + gPGPFile)
            Else
                WipeFile (gPGPFile)
            End If
        Else
            CommonDialog1.DialogTitle = "Save decrypted file as:"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 2
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            cmd = gPGPPath & "\pgp " + TheFileName + " -o " + CommonDialog1.FileName
            
            CheckLen (cmd)
            ExecCmd (cmd)
        End If
    End If
     
    Exit Sub

DecryptError:
    MsgBox "Could not decrypt or verify the message.  Following error was returned by PGP: " & Err.Description
    Err.Clear
End Sub

Private Sub PGPDeleteKey_Click()
Dim Storestring1 As String
Dim msg As String
Dim Storestring2 As String
Dim cmd As String
Dim Ecount As Integer
Dim iResult As Integer
If gPGPVersion = PGP5x Then
    msg = "Deletion of a Key-Pair is not available from this version of PGP. " & vbCrLf
    msg = msg & "It is better that you use the utilites that come with PGP to create your keys." & vbCrLf
    MsgBox msg, vbExclamation, "Create Key Pair Command"
Else
    gCancelButton = False
    gKeyID = ""
    Storestring1 = frmSelectUserID.Label2.Caption
    Storestring2 = frmSelectUserID.Label1.Caption
    frmSelectUserID.Label2.Caption = "Select which user's key to delete"
    frmSelectUserID.Label1.Caption = "from your public ring."
    frmSelectUserID.Show vbModal
    frmSelectUserID.Label2.Caption = Storestring1
    frmSelectUserID.Label1.Caption = Storestring2
    Unload frmSelectUserID
    If Not gCancelButton Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd = gPGPPath & "\PGP -kr " + " +force " + Chr$(34) + gKeyID + Chr$(34) + ""
        CheckLen (cmd$)
        ExecCmd (cmd$)
        
        UpdatePublicKeysFile
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
End If
End Sub
Private Sub PGPEncrypt_Click()
    gKeyID = txtTo.Text
    EncryptMessageArea
End Sub

Private Sub PGPEnSign_Click()

    Dim TheFileName As String
    Dim PGPCmdString As String
    Dim foo As String
    Dim FileNum As Integer
    
    
    On Error GoTo FileEnSError
    
        
    gCancelButton = False
    PGPCmdString = ""
    gPGPKeyID = ""
    TheFileName = ""
    
    '---------------------------------------------
    'handle case of saving a message to file, then encrypting
    '---------------------------------------------
    If PGPFile.Checked Then
        foo = MsgBox("Would you like to encrypt the message area to a file?", vbYesNo, "File Encrypt")
        If foo = vbYes Then
            CommonDialog1.DialogTitle = "Specify file for saving message"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            TheFileName = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            FileNum = FreeFile
            Open CommonDialog1.FileName For Output As FileNum
                Print #FileNum, MessageArea.Text
            Close #FileNum
        End If
    End If
    
    'Dim tempString As String
    'tempString = Form26.Label1.Caption
    'Form26.Label1.Caption = "Select the key to sign the message with."
    '---------------------------------------------
    'display the list of personal keys to sign with
    'if key is selected, it is in global gPGPKeyID
    '---------------------------------------------
    If gPGPVersion = PGP26x Then
        CheckMultipleKey
    Else
        vb2spgpContext.SelectPrivateKeys = True
        frmViewKeyRing.lblContext = "Select from this list or private keys which key you would like to sign the message with"
        frmViewKeyRing.Show vbModal
        If gPGPKeyID = "" Then
            Beep
            Exit Sub
        End If
        vb2spgpContext.Sign = 1
        vb2spgpContext.SignKeyID = gPGPKeyID
    End If
    '------------------------------------------------------------
    'Check to see if the user cancelled anywhere along the way
    '------------------------------------------------------------
    If Not gCancelButton Then
         
        '---------------------------------------------
        'is this an eyes, conventional, other encrypt?
        '---------------------------------------------
        If Not PGPFile.Checked Then
            'If PGPEyes.Checked = True Then
            '    PGPCmdString = "m"
            'End If
            vb2spgpContext.TextMode = 0
            vb2spgpContext.Armor = 1
            If PGPConvent.Checked = True Then
                vb2spgpContext.ConventionalEncrypt = 1
                EncryptMessage "", "-satcw" + PGPCmdString
            Else
                vb2spgpContext.ConventionalEncrypt = 0
                EncryptMessage txtTo.Text, "-satew" + PGPCmdString
            End If
        Else
        '---------------------------------------------
        'no, it is a file encrypt
        '---------------------------------------------
            If Len(TheFileName) = 0 Then
                '---------------------------------------------
                'user has not already specified a file name from above
                '---------------------------------------------
                CommonDialog1.DialogTitle = "Open file to encrypt"
                CommonDialog1.Flags = &H2& + &H4&
                CommonDialog1.Filter = "All Files (*.*)|*.*"
                CommonDialog1.FilterIndex = 1
                CommonDialog1.CancelError = True
                CommonDialog1.InitDir = App.Path
                CommonDialog1.Action = 1
                TheFileName = CommonDialog1.FileName
                If Not InStr(TheFileName, ".asc") = 0 Then
                    MsgBox "Please rename the file to not have the extension '.asc' as this will be the extension of the save file.", vbApplicationModal + vbCritical, "Check for filename extension."
                    gCancelButton = False
                    Exit Sub
                End If
                ChDrive Mid$(App.Path, 1, 3)
                ChDir App.Path
            End If
            If PGPEyes.Checked = True Then
                PGPCmdString = "m"
            End If
            vb2spgpContext.TextMode = 0
            vb2spgpContext.Armor = 1
            vb2spgpContext.Sign = 1
            vb2spgpContext.FileIn = TheFileName
            
            If PGPConvent.Checked Then
                vb2spgpContext.ConventionalEncrypt = 1
                EncryptFile "", "-scat" + PGPCmdString, TheFileName
            Else
                vb2spgpContext.ConventionalEncrypt = 0
                EncryptFile txtTo.Text, "-seat" + PGPCmdString, TheFileName
            End If
        End If
    Else
        '---------------------------------------------
        'user canceled
        '---------------------------------------------
        gCancelButton = False
    End If
 Exit Sub
FileEnSError:
        gCancelButton = False
        MsgBox "There was an error.  The reason given by the operating system is: " & Err.Description, vbApplicationModal, App.Title
        Err.Clear
End Sub
Private Sub PGPEyes_Click()

    PGPEyes.Checked = Not PGPEyes.Checked
End Sub
Private Sub PGPFile_Click()
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
    Dim foo As String
    Dim SectionName As String
    
    On Error GoTo GetKeyError
    If CheckConnection Then
        If Len(MessageArea) > 0 Then
            '---------------------------------------------
            'test first to see if message area is empty
            '---------------------------------------------
            foo = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Get Key From Server")
            If foo = vbYes Then
                MessageArea.Text = ""
            Else
                Exit Sub
            End If
        End If
        If Len(txtTo) = 0 Then
            '---------------------------------------------
            'the user did not specify a recipient in the "to:" box
            '---------------------------------------------
            MessageArea.Text = "No user specified in the To: box." & vbCrLf
            MessageArea.Text = MessageArea.Text + "Please enter a valid e-mail address," & vbCrLf
            MessageArea.Text = MessageArea.Text + "or click on the right arrow button in the" & vbCrLf
            MessageArea.Text = MessageArea.Text + "To: box to choose a name from your address book."
            Exit Sub
        End If
        gServerState = HTTPSTATE
        gWebState = GETSERVERKEY
        If Not frmPI.HTTP1.WinsockLoaded Then frmPI.HTTP1.WinsockLoaded = True
        
        
        If gGetKeyURL = "" Then
            SectionName = "Net Info"
            gGetKeyURL = ReadProfile(SectionName, "GetKeyURL")
            If gGetKeyURL = "" Then
            'frmSelectKeyServer.Show vbModal
            gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
            SectionName = "Net Info"
            WriteProfile SectionName, "GetKeyURL", gGetKeyURL
            'gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
            'WriteProfile SectionName, "GetKeyURL", gGetKeyURL
            End If
        End If
        ShowStatus ("")
        ShowStatus ("Requesting key from server.  Please wait.  Hint: Use 'Add key from message' when done.")
        DoEvents
        GetWebURL (gGetKeyURL & frmPI.txtTo.Text)
        gServerState = 0
    End If
    Exit Sub
GetKeyError:
    HideStatus
    MsgBox Err.Description & " (in PGPGetKey)"
    gServerState = 0
    Err.Clear
End Sub

Private Sub PGPInsertKey_Click()
    InsertKey (IDKEY)
End Sub

Private Sub PGPMin_Click()

    PGPMin.Checked = Not PGPMin.Checked
    If PGPMin.Checked = True Then
        gMinState = 2
    Else
        gMinState = 1
    End If
End Sub

Private Sub PGPMultiple_Click()
    PGPMultiple.Checked = Not PGPMultiple.Checked
End Sub

Private Sub PGPObscurity_Click()

    If PGPObscurity.Checked = True Then
        PGPObscurity.Checked = False
        gEncryptToRemailer = True
        gObscurity = 0
    Else
        PGPObscurity.Checked = True
        gEncryptToRemailer = True
        gObscurity = 1
    End If
End Sub

Private Sub PGPOptions_Click()
    frmPGPOptions.Show vbModal
End Sub

Private Sub PGPSelf_Click()
    PGPSelf.Checked = Not PGPSelf.Checked
End Sub
Private Sub PGPClearSign_Click()

    Dim PGPMutlipleToggle As Boolean
    Dim PGPEncryptSelfToggle As Boolean
    Dim PGPObscurityToggle As Boolean
    Dim PGPCmdString As String
    Dim PGPMultipleToggle As Boolean
    Dim TheFileName As String
    Dim tmpstr2 As String
    'Dim tmpstr As String
    
    
    On Error GoTo ClearSignError
    '------------
    'Initialise Context
    '-------------
    vb2spgpContext.Initialise
    '---------------------------------------------
    'first verify there is something to encrypt
    '---------------------------------------------
    If (Len(MessageArea) = 0) And Not PGPFile.Checked Then
        ShowStatus ("Nothing to sign or encrypt!")
        MessageArea.SelStart = 0
        MessageArea.SelLength = Len(MessageArea)
        Beep
        Exit Sub
    End If
    
    PGPCmdString = ""
    gPGPKeyID = ""
    gCancelButton = False
    
    If gPGPVersion = PGP26x Then
        tmpstr2 = Form26.Label1.Caption
        If PGPFile.Checked Then
            Form26.Label1.Caption = "Select the key to sign the file with."
        Else
            Form26.Label1.Caption = "Select the key to sign the message with."
        End If
        '---------------------------------------------
        'display the list of personal keys to sign with
        'if key is selected, it is in global gPGPKeyID
        '---------------------------------------------
        CheckMultipleKey
    Else
        vb2spgpContext.SelectPrivateKeys = True
        frmViewKeyRing.lblContext = "Select from this list or private keys which key you would like to sign the message with."
        frmViewKeyRing.Show vbModal
        vb2spgpContext.SignKeyID = gPGPKeyID
    End If
    If Not gCancelButton Then
            '---------------------------------------------
            'user has not canceled when selecting the keys.
            '---------------------------------------------
            'tmpstrorarily disable encrypting with multiple keys
            '---------------------------------------------
            PGPMultipleToggle = False
            If PGPMultiple.Checked Then
                PGPMultipleToggle = True
                PGPMultiple.Checked = False
            End If
        
            '---------------------------------------------
            'tmpstrorarily disable encrypting to self
            '---------------------------------------------
            PGPEncryptSelfToggle = False
            If PGPSelf.Checked Then
                PGPEncryptSelfToggle = True
                PGPSelf.Checked = False
            End If
        
            '---------------------------------------------
            'tmpstrorarily disable gObscurity
            '---------------------------------------------
            PGPObscurityToggle = False
            If PGPObscurity.Checked Then
                PGPObscurityToggle = True
                PGPObscurity.Checked = False
            End If
        
            If Form26.Label1.Caption = "" Then
                Form26.Label1.Caption = tmpstr2
                Exit Sub
            End If
        
           ' tmpstr = txtTo.Text
        '---------------------------------------------
        'is this a basic message clear sign
        '---------------------------------------------
        vb2spgpContext.Clear = 1
        vb2spgpContext.Sign = 1
        vb2spgpContext.Armor = 1
        vb2spgpContext.TextMode = 0
        If Not PGPFile.Checked Then
            If PGPConvent.Checked Then
                vb2spgpContext.ConventionalEncrypt = 1
                EncryptMessage "", "-stac"
            Else
                vb2spgpContext.ConventionalEncrypt = 0
                spgpEncryptMessage
                'EncryptMessage txtTo.Text, "-sta"
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
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            TheFileName = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            
            vb2spgpContext.Clear = 1
            vb2spgpContext.Armor = 1
            vb2spgpContext.TextMode = 0
            vb2spgpContext.FileIn = TheFileName
            vb2spgpContext.FileOut = StripExt(TheFileName) & ".asc"
            If PGPConvent.Checked Then
                vb2spgpContext.ConventionalEncrypt = 1
                EncryptFile "", "-stac", TheFileName
            Else
                vb2spgpContext.ConventionalEncrypt = 0
                EncryptFile txtTo.Text, "-sta", TheFileName
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
        're-enable gObscurity
        '---------------------------------------------
        If PGPObscurityToggle Then
            PGPObscurity.Checked = True
        End If
        Exit Sub
    Else
        '---------------------------------------------
        'user canceled at form26
        '---------------------------------------------
        gCancelButton = False
        Exit Sub
    End If
    
ClearSignError:
    MsgBox Err.Description & " -> PGPClearSign. ", vbApplicationModal '& errStr$, 48, gPiStr
    Err.Clear
End Sub



Private Sub PGPVerify_Click()

    Dim tmpstr1 As String
    Dim tmpstr2 As String
    Dim tmpstr3 As String
    Dim tmpstr4 As String
    Dim tmpstr5 As String
    Dim FileNum As Integer
    Dim cmd As String
    Dim Ecount As Integer
    Dim iResult As Integer
    Dim pp As String
    Dim pix As String
    Dim Cyphertext As String
    Dim TextLine As String
    
    On Error GoTo PGPVerifyError
    
    If Len(MessageArea.Text) > 0 Then
        '---------------------------------------------
        'test first to see if message area is empty
        '---------------------------------------------
        tmpstr1 = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Update Remailer")
        If tmpstr1 = vbYes Then
            MessageArea.Text = ""
        Else
            Exit Sub
        End If
    End If
    
    MessageArea.SelText = "Running 3-Step PGP verification." + gCRLF + gCRLF
    MessageArea.SelText = "This will check the signature on your pgp262i.zip file" + gCRLF
    MessageArea.SelText = "to see if it was signed by Jeffrey Schiller at MIT." + gCRLF
    
    tmpstr2 = gPGPPath + "\PGP262I.ASC"
    tmpstr3 = gPGPPath + "\PGP262I.ZIP"
    tmpstr4 = App.Path + "\JKEY.ASC"
    tmpstr5 = gPGPPath + "pubring.pgp"
    
    If iFileExists(tmpstr2) And iFileExists(tmpstr3) And iFileExists(tmpstr4) Then
        
        cmd = gPGPPath & "\pgp -ka +batchmode " + App.Path + "\jkey.asc"
        ExecCmd (cmd)
        
        cmd = gPGPPath & "\pgp " + tmpstr2 + " " + tmpstr3 + "  > " + gPGPPath + "\zcghfoo"
        ExecCmd (cmd)
        
        cmd = gPGPPath & "\pgp -kvc " + Chr$(34) + "Jeffrey I. Schiller" + Chr$(34) + " " + pp + "\pubring.pgp >> " + gPGPPath + "\zcghfoo"
        ExecCmd (cmd)
        
        Dim Okay As Boolean
        'step4:  process the zcghfoo file
        FileNum = FreeFile
        Open gPGPPath + "\" + "zcghfoo" For Input As FileNum
        Cyphertext = ""
        Okay = False
        While Not EOF(FileNum)
            Line Input #FileNum, TextLine
            If InStr(1, TextLine, "DD DC 88 AA 92 DC DD D5  BA 0A 6B 59 C1 65 AD 01", 0) Then
                MessageArea.Text = ""
                MessageArea.SelText = "Your pgp262i.zip file has a valid signature on it." + gCRLF
                MessageArea.SelText = "The file was signed by Jeffrey I. Schiller at MIT." + gCRLF
                MessageArea.SelText = "This is not a 100% certification." + gCRLF
                MessageArea.SelText = "Reason:" + gCRLF
                MessageArea.SelText = "1.  A hacker could have tampered with the PGP source code" + gCRLF
                MessageArea.SelText = "    so that the pgp command outputs that this signature" + gCRLF
                MessageArea.SelText = "    is okay, when in reality, it is not Jeffrey's signature." + gCRLF
                MessageArea.SelText = "Possible Defences:" + gCRLF
                MessageArea.SelText = "1.  Examine the PGP source code and build your own executable (best defense)." + gCRLF
                Okay = True
                Exit Sub
            End If
        Wend
        Close FileNum
        If Not Okay Then
            MessageArea.Text = ""
            MessageArea.Text = "A problem with your PGP installation has been found" + gCRLF
            MessageArea.SelText = "The signature on the pgp262i.zip file is not valid." + gCRLF
            MessageArea.SelText = "Use the send feedback command to get help with this." + gCRLF
        End If
    Else
        MessageArea.Text = ""
        If Not iFileExists(tmpstr2) Then
            MessageArea.Text = "You do not have a copy of pgp262i.asc in your pgp directory." + gCRLF
        End If
        If Not iFileExists(tmpstr3) Then
            MessageArea.SelText = "You do not have a copy of pgp262i.zip in your pgp directory." + gCRLF
        End If
        If Not iFileExists(tmpstr4) Then
            MessageArea.SelText = "You do not have a copy of jkey.asc in your Private Idaho directory." + gCRLF
        End If
        If Not iFileExists(tmpstr5) Then
            MessageArea.SelText = "You do not have the pubring.pgp file in your pgp directory." + gCRLF
        End If
        MessageArea.SelText = "PGP verification cannot complete without the required file(s)." + gCRLF
    End If
   
Exit Sub
    
PGPVerifyError:
    Resume Next

End Sub

Private Sub PGPVersion_PGP26x_Click()
Dim msg As String
Dim SectionName As String

If Len(App.Path & gPGPPath) > 30 Then
    msg = "The application and PGP26 paths appear to be incompatible with DOS and PGP2.6.x" & vbCrLf
    msg = msg & "It is suggested that if you want to use PGP2.6.x that your either re-install PGP 2.6.x into a shorter directory, " & vbCrLf
    msg = msg & "or re-install PI into a shorter directory."
    MsgBox msg, vbCritical, "PGP Installation Warning"
End If

PGPVersion_PGP26x.Checked = True
PGPVersion_PGP5x.Checked = False
gPGPVersion = PGP26x

MsgBox "Private Idaho needs to confirm the PGP 2.6.x paths are correct.", 64, gPiStr
frmPGPOptions.Command1.Enabled = False
frmPGPOptions.Show vbModal
DoEvents
If UpdatePublicKeysFile Then
    EnablePGPMenuItems
Else
    gPGPVersion = NoPGP
    DisablePGPMenuItems
    PGPVersion_PGP26x.Checked = False
    PGPVersion_PGP5x.Checked = False
    gPGPVersion = NoPGP
End If
SectionName = "PGP Info"
WriteProfile SectionName, "PGP Version", gPGPVersion
'PGPVersion_PGP5x.Checked = True
'Enable 2.6.x functions
'PGPAdd.Enabled = True
'End If
End Sub

Private Sub PGPVersion_PGP5x_Click()
Dim SectionName As String
PGPVersion_PGP26x.Checked = False
PGPVersion_PGP5x.Checked = True
gPGPVersion = PGP5x
SectionName = "PGP Info"
WriteProfile SectionName, "PGP Version", gPGPVersion
'Disable 2.6.x functions
'PGPAdd.Enabled = False
EnablePGPMenuItems
End Sub

Private Sub PGPWrap_Click()
    PGPWrap.Checked = Not PGPWrap.Checked
End Sub

Private Sub Prepare_Usenet_Nym_Click()
frmUSENETGateways.Show vbModal
If txtTo.Text = "" Then
        ShowStatus ("You need to specify a recipient or have something to send.")
        Beep
        DoEvents
        Exit Sub
    End If
gNymState = gNYM_USENET_PREPARE
frmMultiNyms.Show vbModal
gNymState = gintNYM_IDLE
End Sub
Private Sub Prepare_usenet_standard_Click()
'Dim tmpstr1 As String
'Dim i As Integer
mSendNewsGroupMessage.Enabled = True
frmUSENETGateways.Show vbModal

'PrepareUSENETMessage (gEmailAddress) 'Pass from argument
End Sub

Private Sub PrintSetup_Click()
CommonDialog1.Flags = &H40&
CommonDialog1.Action = 5
End Sub

Private Sub RemailerAppend_Click()
    
    'gEncryptToRemailer = False
    'USENETFi.Checked = False
    'UseNetSoda.Checked = False
    'USENETGate.Checked = False
    'USENETNone.Checked = True
    gNewsgroupType = 0
    'TransferAES.Checked = True
End Sub

Public Sub RemailerKeys()
    Dim tmpstr As String
    
    On Error GoTo RemailerKeysError
    
    If Len(MessageArea.Text) > 0 Then
        '---------------------------------------------
        'test first to see if message area is empty
        '---------------------------------------------
        tmpstr = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Fetch Remailer Keys")
        If tmpstr = vbYes Then
            MessageArea.Text = ""
        Else
            Exit Sub
        End If
    End If
    
    If CheckConnection Then
        gWebState = GETREMAILERKEYS
        If Not frmPI.HTTP1.WinsockLoaded Then frmPI.HTTP1.WinsockLoaded = True
        ShowStatus ("Getting current remailer PGP keys.  Please wait.  Hint: Use 'Add key from message' when done.")
        DoEvents
        If gPGPKeysURL = "" Then
            Dim SectionName As String
            SectionName = "Net Info"
            gPGPKeysURL = ReadProfile(SectionName, "PGPKeysURL")
            If gPGPKeysURL = "" Then
                gPGPKeysURL = "http://kiwi.cs.berkeley.edu/pgpkeys"
                WriteProfile SectionName, "PGPKeysURL", gPGPKeysURL
            End If
        End If
        DoEvents
        GetWebURL (gPGPKeysURL)
        HideStatus
        DoEvents
    End If
    Exit Sub
RemailerKeysError:
    HideStatus
    MsgBox Err.Description + "RemailerKeys"
    Err.Clear
End Sub




Private Sub RemailersCP_Click()

 UseCypherPunk
End Sub

Private Sub RemailersMix_Click()
UseMixmaster
End Sub

Public Sub RemailerUpdate()
Dim Response As String
    On Error GoTo RemailerError
    
    If Len(MessageArea.Text) > 0 Then
        '---------------------------------------------
        'test first to see if message area is empty
        '---------------------------------------------
        Response = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Update Remailer")
        If Response = vbYes Then
            MessageArea.Text = ""
        Else
            Exit Sub
        End If
    End If
    
    If CheckConnection Then
        gServerState = HTTPSTATE
        '---------------------------------------------
        'fetch the URL for obtaining remailer data
        '---------------------------------------------
       ' If gRemailerInfoURL = "" Then
            Dim SectionName As String
            SectionName = "Net Info"
            If gRemailerType = REMAILER_MIX Then
                gMixListURL = ReadProfile(SectionName, "MixListURL")
                If Len(gMixListURL) = 0 Then
                    gMixListURL = "http://www.publius.net/mixmaster-list"
                    WriteProfile SectionName, "MixListURL", gMixListURL
                End If
                gMixType2URL = ReadProfile(SectionName, "MixType2URL")
                If Len(gMixType2URL) = 0 Then
                    gMixType2URL = "http://www.publius.net/type2.list"
                    WriteProfile SectionName, "MixType2URL", gMixType2URL
                End If
                gMixPubRingURL = ReadProfile(SectionName, "MixPubRingURL")
                If Len(gMixPubRingURL) = 0 Then
                    gMixPubRingURL = "http://www.publius.net/pubring.mix"
                    WriteProfile SectionName, "MixPubRingURL", gMixPubRingURL
                End If
            Else
                gRemailerInfoURL = ReadProfile(SectionName, "RemailerInfoURL")
                If Len(gRemailerInfoURL) = 0 Then
                    gRemailerInfoURL = "http://www.publius.net/rlist"
                    WriteProfile SectionName, "RemailerInfoURL", gRemailerInfoURL
                End If
            End If
        'End If
        '---------------------------------------------
        'mixmaster option selected on menu
        '---------------------------------------------
        If gRemailerType = REMAILER_MIX Then
            gWebState = MIXUPDATE
            If Not frmPI.HTTP1.WinsockLoaded Then frmPI.HTTP1.WinsockLoaded = True
            frmRemailerList.lblstatus = "Downloading and updating Mixmaster remailer information.  Please wait."
            DoEvents
            GetWebURL (gMixListURL)
            gWebState = TYPE2UPDATE
            If Not frmPI.HTTP1.WinsockLoaded Then frmPI.HTTP1.WinsockLoaded = True
            frmRemailerList.lblstatus = "Downloading and updating the Mixmaster Type2.lis service.  Please wait."
            DoEvents
            GetWebURL (gMixType2URL)
            gWebState = PUBRINGUPDATE
            If Not frmPI.HTTP1.WinsockLoaded Then frmPI.HTTP1.WinsockLoaded = True
            frmRemailerList.lblstatus = "Downloading and updating the Mixmaster Pubring.mix file.  Please wait."
            DoEvents
            GetWebURL (gMixPubRingURL)
        Else
            '---------------------------------------------
            'mixmaster is not checked on the menu
            '---------------------------------------------
            gWebState = GETREMAILERUPDATE
            If Not frmPI.HTTP1.WinsockLoaded Then frmPI.HTTP1.WinsockLoaded = True
            frmRemailerList.lblstatus = "Downloading and remailer info..."
            DoEvents
            GetWebURL (gRemailerInfoURL)
        End If
        If gCancelButton Then
            gCancelButton = False
            Exit Sub
        End If
       
        
        'And do the private stuff as well
        If iFileExists(App.Path & "\private.txt") Then
            frmRemailerList.InitializeRemailers (App.Path & "\private.txt")
        End If
        '---------------------------------------------
        'set the newsgroup menu for cp state
        '---------------------------------------------
        'USENETGate.Visible = True
       ' USENETFi.Visible = False
        'UseNetSoda.Visible = False
        '------------------------------
        ' This will sort the remailers as well and fill the matched remailers list
        '  (strange place to put it...need to fix
        '-----------------------------------------
        'SortRemailers
        'FillRemailerList
    End If
    Exit Sub
RemailerError:
    HideStatus
    MsgBox Err.Description & "(in RemailerUpdate)"
    gServerState = 0
    Err.Clear
End Sub

Private Sub SelectKeyServer_Click()
frmSelectKeyServer.Show
End Sub

Private Sub SendSysInfo_Click()
   PrepareFeedback
End Sub


Private Sub SMTP1_Error(ErrorCode As Integer, Description As String)


    MsgBox "Error" & Str$(ErrorCode) + ". Mail not sent.", vbApplicationModal, App.Title
End Sub


Private Sub SMTP1_PITrail(Direction As Integer, Message As String)


   ' ShowStatus = Message
    If gSMTPLog = 1 Then
        Print #gSMTPFile, "SMTP: " + Format$(Direction) + ": " + Message
    End If
End Sub
Private Sub SMTP1_StartTransfer()


    'ShowStatus = "Sending mail."
End Sub

Private Sub SMTP1_Transfer(BytesTransfered As Long)
    frmStatus.Label3.Caption = Format(BytesTransfered)
End Sub

Private Sub SSRibbon1_Click(Index As Integer, Value As Integer)
SSRibbon1(Index).Value = False

'StatusBar.Item(0).Style = sbrSimple
DoEvents
If Value = 0 Then
    ShowStatus ("")
    Select Case Index
        Case 0
           ' SSRibbon1(index).Picture = SSRibbon1(21).Picture
            SendPIMessage
            Unload Me
        Case 1
            EditPerform WM_COPY
            
        Case 2
            EditPerform WM_PASTE 'Win.EditPaste  'MessageArea.SelText = Win.EditPaste 'Clipboard.GetText()
        Case 3
            txtTo.Text = ""
            txtSubject.Text = ""
            txtCC.Text = ""
            MessageArea.Text = ""
            
        Case 4
            ImportFile
        Case 5
            SaveMessage
        Case 6
            If Not gRemailerType = REMAILER_NONE Then AppendInfo
        Case 7
            AddAttachment
        Case 8
            EditPerform WM_CUT
        Case 9
            EncryptMessageArea
        Case 10
            frmMain.ReplyToSender
    End Select
End If
End Sub



Private Sub StepAddKey_Click()

    Form32.Caption = "Adding a public key"
    Form32.Text1.Text = "Before you can send someone an encrypted message, you need a copy of their public key.  To add a copy of your key:" + gCRLF + gCRLF + "1. - Copy the public key (either from an e-mail message or key server) and paste it into the Message text box." + gCRLF + gCRLF + "2. - From the Keys menu, choose the 'Add key from message' command." + gCRLF + gCRLF + "3. - PGP will run in the DOS window.  If you are running Windows 95, click the icon in the taskbar.  Certify the key." + gCRLF + gCRLF + "The key is added to your public key ring."
    Form32.Show
End Sub

Private Sub StepAttach_Click()

    Form32.Caption = "Sending an attachment"
    Form32.Text1.Text = "You can attach a file to a message sent from Private Idaho.  To send a message with an attachment:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - Make sure you have a connection to the Internet if you are using the 16-bit version." + gCRLF + gCRLF + "4. - Check the Attachment checkbox and specify the file to attach." + gCRLF + gCRLF + "5. - In the drop-down box, specify if you'd like the file encrypted, you may also want to encrypt the text message before you send it." + gCRLF + gCRLF + "6. - Select the Send button (or choose 'Send' from the Message menu)." + gCRLF + gCRLF + "Note: The file is Base64 encoded to a MIME compliant attachment." + gCRLF + gCRLF + "Private Idaho currently doesn't support sending attachments through remailers."
    Form32.Show
End Sub

Private Sub StepCreateKey_Click()

    Form32.Caption = "Creating a PGP key pair"
    Form32.Text1.Text = "If you'd like to create a PGP secret and public key to use with a gNym:" + gCRLF + gCRLF + "1. - From the Keys menu, choose the 'Create key pair' command." + gCRLF + gCRLF + "2. - PGP will run in the DOS window.  Follow the steps for creating a key." + gCRLF + gCRLF + "Hint: Use a key size of 1024 bits or higher."
    Form32.Show
End Sub

Private Sub StepDecrypt_Click()

    Form32.Caption = "Decrypting a message"
    Form32.Text1.Text = "To decrypt a PGP message you've received:" + gCRLF + gCRLF + "1. - Copy the message from your e-mail application." + gCRLF + gCRLF + "2. - Paste it into Private Idaho's Message text box." + gCRLF + gCRLF + "3. - In the PGP menu, choose the 'Decrypt message' command." + gCRLF + gCRLF + "4. - PGP will run in the DOS window.  Enter your PGP passphrase." + gCRLF + gCRLF + "The encrypted text in the Message box is replaced by the decrypted text."
    Form32.Show
End Sub

Private Sub StepDelete_Click()

    Form32.Caption = "Deleting a public key"
    Form32.Text1.Text = "To remove a key from your public key ring:" + gCRLF + gCRLF + "1. - From the Keys menu, choose the 'Delete key' command." + gCRLF + gCRLF + "2. - Select the key to remove and click OK." + gCRLF + gCRLF + "3. - PGP will run in the DOS window.  If you are running Windows 95, click the icon in the taskbar.  Verify you want to remove the key." + gCRLF + gCRLF + "The key is removed from your public key ring."
    Form32.Show
End Sub

Private Sub StepEncrypt_Click()

    Form32.Caption = "Encrypting a message"
    Form32.Text1.Text = "To PGP encrypt a message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the PGP user ID of person you'll be sending the message to.  This is usually an e-mail address.  You must have already added the person's key to your public key ring." + gCRLF + gCRLF + "Hint: You can use the Address book in the File menu to store common user IDs and addresses.  You can access these by clicking on the button on the right side of the To: box.  Another way is to click to the left of the To: field where it says Click Here.  This will bring up a copy of the recipients on your public key ring." + gCRLF + gCRLF + "3. - From the PGP menu, choose the 'Encrypt message' command." + gCRLF + gCRLF + "PGP will run in the DOS window.  If you are running Windows 95, you will need to double click on the PGP icon in the task bar." + gCRLF + gCRLF + "The original text in the Message box is replaced by the encrypted text."
    Form32.Show
End Sub

Private Sub StepGetKey_Click()

    Form32.Caption = "Sending/getting MIT server keys"
    Form32.Text1.Text = "MIT has an experimental PGP key server where you can publish and retrieve keys.  Private Idaho can directly access the server without going to the Web page.  To request a copy of someone's key:" + gCRLF + gCRLF + "1. - Make sure you have an Internet connection (only applies to the 16-bit version)." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person's key you want to retrieve." + gCRLF + gCRLF + "3. - From the Keys menu, choose the 'Get key from server' command." + gCRLF + gCRLF + "4. - Private Idaho will connect with the server.  If the key is in the data base, it will be displayed in the Message text box.  You can then add the key to your public key ring with the 'Add key from message' command." + gCRLF + gCRLF + "To add a key to the server, use the 'Insert key in message' command in the Keys menu to insert your key.  Then choose the 'Submit key to server' command."
    Form32.Show
End Sub

Private Sub StepInfo_Click()


    Form32.Caption = "Internet privacy info"
    Form32.Text1.Text = "Private Idaho has links to a variety of Internet privacy sources.  To access them:" + gCRLF + gCRLF + "1. - Ensure Private Idaho can communicate with your Web browser.  The default is Netscape Navigator.  If you're using another browser, choose the Options command in the Web menu." + gCRLF + gCRLF + "2. - You should be connected to the Internet with the browser running and not minimized." + gCRLF + gCRLF + "3. - From the Web menu, choose the information you'd like to access." + gCRLF + gCRLF + "Your browser will display the associated Web page."
    Form32.Show
End Sub

Private Sub StepNym_Click()
   Form32.Caption = "Creating a Nym"
    Form32.Text1.Text = "A Nym(as in ano'Nym'ous) is an alias used for private communications.  Once you've created a gNym account, you can send messages through it to people.  Unlike anonymous remailers, they can reply back to you without knowing your identity.  Various free servers are available for setting up gNym accounts.  These are much more secure than using anon.penet.fi.  To create a gNym:" + gCRLF + gCRLF + "1. - From the gNym menu, choose the 'Create gNym' command." + gCRLF + gCRLF + "This command steps you through the entire gNym creation process with a series of easy to follow dialog boxes." + gCRLF + gCRLF + "Refer to the on-line help for additional information."
    Form32.Show
End Sub

Private Sub StepNymDelete_Click()

    Form32.Caption = "Deleting a Nym"
    Form32.Text1.Text = "To delete a Nym:" + gCRLF + gCRLF + "1. - From the Nym menu, choose the 'Delete Nym' command." + gCRLF + gCRLF + "2. - Select the Nym account to delete and click OK." + gCRLF + gCRLF + "3. - Send the prepared message to the Nym server.  Your Nym will be deleted."
    Form32.Show
End Sub

Private Sub StepNymPass_Click()

    Form32.Caption = "Changing a Nym password"
    Form32.Text1.Text = "If you want to change your Nym password:" + gCRLF + gCRLF + "1. - From the Nym menu, choose the 'Change Nym password' command." + gCRLF + gCRLF + "2. - Select the gNym to change and click OK." + gCRLF + gCRLF + "3. - In the Message text box enter your old and new passwords." + gCRLF + gCRLF + "4. - From the gNym menu, choose the 'Encrypt Nym message' command." + gCRLF + gCRLF + "5. Send the message."
    Form32.Show
End Sub

Private Sub StepNymReply_Click()


    Form32.Caption = "Changing a Nym reply block"
    Form32.Text1.Text = "If you want to change the routing of your Nym messages:" + gCRLF + gCRLF + "1. - Enter the final destination e-mail address in the To: text box." + gCRLF + gCRLF + "2. - Specify the new remailer or chain in the Remailer drop-down list." + gCRLF + gCRLF + "3. - From the Nym menu, choose the 'Change reply block' command." + gCRLF + gCRLF + "4. - Select the Nym you want to change." + gCRLF + gCRLF + "5. - Enter your Nym password in the Message text box after Password:.  Make sure the reply block is correct." + gCRLF + gCRLF + "6. - For alias type nyms, from the gNym menu, choose the 'Encrypt Nym message' command." + gCRLF + gCRLF + "7. - Send the message."
    Form32.Show
End Sub

Private Sub StepNymSend_Click()


    Form32.Caption = "Sending a Nym message"
    Form32.Text1.Text = "Once you've created a gNym account, you can send messages through it.  To do so:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - From the gNym menu, choose the 'Prepare gNym message' command." + gCRLF + gCRLF + "4. - Select the gNym account to use and click OK." + gCRLF + gCRLF + "5. - For alias type nyms, from the gNym menu, choose the 'Encrypt gNym message' command." + gCRLF + gCRLF + "7. Send the message."
    Form32.Show
End Sub

Private Sub StepRemailer_Click()


    Form32.Caption = "Anonymous messages"
    Form32.Text1.Text = "You can send some a message without revealing your identity by using an anonymous remailer.  To send an anonymous message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - From the Remailers menu, choose the type of remailer to use." + gCRLF + gCRLF + "4. - Select the remailer to send the message through from the Remailer drop-down list.  Selecting 'chain' routes the message through several remailers." + gCRLF + gCRLF + "5. - From the Message menu, choose the 'Append info' command.  This formats the message for sending through a remailer.  If you selected 'chain,' a dialog box will prompt you for the remailers to use." + gCRLF + gCRLF + "6. - Send the message." + gCRLF + gCRLF + "Note: The Cypherpunk type remailers support a variety of advanced features.  Refer to the on-line help."
    Form32.Show
End Sub

Private Sub StepSend_Click()


    Form32.Caption = "Sending a message from Private Idaho"
    Form32.Text1.Text = "If you have a connection to the Internet you can send a message directly from Private Idaho.  To send a message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - Make sure you have a connection to the Internet." + gCRLF + gCRLF + "4. - Select the Send button (or choose 'Send' from the Message menu)." + gCRLF + gCRLF + "Note: Before sending a message, you need to provide information about your e-mail server.  From the File menu, choose the 'Options' command." + gCRLF + gCRLF + "If you can't send e-mail directly from Private Idaho, you can easily transfer the message to an e-mail application such as Eudora or Pegasus.  Refer to the on-line help."
    Form32.Show
End Sub

Private Sub StepSendKey_Click()

    Form32.Caption = "Sending your public key"
    Form32.Text1.Text = "Before someone can send you an encrypted message, they need a copy of your public key.  To send a copy of your key:" + gCRLF + gCRLF + "1. - From the Keys menu, choose the 'Insert key in message' command." + gCRLF + gCRLF + "2. - Select your key from the user ID dialog box and click OK." + gCRLF + gCRLF + "3. - PGP will run and fetch the key from your public key ring." + gCRLF + gCRLF + "4. - The public key is inserted in the Message text box.  You can now send the key to someone you want to privately correspond with."
    Form32.Show
End Sub

Private Sub StepSign_Click()


    Form32.Caption = "Signing a message"
    Form32.Text1.Text = "There are two ways to sign a message.  One way is to clear-sign a message.  This method leaves the text alone, but wraps it with a signature.  Signing the message must be the last step as the signature depends on the contents of the message.  The second way is to sign an encrypted file.  Either way, the intent is to let the recipient know that you are the one who created the message because it requires your secret key and password to compute the signature." + gCRLF + gCRLF + "To sign a message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - From the PGP menu, choose the 'Clear sign message' or 'Encrypt and Sign' command." + gCRLF + gCRLF + "3. - PGP will run in the DOS window.  Enter your passphrase." + gCRLF + gCRLF + "A signature is attached to the message."
    Form32.Show
End Sub

Private Sub StepUpdateInfo_Click()
    Form32.Caption = "Updating remailer info"
    Form32.Text1.Text = "To update remailer info:" + gCRLF + gCRLF + "1. - Make sure you have a Net connection." + gCRLF + gCRLF + "2. - From the Remailers menu, choose the 'Update remailer info' command." + gCRLF + gCRLF + "Private Idaho will connect to Raph's Web page and download the current remailer data and update the remailer list box.  If Cypherpunk is checked in the Remailer menu, Cypherpunk info is updated.  If Mixmaster is checked, Mixmaster info is updated."
    Form32.Show
End Sub

Private Sub StepUSENET_Click()
Dim zap As String
    Form32.Caption = "USENET articles"
    zap = "You can post a USENET article without revealing your identity by using a remailer. "
    zap = zap & "To post an anonymous article:" + gCRLF + gCRLF
    zap = zap & "1. - Compose the message in the Message text box." + gCRLF + gCRLF
    zap = zap & "2. - Click on 'Newsgroups' and fill in the required info. " & vbCrLf & vbCrLf
    zap = zap & "3. - Click on 'Prepare USENET Message' to creat the Nym message (you would have previously created a Nym." & vbCrLf & vbCrLf
    zap = zap & "4. - Send using a remailer"
    Form32.Text1.Text = zap
    Form32.Show
End Sub

Private Sub StepWeb_Click()


    Form32.Caption = "Anonymous Web page access"
    Form32.Text1.Text = "It's very easy for someone to log Web pages you visit.  "
    Form32.Text1.Text = Form32.Text1.Text & " Community Connexion (c2.org), has a free Web anonymizer service. "
    Form32.Text1.Text = Form32.Text1.Text & "To use it from Private Idaho:" + gCRLF + gCRLF + "1. - Type the Web page URL you want to anonymously visit in the Message text box and select (highlight) it." + gCRLF + gCRLF + "2. Ensure Private Idaho can communicate with your Web browser.  The default is Netscape Navigator.  If you're using another browser, choose the Options command in the Web menu." + gCRLF + gCRLF + "3. - You should be connected to the Internet with the browser running and not minimized." + gCRLF + gCRLF + "4. - From the Web menu, choose the 'Anonymous jump to URL' command." + gCRLF + gCRLF + "Your browser will anonymously access the Web page." + gCRLF + gCRLF + "Hint: Enter frequently accessed Web pages in Private Idaho's address book."
    Form32.Show
End Sub






Private Sub TabStrip1_Click()

End Sub

Private Sub TransferApp1_Click()

    Dim SectionName As String
    Dim AppName As String
    Dim AppScript As String
    Dim WindApp As String
    
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App1Wind")
    AppScript = ReadProfile(SectionName, "App1Script")
    TransferInfo WindApp$, AppScript
    
End Sub

Private Sub TransferApp2_Click()


    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    Dim WindApp As String
    
    SectionName = "Options"
    WindApp = ReadProfile(SectionName, "App2Wind")
    AppScript = ReadProfile(SectionName, "App2Script")
    TransferInfo WindApp, AppScript
    
End Sub

Private Sub TransferApp3_Click()


    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    Dim WindApp As String
    
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App3Wind")
    AppScript = ReadProfile(SectionName, "App3Script")
    TransferInfo WindApp$, AppScript
    
End Sub

Private Sub TransferApp4_Click()


    Dim WindApp As String
    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App4Wind")
    AppScript = ReadProfile(SectionName, "App4Script")
    TransferInfo WindApp$, AppScript
End Sub

Private Sub TransferAS_Click()
AppendInfo
End Sub
   

Private Sub TransferAT_Click()
    If Not gRemailerType = REMAILER_NONE Then
        AppendInfo
        TransferInfo gEmailer, gtranScript
    End If
End Sub

Private Sub TransferEncrypt_Click()
    EncryptMessage txtTo.Text, "-eatw"
End Sub

Private Sub TransferEu_Click()
    TransferInfo gEmailer, gtranScript
End Sub

Private Sub TransferNym_Click()
    frmCreateNymStep1.Show
End Sub

Private Sub TransferOptions_Click()
    SetEmailers
    Form5.Show 1
End Sub

Private Sub TransferPrepare_Click()
If txtTo.Text = "" Then
    ShowStatus ("You need to specify a recipient or have something to send.")
    Beep
    DoEvents
    Exit Sub
End If
gNymState = gNYMPREPARE
frmMultiNyms.Show vbModal
gNymState = gintNYM_IDLE
End Sub

Private Sub TransferReply_Click()
Dim foo As String
Dim fie As String
Dim tmpChar As String
Dim i As Integer
    MousePointer = vbHourglass
    If MessageArea.Text <> "" Then
        fie = ">"
        foo = InsertCRLFs()
        For i = 1 To Len(foo)
            tmpChar = Mid(foo, i, 1)
            fie = fie + tmpChar
            If tmpChar = Chr(10) Then
                fie = fie + ">"
            End If
        Next
        MessageArea.Text = fie
        txtTo.Text = gMessageRecord.From
    End If
    MousePointer = vbDefault
End Sub

Private Sub TransferSend_Click()
   SendPIMessage
End Sub



Private Sub USENETheader_Click()
'    USENETheader.Checked = Not USENETheader.Checked
End Sub

Private Sub UseNetSoda_Click()

    'USENETNone.Checked = False
    'USENETGate.Checked = False
    'USENETFi.Checked = False
    'UseNetSoda.Checked = True
    gNewsgroupType = Soda
End Sub




Private Sub EditSetFont_Click()

    '---------------------------------------------
    'display the common font dialog
    '---------------------------------------------
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    
    '---------------------------------------------
    'set the font for the message area
    '---------------------------------------------
    MessageArea.SelFontName = CommonDialog1.FontName
    MessageArea.SelBold = CommonDialog1.FontBold
    MessageArea.SelItalic = CommonDialog1.FontItalic
    MessageArea.SelStrikeThru = CommonDialog1.FontStrikethru
    MessageArea.SelFontSize = CommonDialog1.FontSize
    
    '---------------------------------------------
    'set the font for the to: box
    '---------------------------------------------
    txtTo.Font = CommonDialog1.FontName
    txtTo.FontBold = CommonDialog1.FontBold
    txtTo.FontItalic = CommonDialog1.FontItalic
    txtTo.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the subject: box
    '---------------------------------------------
    txtSubject.Font = CommonDialog1.FontName
    txtSubject.FontBold = CommonDialog1.FontBold
    txtSubject.FontItalic = CommonDialog1.FontItalic
    txtSubject.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the cc: box
    '---------------------------------------------
    txtCC.Font = CommonDialog1.FontName
    txtCC.FontBold = CommonDialog1.FontBold
    txtCC.FontItalic = CommonDialog1.FontItalic
    txtCC.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the bcc: box
    '---------------------------------------------
    'Text4.Font = CommonDialog1.FontName
   ' Text4.FontBold = CommonDialog1.FontBold
    'Text4.FontItalic = CommonDialog1.FontItalic
   ' Text4.FontStrikethru = CommonDialog1.FontStrikethru

End Sub

Public Property Get BusyCancel() As Boolean
    BusyCancel = m_BusyCancel
End Property

Public Property Let BusyCancel(ByVal bBusyCancel As Boolean)
    m_BusyCancel = bBusyCancel
End Property

Public Sub InitialiseDisplay()

If iFileExists(App.Path + "\remailer.txt") Then
    frmRemailerList.InitializeRemailers (App.Path + "\remailer.txt")
Else
    If iFileExists(App.Path + "\remailer.htm") Then frmRemailerList.InitializeRemailers (App.Path + "\remailer.htm")
End If
End Sub



Private Sub ImportFile()
Dim FileNum As Integer
    Dim TextLine As String
    Dim NumBytes As Long
    Dim msg As String
    Dim foo As Long
    Dim FileSize As Long
    Dim LineCount As Long
    
    On Error GoTo ImportError
    CommonDialog1.DialogTitle = "Open message text file"
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|Asc Files (*.asc)|*.asc|All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Action = 1
    FileNum = FreeFile
    Open CommonDialog1.FileName For Input As FileNum
        If Len(MessageArea.Text) > 0 Then
            foo = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Send Feedback")
            If foo = vbNo Then
                Exit Sub
            End If
        End If
        frmPI.MessageArea.SelStart = 0
        frmPI.MessageArea.Text = ""
        msg = ""
        MousePointer = vbHourglass
        frmPI.ShowStatus ("Loading file " & CommonDialog1.FileName & "...")
        DoEvents
        FileSize = FileLen(CommonDialog1.FileName)
        frmBusy.Style = 1
        frmBusy.AllowCancel = True
       ' frmBusy.BarCaption =
        frmBusy.CallingForm = frmPI
        frmBusy.Message = "Loading file " & CommonDialog1.FileName & ".  Please wait."
        frmBusy.Show
        Me.BusyCancel = False
        
        Do While Not EOF(FileNum)
            Line Input #FileNum, TextLine
            LineCount = LineCount + 1
            msg = msg & TextLine & vbCrLf
            NumBytes = NumBytes + Len(TextLine)
            frmBusy.BarPercent = 100 * NumBytes / FileSize
            frmBusy.BarCaption = "Processing line " & LineCount
            If Me.BusyCancel Then Exit Do
            DoEvents
        Loop
    MessageArea.Text = msg
    MousePointer = vbDefault
    Unload frmBusy
    'End If
    Close FileNum
    ChDir App.Path
    Exit Sub

ImportError:
    Close FileNum
    MsgBox Err.Description & " in FileImport", vbApplicationModal, App.Title
    Err.Clear
End Sub

Private Sub SaveMessage()
Dim FileNum As Integer
Dim rs As Recordset
Dim MessageFileName As String
    
    'Write to datbase here....
    '
    On Error GoTo WriteMessError
    'First Find the Inbox folder
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "[Folder] =" & "'" & "Inbox" & "'"
    If rs.NoMatch Then
        Err.Raise 1002, , "Inbox folder missing from database."
    End If
    frmMain.Tree.FolderId = rs("folder id")
    Set rs = DB.OpenRecordset("Messages", dbOpenDynaset)
    rs.AddNew
    rs("Folder ID") = frmMain.Tree.FolderId
    rs("To") = txtTo.Text
    rs("From") = gEmailAddress
    rs("Date") = CDate(Now())
    MessageFileName = DatePart("d", Now())
    MessageFileName = MessageFileName & DatePart("m", Now())
    MessageFileName = MessageFileName & Mid$(DatePart("yyyy", Now()), 3, 2)
    MessageFileName = MessageFileName & rs("Message ID") & ".msg"
    rs("Subject") = txtSubject.Text
    rs("CC") = txtCC.Text
    FileNum = FreeFile
    Open App.Path & "\mailbox\" & MessageFileName For Binary As FileNum
    Put #FileNum, , MessageArea.Text
    Close FileNum
    rs("Attachment") = False
    rs("Message File Name") = MessageFileName
    
    'This is crude but last thing is to strip out PGP message
    rs("Message Read") = True
    rs("Incoming Message") = False
    rs("Message Deleted") = False
    'If message has been decrypted then flag danger
   ' If MessageType.Decrypted Then
       ' rs("Message Status") = MessageType.DecryptedText
        'MessageType.Decrypted = False
   ' Else
        Select Case spgpAnalyseMessage(MessageArea.Text)
            Case "Encrypted"
                rs("Message Status") = MessageType.Encrypted
            Case "Signed"
                rs("Message Status") = MessageType.Signed
            Case "Detached Signature"
                rs("Message Status") = MessageType.DetachedSignature
            Case "Key"
                rs("Message Status") = MessageType.Key
            Case Else
                rs("Message Status") = MessageType.Unknown
        End Select
    'End If
    rs.Update
    rs.Close
    Exit Sub

WriteMessError:
    Close FileNum
    ShowStatus ("Following error occurred: " & Err.Description & " (Save Message)")
    Err.Clear
End Sub

Private Sub mNoRemailers_Click()
DontUseRemailer
End Sub

Private Sub EditPerform(EditFunction As Integer)
If TypeOf Me.ActiveControl Is TextBox Then
    Call SendMessage(Me.ActiveControl.hwnd, EditFunction, 0, 0&)
ElseIf TypeOf Me.ActiveControl Is RichTextBox Then
    If m_ControlKey = False Then
        Call SendMessage(Me.ActiveControl.hwnd, EditFunction, 0, 0&)
    End If
Else
    Beep
End If
End Sub

Private Sub UseCypherPunk()
   gRemailerType = REMAILER_CYPHERPUNK
   gEncryptToRemailer = True
    SSRibbon1(7).Enabled = True
    SSRibbon1(6).Enabled = True
    SSRibbon1(0).Picture = SSRibbon1(21).Picture
    frmRemailerList.InitializeRemailers (App.Path + "\remailer.htm")
    frmRemailerList.lblstatus = "Filling Remailer List..."
    frmRemailerList.SortRemailers
    frmRemailerList.FillRemailerList
   Exit Sub
   
    
End Sub

Private Sub UseMixmaster()
Dim SectionName As String
    gEncryptToRemailer = True
    gRemailerType = REMAILER_MIX
    SSRibbon1(7).Enabled = False
    frmRemailerList.InitializeRemailers (App.Path + "\mixmaster.htm")
    frmRemailerList.lblstatus = "Filling Remailer List..."
    frmRemailerList.SortRemailers
    frmRemailerList.FillRemailerList
    SSRibbon1(7).Enabled = False
    SSRibbon1(6).Enabled = False
    SSRibbon1(0).Picture = SSRibbon1(21).Picture
End Sub

Private Sub DontUseRemailer()
gRemailerType = REMAILER_NONE
gEncryptToRemailer = False
SSRibbon1(7).Enabled = True
End Sub

Private Sub AddAttachment()
Dim lListItem As ListItem
Dim j As Integer

   On Error GoTo AttachError

        '---------------------------------------------
        'File Open/Save Dialog Box Flags
        
        'Do this for PGP only - PGP2.6.3 can't handle long file names
        'CommonDialog1.FileTitle = cdlOFNNoLongNames
    
        CommonDialog1.DialogTitle = "Open file to attach."
        CommonDialog1.Flags = &H2& Or &H4& Or &H40000 'cdlOFNNoLongNames
        CommonDialog1.Filter = "All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path
        CommonDialog1.Action = 1
        
        j = FileIconsImageList1.GetFileIconNum(StripFileName(CommonDialog1.FileName))
        Set lListItem = lvwAttachments.ListItems.Add(, , CommonDialog1.FileName, j, j)
        
        ChDrive Mid$(App.Path, 1, 3)
        ChDir App.Path
    Exit Sub

AttachError:
    ShowStatus ("Error processing attachment - The following error occurred: " & Err.Description)
    Beep
    Err.Clear
End Sub

Private Sub EncryptMessageArea()
Dim PGPCmdString As String
Dim TheFileName As String
Dim tmpstr As String
Dim FileNum As Long
    
    On Error GoTo PGPEncryptError
    vb2spgpContext.Initialise
    '---------------------------------------------
    'verify there is something to encrypt
    '---------------------------------------------
   ' If (Len(MessageArea.Text) = 0) And Not PGPFile.Checked Then
       ' Beep
       ' MessageArea.Text = "Sorry...there's nothing to Encrypt!"
        'MessageArea.SelStart = 0
        'MessageArea.SelLength = Len(MessageArea.Text)
       ' Exit Sub
   ' End If
    'If (Len(txtTo.Text) = 0) Then
       ' Beep
       ' txtTo.Text = "You need to select a user id or recepient!"
       ' Exit Sub
   ' End If
       
    gCancelButton = False
    PGPCmdString = ""
    gPGPKeyID = ""
    
    '---------------------------------------------
    'handle case of saving a message to file, then encrypting
    '---------------------------------------------
    If PGPFile.Checked Then
        If Not Len(MessageArea.Text) = 0 Then
         tmpstr = MsgBox("Would you like to encrypt the message area and save it in a file?", vbYesNo, "File Encrypt")
         If tmpstr = vbYes Then
            CommonDialog1.DialogTitle = "Specify file for saving message"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            TheFileName = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            FileNum = FreeFile
            Open CommonDialog1.FileName For Output As FileNum
                Print #FileNum, MessageArea.Text
            Close #FileNum
         End If
        End If
        '---------------------------------------------
        'no, it is a file encrypt
        '---------------------------------------------
        If Len(TheFileName) = 0 Then
            '---------------------------------------------
            'user has not already specified a file name from above
            '---------------------------------------------
            CommonDialog1.DialogTitle = "Open file to encrypt"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            TheFileName = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
    
        End If
        If PGPEyes.Checked = True Then
            PGPCmdString = "m"
        End If
        vb2spgpContext.TextMode = 0
        vb2spgpContext.Armor = 1
        vb2spgpContext.FileIn = TheFileName
        vb2spgpContext.FileOut = StripExt(TheFileName) & ".asc"
        If PGPConvent.Checked Then
            vb2spgpContext.ConventionalEncrypt = 1
            EncryptFile "", "-cat" + PGPCmdString, TheFileName
        Else
            vb2spgpContext.ConventionalEncrypt = 0
            EncryptFile txtTo.Text, "-eat" + PGPCmdString, TheFileName
        End If
    Else
    '---------------------------------------------
    'is this a eyes, conventional, other encrypt?
    '---------------------------------------------
        If PGPEyes.Checked = True Then
            PGPCmdString = "m"
        End If
        vb2spgpContext.TextMode = 0
        vb2spgpContext.Armor = 1
        If PGPConvent.Checked = True Then
            vb2spgpContext.ConventionalEncrypt = 1
            vb2spgpContext.KeyEncrypt = 0
            EncryptMessage "", "-catw" + PGPCmdString
        Else
            vb2spgpContext.ConventionalEncrypt = 0
            vb2spgpContext.KeyEncrypt = 1
            EncryptMessage txtTo.Text, "-eatw" + PGPCmdString
        End If
    End If
    
    Exit Sub
PGPEncryptError:
     MsgBox Err.Description & " in PGP Encryption", vbCritical + vbApplicationModal, App.Title
     Err.Clear
End Sub

Public Sub CreateReplyBlock()
    AppendNymInfo
    MessageArea.SelStart = 0
    MessageArea.SelText = "Reply-Block:" + vbLf
    MessageArea.SelText = "::" & vbLf
    MessageArea.SelText = "Anon-To: " & frmPI.txtTo.Text & vbLf
    MessageArea.SelText = "Latent-Time: " & gLatentTime & vbLf
   ' MessageArea.SelText = "Nym-Commands: " & IIf(gAcksend, "+acksend", "-acksend") & vbLf
    If Not gNymPassPhrase(1) = "" Then
        MessageArea.SelText = "Encrypt-Key: " & gNymPassPhrase(1) & vbLf & vbLf
    Else
        MessageArea.SelText = vbLf
    End If

End Sub



Public Sub DisablePGPMenuItems()
Exit Sub
mPGP.Visible = False
'mRemailers.Visible = False
mNewsgroups.Visible = False
mFingerOps.Visible = False
mMessage.Visible = False
mNym.Visible = False
End Sub

Public Sub EnablePGPMenuItems()
Exit Sub
mPGP.Visible = True
'mRemailers.Visible = True
mNewsgroups.Visible = True
mFingerOps.Visible = True
mMessage.Visible = True
mNym.Visible = True
End Sub

Public Function spgpAnalyseMessage(inBuffer As String) As String
Dim Buffer As String
  Dim i As Long
 ' If inBuffer = "" Then
   ' Buffer = String(Len(MessageArea.Text & Chr(0)), Chr(0))
   ' Buffer = MessageArea.Text & Chr(0)
  'Else
    Buffer = String(Len(inBuffer & Chr(0)), Chr(0))
    Buffer = inBuffer & Chr(0)
 'End If
  i = spgpAnalyze(Buffer)

  Select Case i
  Case PGPAnalyze_Encrypted ' Encrypted message
        spgpAnalyseMessage = "Encrypted"
  Case PGPAnalyze_Signed ' Signed message
        spgpAnalyseMessage = "Signed"
  Case PGPAnalyze_DetachedSignature ' Detached signature
        spgpAnalyseMessage = "Detached Signature"
  Case PGPAnalyze_Key ' Key data
        spgpAnalyseMessage = "Key"
  Case Else
        spgpAnalyseMessage = "Unknown" ' Global Const PGPAnalyze_Unknown = 4              ' Non-pgp message
  End Select
  
End Function

Public Sub GetRecipient()
Dim foostr As String
    Dim x As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim i As Integer
        
    '---------------------------------------------
    'clear any current name from the "to:" box
    '---------------------------------------------
    txtTo.Text = ""
    gKeyID = ""
    '---------------------------------------------
    'present the list of public keys
    '---------------------------------------------
    frmSelectUserID.Label2 = ""
    frmSelectUserID.Label1 = "Select a recipient from either your public key ring or your personal address book."
    frmSelectUserID.Show vbModal
    '---------------------------------------------
    'process the key, if any
    '---------------------------------------------
    If gKeyID <> "null" Then
        '---------------------------------------------
        'if user has selected a key, strip off quotes
        '---------------------------------------------
        pos2 = Len(gKeyID)
        foostr = ""
        For i = 1 To pos2
            x = Mid(gKeyID, i, 1)
            If (x <> Chr$(34)) Then
                foostr = foostr + x
            End If
        Next
        txtTo.Text = LTrim(foostr)
    End If
End Sub

Public Sub ShowStatus(Status As String)
StatusBar.Style = sbrNormal
If TextWidth(Status & "  ") > StatusBar.Panels(1).Width Then StatusBar.Panels(1).Width = TextWidth(Status & " ")
StatusBar.Panels.Item(1) = Status
End Sub

Private Function CheckConnection()
    Dim ReturnValue As Boolean
    Dim WaitAWhile As Variant
    Dim Req1 As Long
    Dim Response As String

    On Error GoTo BadConnection
    If IPInfo1.WinsockLoaded Then
        CheckConnection = True
    Else
        IPInfo1.WinsockLoaded = True
        CheckConnection = True
    End If
    DoEvents
    Exit Function
    
BadConnection:
    CheckConnection = False
    Beep
    MsgBox "Can't create connection", vbApplicationModal + vbCritical
    Err.Clear
End Function

Public Sub ShowAttachment(Attachment As String)
StatusBar.Style = sbrNormal
If TextWidth(Attachment & "  ") > StatusBar.Panels(2).Width _
    Then StatusBar.Panels(2).Width = TextWidth(Attachment & " ")
StatusBar.Panels.Item(2) = Attachment
End Sub

Public Sub ShowRemailer(Remailer As String)
StatusBar.Style = sbrNormal
If TextWidth(Remailer & "  ") > StatusBar.Panels(3).Width _
    Then StatusBar.Panels(3).Width = TextWidth(Remailer & " ")
StatusBar.Panels.Item(3) = Remailer
End Sub

Private Sub SetMessageReadState()
'Turn off send button
SSRibbon1(0).Enabled = False
TransferSend.Enabled = False
cmbRemailerSelect.Enabled = False
cmbRemailerSelect.ListIndex = 0
btnTo(0).Enabled = False
btnTo(1).Enabled = False

txtTo.Appearance = 0
txtTo.BackColor = vbMenuBar
txtTo.BorderStyle = 0
txtTo.Enabled = False
txtCC.Appearance = 0
txtCC.BackColor = vbMenuBar
txtCC.BorderStyle = 0
txtCC.Enabled = False
txtSubject.Appearance = 0
txtSubject.BackColor = vbMenuBar
txtSubject.BorderStyle = 0
txtSubject.Enabled = False


End Sub

Private Sub SetAttachmentEncryptionOptions()
    mDontEncryptAttachment.Checked = True
    mEncryptAttachmentWithKey.Checked = False
    mConventionallyEncryptAttachment.Checked = False
End Sub

