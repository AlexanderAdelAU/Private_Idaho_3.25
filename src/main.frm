VERSION 5.00
Object = "{33337143-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "netcod40.ocx"
Object = "{33337153-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ipinfo40.ocx"
Object = "{33337233-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "smtp40.ocx"
Object = "{33337243-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "pop40.ocx"
Object = "{33337313-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "imap40.ocx"
Object = "{33337293-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "mx40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "PI Mail and File "
   ClientHeight    =   7395
   ClientLeft      =   2835
   ClientTop       =   4215
   ClientWidth     =   12180
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7395
   ScaleWidth      =   12180
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   10680
      TabIndex        =   12
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   7095
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15822
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   10680
      Top             =   6600
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Align           =   1  'Align Top
      Height          =   4710
      Left            =   0
      TabIndex        =   2
      Top             =   435
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   8308
      _Version        =   131074
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      BorderStyle     =   1
      BackColor       =   -2147483644
      PaneTree        =   "main.frx":030A
      Begin SSActiveTreeView.SSTree SSTree1 
         CausesValidation=   0   'False
         Height          =   4680
         Index           =   0
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   8255
         _Version        =   65538
         LineStyle       =   1
         TreeTips        =   1
         Indentation     =   315
         AllowDelete     =   -1  'True
         HideSelection   =   0   'False
         HasFont         =   0   'False
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "ImageList1"
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "main.frx":035C
         Height          =   4680
         Left            =   3030
         TabIndex        =   0
         Top             =   15
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8255
         _Version        =   393216
         Rows            =   7
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   -2147483644
         BackColorBkg    =   -2147483643
         GridColorFixed  =   16777215
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         AllowUserResizing=   1
      End
      Begin MXLibCtl.MX MX1 
         Left            =   10440
         Top             =   7800
         DNSServer       =   ""
         WinsockLoaded   =   -1  'True
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   767
      _Version        =   131074
      PictureBackgroundStyle=   2
      PictureBackground=   "main.frx":0375
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   7
         Left            =   3570
         TabIndex        =   13
         ToolTipText     =   "File Safe"
         Top             =   30
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":10C9F
         Alignment       =   4
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   6
         Left            =   840
         TabIndex        =   9
         ToolTipText     =   "Send and Receive Messages"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":1177D
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   5
         Left            =   2250
         TabIndex        =   8
         ToolTipText     =   "Forward Message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":1188F
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   4
         Left            =   450
         TabIndex        =   7
         ToolTipText     =   "New sub-folder"
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":119A1
         ButtonStyle     =   3
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   7
         Left            =   7290
         Picture         =   "main.frx":11DE3
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   3
         Left            =   2880
         TabIndex        =   6
         ToolTipText     =   "Print Message"
         Top             =   60
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":1234B
         Alignment       =   4
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   2
         Left            =   1860
         TabIndex        =   5
         ToolTipText     =   "Reply To Sender"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":1288D
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   4
         ToolTipText     =   "New message"
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":129CB
         ButtonStyle     =   3
      End
      Begin VB.Image ImgList 
         Height          =   480
         Index           =   5
         Left            =   9450
         Picture         =   "main.frx":12E91
         Top             =   -30
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgList 
         Height          =   480
         Index           =   4
         Left            =   8850
         Picture         =   "main.frx":1319B
         Top             =   -30
         Visible         =   0   'False
         Width           =   480
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   1
         Left            =   1470
         TabIndex        =   3
         ToolTipText     =   "Delete selected message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main.frx":134A5
         ButtonStyle     =   3
      End
      Begin VB.Image ImgList 
         Height          =   195
         Index           =   3
         Left            =   8490
         Picture         =   "main.frx":135E7
         Top             =   90
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   0
         Left            =   7590
         Picture         =   "main.frx":136D1
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   1
         Left            =   7950
         Picture         =   "main.frx":13BDF
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   2
         Left            =   8370
         Picture         =   "main.frx":14133
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin NETCODELibCtl.NetCode NetCode1 
      Left            =   6120
      Top             =   6600
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
   Begin SMTPLibCtl.SMTP SMTP1 
      Left            =   7080
      Top             =   6600
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
   Begin IMAPLibCtl.IMAP IMAP1 
      Left            =   8280
      Top             =   6600
      LocalFile       =   ""
      Mailbox         =   "Inbox"
      MailServer      =   ""
      MessageSet      =   ""
      Password        =   ""
      SearchCriteria  =   ""
      User            =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin POPLibCtl.POP POP1 
      Left            =   8880
      Top             =   6600
      LocalFile       =   ""
      MailServer      =   ""
      MaxLineLength   =   2048
      MaxLines        =   0
      MessageNumber   =   0
      Password        =   ""
      User            =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   23
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":14263
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":145B5
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":14907
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":14C59
            Key             =   "happy"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":14F73
            Key             =   "apathy"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":1528D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":155A7
            Key             =   "bulbon"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":158C1
            Key             =   "bulboff"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":15BDB
            Key             =   "question"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":15EF5
            Key             =   "openlock"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":1620F
            Key             =   "closedlock"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":16529
            Key             =   "exclamation"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":16843
            Key             =   "BrokenEnvelope"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":16B5D
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":16CAF
            Key             =   "Mask"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":16E01
            Key             =   "EncryptedEnvelope"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":17353
            Key             =   "Closed Envelope"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":178A5
            Key             =   "Open Envelope"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":17DF7
            Key             =   "FolderGroup"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":17FD1
            Key             =   "DragIcon"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":182EB
            Key             =   "DropIcon"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":18605
            Key             =   "New Folder"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "main.frx":18A47
            Key             =   "Open Eye"
         EndProperty
      EndProperty
   End
   Begin IPINFOLibCtl.IPInfo IPInfo1 
      Left            =   7680
      Top             =   6600
      PendingRequests =   1
      ServiceName     =   ""
      ServicePort     =   0
      ServiceProtocol =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mFile_Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mFile_Delete 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu Filebreak 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mFile_NewFolderGroup 
         Caption         =   "New Folder Group"
      End
      Begin VB.Menu mFile_NewFolder 
         Caption         =   "New Sub Folder"
      End
      Begin VB.Menu Filebreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mViewHeaders 
         Caption         =   "View Headers"
      End
      Begin VB.Menu Filebreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mFile_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mCompose 
      Caption         =   "Compose"
      Begin VB.Menu mCompose_NewMessage 
         Caption         =   "New Message"
      End
      Begin VB.Menu mCompose_ReplyTo 
         Caption         =   "Reply To"
      End
   End
   Begin VB.Menu Mview 
      Caption         =   "View"
      Begin VB.Menu MviewFileList 
         Caption         =   "View Files List"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "Tools"
      Begin VB.Menu mTools_EmailScan 
         Caption         =   "Send and Receive Messages"
      End
      Begin VB.Menu mViewIMAPMailboxes 
         Caption         =   "View IMAP Mailboxes"
      End
      Begin VB.Menu mbreak5 
         Caption         =   "-"
      End
      Begin VB.Menu mToolsAddressBook 
         Caption         =   "Address Book"
      End
      Begin VB.Menu mbreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mEmail_Options_Header 
         Caption         =   "Email Options"
         Begin VB.Menu mEmailOptions 
            Caption         =   "Connection Options"
            Index           =   0
         End
         Begin VB.Menu mEmailOptions 
            Caption         =   "Email Retrieve Options"
            Index           =   1
         End
         Begin VB.Menu mEmailOptions 
            Caption         =   "Mail Groups"
            Index           =   2
         End
         Begin VB.Menu mPreviewMessages 
            Caption         =   "Preview Messages"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mbreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mCompressDatabase 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu mFileSafe 
         Caption         =   "File Safe"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuTreeListContext 
      Caption         =   "TreeListContext Menu"
      Visible         =   0   'False
      Begin VB.Menu popupNewGroupFolder 
         Caption         =   "New Group Folder"
      End
      Begin VB.Menu popupNewSubFolder 
         Caption         =   "New Sub Folder"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Folders(100) As String
Private Const CLOSED_ENVELOPE As Integer = 0
Private Const OPEN_ENVELOPE As Integer = 1
Private Const Attachment As Integer = 2
'Private Const DANGER As Integer = 6
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
'Private Const TIMER_DRAG_DROP As Integer = 0
'Private Const TIMER_MESSAGE_SCAN As Integer = 1

Private Const MAX_OPEN_MESSAGES = 10

Private Type MouseCoordinates
    x As Integer
    y As Integer
End Type
Private MouseDownPosition As MouseCoordinates
Private gDragCommenced As Boolean

Private sMailServer As String
'Private frmMessages(MAX_OPEN_MESSAGES) As New frmPI
'Public InstanceNumber As Integer
'Private MousePressed As Boolean
Private m_GroupNodeMarkedForDelete As Boolean



Public Sub BuildTree(Index)
    Dim i As Integer
    Dim FolderCount As Integer
    Dim Key As String
    Dim start As Long
    Dim n As SSNode
    Dim rsNode As Recordset
    Dim rsItems As Recordset
    Dim ItemsinFolder As Long
    Dim rsFolder As Recordset
    Dim qd As QueryDef
    Dim qdItems As QueryDef
    Dim sNodeName As String
    Dim sFolderName As String
    'Root folders
    'Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    On Error GoTo BadTreeBuild
    SSTree1(Index).Nodes.Clear
    DoEvents
    If Index = 0 Then
        Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
        Set n = SSTree1(Index).Nodes.Add(, , , "Private Idaho E-Mail", "EncryptedEnvelope", "EncryptedEnvelope")
    Else
        Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
        Set n = SSTree1(Index).Nodes.Add(, , , "Private Idaho File Safe", "EncryptedEnvelope", "EncryptedEnvelope")
    
    End If
    n.Font.Bold = True
    Set n = Nothing
  ' If Not rsNode.EOF Then rsNode.MoveFirst
    On Error Resume Next
    Do While Not rsNode.EOF
        sNodeName = rsNode("Node Name")
        Set n = SSTree1(Index).Nodes.Add(, , sNodeName, sNodeName, "FolderGroup", "FolderGroup")
       ' n.Expanded = True
        n.Font.Bold = True
        n.Expanded = True
       ' Set n = Nothing
        If Index = 0 Then
            Set qd = DB.QueryDefs("qdSubFolders")
        Else
            Set qd = DB.QueryDefs("qdFileSubFolders")
        End If
        qd.Parameters![NodeID] = rsNode("Node ID")
        Set rsFolder = qd.OpenRecordset()
        Do While Not rsFolder.EOF
            sFolderName = rsFolder("Folder")
         '   If Index = 0 Then
         '       Set qdItems = DB.QueryDefs("qdNumberofMessagesinFolder")
          '      qdItems.Parameters![FolderId] = rsFolder("[Folder ID]")
         '       Set rsItems = qdItems.OpenRecordset()
          '      If Not rsItems.RecordCount = 0 Then sFolderName = sFolderName & " (" & rsItems("[Number of Messages]") & ")"
          '  Else
         '       Set qdItems = DB.QueryDefs("qdNumberofFilesinFolder")
            '    qdItems.Parameters![FolderId] = rsFolder("[Folder ID]")
           '     Set rsItems = qdItems.OpenRecordset()
          '      If Not rsItems.RecordCount = 0 Then sFolderName = sFolderName & " (" & rsItems("[Number of Files]") & ")"
           ' End If
            Set n = SSTree1(Index).Nodes.Add(sNodeName, ssatChild, , sFolderName, "closed", "open")
           ' n.Expanded = True
           ' n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
            
           ' Set n = Nothing
            rsFolder.MoveNext
        Loop
         'MsgBox "Debug 24", vbApplicationModal
        rsNode.MoveNext
    Loop
    rsFolder.Close
    rsNode.Close

SSTree1(Index).Visible = True
 AddFolderParameters (Index)
n.Expanded = True
Exit Sub
BadTreeBuild:
    MsgBox "Private Idaho Error:  " & Err.Description, vbCritical + vbApplicationModal, "Build Tree"
    Err.Clear
    Exit Sub

End Sub




Private Sub Form_Activate()
Dim Win As New CWindow
Win.OnTop(Me) = False
FillGrid
Set Win = Nothing
'lblstatus.Visible = False



End Sub
Private Sub InitialiseProgressBar()

ProgressBar1.Top = StatusBar1.Top + 10
ProgressBar1.Left = StatusBar1.Panels.Item(3).Left + 10
ProgressBar1.Height = 0.9 * StatusBar1.Height
ProgressBar1.Width = StatusBar1.Panels.Item(3).Width - 10
ProgressBar1.Visible = True
ProgressBar1.Min = 0
ProgressBar1.Max = 100
ProgressBar1.Value = 0
'StatusBar1.Visible = True
End Sub
Private Sub ShowStatus(Panel As Integer, Status As String)
StatusBar1.Style = sbrNormal
If TextWidth(Status & "  ") > StatusBar1.Panels(Panel).Width Then StatusBar1.Panels(Panel).Width = TextWidth(Status & " ")
StatusBar1.Panels.Item(Panel) = Status
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
MousePointer = vbDefault
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
MousePointer = vbNoDrop
DoEvents
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
Dim App As New CApplication

Dim i As Integer
'Dim Headings As String
Dim RowIndex As Integer
Dim ItemString As String
Dim Sqlq As String

'==============================
'Check for previous instance and initialise number of PI form instances
'================================
If ProgramIsAlreadyRunning Then Stop
gFormInstance = 0
Win.Center Me, Null
On Error Resume Next
'Make sure directories exisit

If Dir(App.Path & "mailbox\") = "" Then
    MkDir App.Path & "mailbox"
End If

If Dir(App.Path & "temp\") = "" Then
    MkDir App.Path & "temp"
Else
    Kill App.Path & "\temp\*.*"
End If
On Error GoTo BadLoad

Select Case SecurityCheck
    Case Is = "INVALID"
        gFullRelease = -1
    
    Case Is = "TRIAL"
        gFullRelease = 0
    
    Case Else
        gFullRelease = 1
End Select

Me.Caption = "Private Idaho E-Mail (" & App.Version & ") for Win9x/Win2k/NT"
Call InitialiseGrid
SSTree1(0).ImageList = ImageList1
BuildTree (0)
DisplayInBox

RestoreMainSettings
'
' Align all the controls
'
InitialiseProgressBar
Exit Sub
BadLoad:
    MsgBox Err.Description & " Can't load properly...", vbApplicationModal + vbCritical, App.Title
    Err.Clear
    Resume Next
End Sub

Private Sub Form_Resize()
 Dim lWidth As Long
 On Error Resume Next
 If WindowState <> 1 Then
    'frmMain.Refresh
    StatusBar1.Visible = False
    
    ProgressBar1.Visible = False
    'InitialiseProgressBar
    'If mPGP.Enabled Then
      ' SSSplitter1.Width = Width - SSSplitter1.Left - Width * 0.02
        SSSplitter1.Height = StatusBar1.Top - StatusBar1.Height - 100 'SSSplitter1.Top - Height * 0.12
   ' End If
   ' MSFlexGrid1.Visible = False
    lWidth = MSFlexGrid1.Width
    'MSFlexGrid1.ColWidth(0) = lWidth * 0.03
    MSFlexGrid1.ColWidth(0) = lWidth * 0.08
    MSFlexGrid1.ColWidth(1) = lWidth * 0.05
    MSFlexGrid1.ColWidth(2) = lWidth * 0.28
    MSFlexGrid1.ColWidth(3) = lWidth * 0.3
    MSFlexGrid1.ColWidth(4) = lWidth * 0.22
   
    'DoEvents
     InitialiseProgressBar
    ' DoEvents
    'MSFlexGrid1.Visible = True

    'DoEvents
    ProgressBar1.Visible = True
        StatusBar1.Visible = True
 End If
 
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form
On Error Resume Next
For Each Form In Forms
    If Form.name <> Me.name Then
        Unload Form
        Set Form = Nothing
    End If
Next Form
Set frmMain = Nothing
'Delete all temporary files
Kill App.Path & "\temp\*.*"
'End
End Sub



Private Sub IMAP1_Error(ErrorCode As Integer, Description As String)
MsgBox "An IMAP Server error has been detected.  Code = " & CStr(ErrorCode) & ", Description: " & Description, vbApplicationModal + vbCritical, " IMAP Error"
End Sub

Private Sub IMAP1_MailboxList(Mailbox As String, Separator As String, Flags As String)
'Dim ObjItem As ListItem
frmIMAPMailBoxList.List1.AddItem Mailbox & " , " & Separator & " , " & Flags
frmIMAPMailBoxList.Show
'ObjItem = frmIMAPMailBoxList.ListView1.ListItems.Add(, "Key", Mailbox & Separator & Flags)
End Sub

Private Sub IMAP1_Transfer(BytesTransferred As Long, Text As String)
Dim ptrText As Long
Dim sText As String
    frmStatus.lblBytesTransferred.Caption = Val(frmStatus.lblBytesTransferred.Caption) + BytesTransferred
    If gBuffer.hMem = 0 Then
        DoEvents
        gMessage = gMessage & Text & vbCrLf
    Else
        If Len(Text) > 8046 Then
            ShowStatus 1, "Buffer overflow error..."
            Beep
            Exit Sub
        End If
        sText = Text & vbCrLf
        Call agCopyData(ByVal sText, ByVal gBuffer.Address, ByVal CLng(Len(sText)))
        gBuffer.Address = gBuffer.Address + CLng(Len(sText))
    End If
    DoEvents
End Sub




Private Sub mAbout_Click()
 frmAbout.Show vbModal
End Sub

Private Sub mCompose_NewMessage_Click()
Me.MousePointer = vbHourglass
gComposeMode = True
CreateNewPIInstance
Me.MousePointer = vbDefault
End Sub

Private Sub mCompose_ReplyTo_Click()
Dim iMessageInstance As Integer
Me.MousePointer = vbHourglass
iMessageInstance = DisplayMessage
ReplyToSender (iMessageInstance)
Me.MousePointer = vbDefault
End Sub



Private Sub mCompressDatabase_Click()
Dim DBName As String
Dim DBBack As String
Dim dbNew As String

On Error GoTo BackupProblem

MousePointer = vbHourglass
ShowStatus 1, "Making backup first...."
'lblstatus.Visible = True
DoEvents
DB.Close
Set DB = Nothing
DBName = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "PI32PostOffice.MDB"
DBBack = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "PI32PostOffice.BAK"
dbNew = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "PI32PostOffice_NEW.MDB"

FileCopy DBName, DBBack

'It may exist due to a bad procedure last time so kill it
On Error Resume Next
Kill dbNew

On Error GoTo WarnUser
ShowStatus 1, "Now compacting your database...."
DoEvents
Dim spass As String

'spass = "redemption"
'spass = spass & "loosesteps"
'spass = spass & "onetwoone"

'spass = Mid(spass, 1, 3) & Mid(spass, 16, 2) & Mid(spass, 21, 3) & "121"
spass = "redim"
spass = spass & "string"
spass = spass & "integer one1"

spass = Mid(spass, 1, 3) & Mid(spass, Len(spass), 1) & "21st" & Mid(spass, Len(spass) - 3, 3) '& Mid(spass, 6, 3)


DBEngine.CompactDatabase DBName, dbNew, , dbEncrypt, ";pwd=" & spass
Kill DBName
Name dbNew As DBName
ShowStatus 1, "Database successfully compacted..."
DoEvents
MousePointer = vbDefault
OpenDatabase
Exit Sub

WarnUser:
    DoEvents
    MousePointer = vbDefault
    MsgBox "An serious error has been detected.  " & Err.Description, vbCritical, "Compress Database"
    'Beep
    Err.Clear
    OpenDatabase
    'lblstatus.Visible = False
    Exit Sub
BackupProblem:
    DoEvents
    MousePointer = vbDefault
    MsgBox "Can't create back - cannot proceed.  " & Err.Description, vbCritical, "Compress Database"
   ' Beep
    Err.Clear
    OpenDatabase
    'lblstatus.Visible = False
    
End Sub



Private Sub mEmailOptions_Click(Index As Integer)

Dim SectionName As String

'Set up all the defaults
frmMailServerOptions.ScanInterval = giEmailScanInterval
frmMailServerOptions.EmailScanOption = giEmailScanOption
'Okay show the form
frmMailServerOptions.pOptionsTab = Index
frmMailServerOptions.Show vbModal

'Now get the data
giEmailScanInterval = frmMailServerOptions.ScanInterval
giTimerCounter = giEmailScanInterval

If giEmailScanInterval > 0 Then Timer1.Enabled = True
giEmailScanOption = frmMailServerOptions.EmailScanOption

SectionName = "MailServerOptions"
WriteProfile SectionName, "MessageScanInverval", CStr(giEmailScanInterval)
WriteProfile SectionName, "EmailScanOption", CStr(giEmailScanOption)

Unload frmMailServerOptions
DoEvents

End Sub

Private Sub mFile_Delete_Click()
DeleteMessage
InitialiseGrid
FillGrid
End Sub

Private Sub mFile_Exit_Click()
'Unload PIForm(1)
Unload Me
End Sub



Private Sub mFile_NewFolder_Click()
CreateSubFolder
DisplayInBox
End Sub

Private Sub mFile_NewFolderGroup_Click()
Dim rs As Recordset
Dim Index As Integer
Dim sFolderName As String


On Error Resume Next
'Index = giTreeIndex
frmFolderName.Caption = "Create New Group Folder"
frmFolderName.lblNamePrompt = "Enter the name of the Group Folder"
frmFolderName.Show vbModal
sFolderName = frmFolderName.txtFolderName
Set frmFolderName = Nothing

   'If Index = 0 Then
        Set rs = DB.OpenRecordset("Nodes", dbOpenDynaset)
   ' Else
    '    Set rs = DB.OpenRecordset("File Nodes", dbOpenDynaset)
   ' End If
   
    rs.AddNew
    rs("Node Name") = sFolderName
    rs("Can Delete") = True
    rs.Update
    rs.Close
    Set rs = Nothing
    BuildTree (Index)
 
'End If
End Sub

Private Sub mFile_Open_Click()
Dim iCurrentMessage As Integer

Me.MousePointer = vbHourglass
If MSFlexGrid1.Row = 0 Then
    MsgBox "No messages selected", vbApplicationModal + vbCritical, "Message Open Error"
    Me.MousePointer = vbHourglass
    Exit Sub
End If
iCurrentMessage = DisplayMessage
 If iCurrentMessage Then
    PIForm(iCurrentMessage).SSRibbon1(0).Enabled = False ' don't allow them to send
    PIForm(iCurrentMessage).SSRibbon1(6).Enabled = False ' don't allow them to send
    PIForm(iCurrentMessage).SSRibbon1(9).Enabled = False ' don't allow them to send
    PIForm(iCurrentMessage).SSRibbon1(4).Enabled = False ' don't allow them to send
    PIForm(iCurrentMessage).SSRibbon1(7).Enabled = False ' don't allow them to send
    PIForm(iCurrentMessage).SSRibbon1(10).Enabled = True ' allow reply to sendere
    PIForm(iCurrentMessage).SSRibbon1(12).Enabled = True ' allow forware to sendere
End If
Me.MousePointer = vbDefault
End Sub




Private Sub mFileSafe_Click()
 frmFileSafe.Show
End Sub

Private Sub mPGPFileSafeOptions_Click()
frmPGPOptions.Show vbModal
End Sub

Private Sub mPreviewMessages_Click()
mPreviewMessages.Checked = Not mPreviewMessages.Checked
End Sub

Private Sub mRegistrationCheck_Click()
frmLicence.Show
End Sub

Private Sub MSFlexGrid1_Click()
'Grid.SelectedRow = MSFlexGrid1.Row
End Sub

Private Sub MSFlexGrid1_DblClick()
Dim n As SSNode
Dim lFolderIndex As Long
Dim iCurrentMessage As Integer
 '
'First Find the node name
'
On Error Resume Next
lFolderIndex = SSTree1(0).SelectedItem.Index
Set n = SSTree1(0).Nodes.Item(lFolderIndex)
Me.MousePointer = vbHourglass
DoEvents

iCurrentMessage = DisplayMessage

If Not StripItemCount(n.Text) = "Drafts" Then
    ConfigureDisplayControls (iCurrentMessage)
    PIForm(iCurrentMessage).DontUseRemailer
End If
'lFolderIndex = SSTree1(0).SelectedItem.Index
Me.MousePointer = vbDefault
'MSFlexGrid1.SetFocus
'Now deselect the tree
Set n = SSTree1(0).SelectedNodes

n.Expanded = True
'MSFlexGrid
'n.Selected = False
'ssNodeTmp.Selected = False
End Sub

Private Sub MSFlexGrid1_DragDrop(Source As Control, x As Single, y As Single)
MousePointer = vbDefault
End Sub



Private Sub MSFlexGrid1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
On Error Resume Next
    
   If Not Source.name = "MSFlexGrid1" Then
        Source.Value.Drag vbEndDrag
        MSFlexGrid1.Drag vbEndDrag
        MSFlexGrid1.MousePointer = flexNoDrop
   End If
End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'Dim
On Error Resume Next

Me.MousePointer = vbHourglass
DoEvents
If KeyCode = vbKeyDelete Then
    DeleteMessage
    'This selects a single row again to stop multiple deletes
    Grid.RowSelection = Grid.SelectedRow
End If


'Now highlight selected cell
   
   If Grid.SelectedRow <= MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Row = Grid.SelectedRow
   End If
 'Highlight the selected row
    For i = 0 To 4
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbHighlight
        MSFlexGrid1.CellForeColor = vbWhite
    Next
   ' MSFlexGrid1.SetFocus
    
Me.MousePointer = vbDefault
End Sub

Private Sub MSFlexGrid1_LostFocus()
DisableToolBarButtons 'moved to tr
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim LastRow As Integer
Dim FirstRow As Integer
Dim J As Integer
Dim PreviousRow As Integer
Dim PreviousSelection As Integer

On Error Resume Next

'Reset this for the drag and drop

gDragCommenced = False

'And this to allows a delay before showing the drag icon
'Mouse down marks the event - so we need to grad the co-ordinates here
MouseDownPosition.x = x
MouseDownPosition.y = y

'Exit Sub


MousePointer = vbDefault

'This is better than making the grid non visible as that will fire off a lost focus


'First clear previous selection
'DoEvents
'if we have clicked on the field, then sort
'The 50 stops false triggers
If y < MSFlexGrid1.Top + MSFlexGrid1.CellHeight Then 'CellTop - 50 Then
  Grid.SelectedColToSort = MSFlexGrid1.Col
  FillGrid
  Exit Sub
End If
'If we have selected in the previous selection don't do anything so we can drag
'If Shift = 0 it is a free select
If Shift = 0 Then
    If MSFlexGrid1.RowSel >= Grid.SelectedRow And MSFlexGrid1.RowSel <= Grid.SelectedRow + Grid.RowSelection Then
        Exit Sub
    End If
End If
MSFlexGrid1.Visible = False
    PreviousRow = Grid.SelectedRow
    PreviousSelection = Grid.RowSelection
    Grid.SelectedCol = MSFlexGrid1.Col
    Grid.RowSelection = MSFlexGrid1.RowSel
    Grid.SelectedRow = MSFlexGrid1.Row
    
If Not PreviousSelection = 0 Then
    If PreviousRow < PreviousSelection Then
        FirstRow = PreviousRow
        LastRow = PreviousSelection
    Else
        FirstRow = PreviousSelection
        LastRow = PreviousRow
    End If
    
    For J = FirstRow To LastRow
        For i = 0 To 4
            MSFlexGrid1.Row = J
            MSFlexGrid1.Col = i
            MSFlexGrid1.CellBackColor = vbWhite
            MSFlexGrid1.CellForeColor = vbBlack
        Next
    Next
End If

If Shift = 0 Then
    MSFlexGrid1.Row = PreviousRow
    For i = 0 To 4
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbWhite
        MSFlexGrid1.CellForeColor = vbBlack
    Next
    MSFlexGrid1.Row = Grid.SelectedRow
    For i = 0 To 4
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbHighlight
        MSFlexGrid1.CellForeColor = vbWhite
    Next
Else
    
    If PreviousRow < Grid.RowSelection Then
        FirstRow = PreviousRow
        LastRow = Grid.RowSelection
    Else
         FirstRow = Grid.RowSelection
        LastRow = PreviousRow
    End If
    
    For J = FirstRow To LastRow
    For i = 0 To 4
        MSFlexGrid1.Row = J
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbHighlight
        MSFlexGrid1.CellForeColor = vbWhite
    Next
Next

End If
'MSFlexGrid1.Enabled = True
'
MSFlexGrid1.Visible = True
MSFlexGrid1.SetFocus
DoEvents

'Now enable the toolbar buttons
EnableToolBarButtons
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Not Button = vbLeftButton Then Exit Sub
If Abs(MouseDownPosition.x - x) > 200 Or Abs(MouseDownPosition.y - y) > 200 Then
    'Keep track of present folder
        SSTree1(0).Tag = SSTree1(0).SelectedItem.Index
        MSFlexGrid1.DragIcon = ImgList(4).Picture
        MSFlexGrid1.Drag vbBeginDrag
Else
    MSFlexGrid1.MousePointer = flexArrow
End If

End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim LastRow As Integer
Dim FirstRow As Integer
Dim J As Integer
Dim PreviousRow As Integer
Dim PreviousSelection As Integer

On Error Resume Next

'Reset this for the drag and drop
If Shift = 1 Then Exit Sub
gDragCommenced = False

'And this to allows a delay before showing the drag icon

'MouseDownPosition.x = x
'MouseDownPosition.y = y

MousePointer = vbDefault

'This is better than making the grid non visible as that will fire off a lost focus
MSFlexGrid1.Visible = False

PreviousRow = Grid.SelectedRow
'If Not Shift Then
  ' PreviousSelection = 0
'Else
PreviousSelection = Grid.RowSelection
'End If
Grid.SelectedCol = MSFlexGrid1.Col
Grid.RowSelection = MSFlexGrid1.RowSel
Grid.SelectedRow = MSFlexGrid1.Row

If Not PreviousSelection = 0 Then
    If PreviousRow < PreviousSelection Then
        FirstRow = PreviousRow
        LastRow = PreviousSelection
    Else
        FirstRow = PreviousSelection
        LastRow = PreviousRow
    End If
    'Blank out previous selection
    For J = FirstRow To LastRow
        For i = 0 To 4
            MSFlexGrid1.Row = J
            MSFlexGrid1.Col = i
            MSFlexGrid1.CellBackColor = vbWhite
            MSFlexGrid1.CellForeColor = vbBlack
        Next
    Next
End If

If Shift = 0 Then
   ' MSFlexGrid1.Clear
    MSFlexGrid1.Row = PreviousRow
    For i = 0 To 4
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbWhite
        MSFlexGrid1.CellForeColor = vbBlack
    Next
    MSFlexGrid1.Row = Grid.SelectedRow
    For i = 0 To 4
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbHighlight
        MSFlexGrid1.CellForeColor = vbWhite
    Next
Else
    
    If PreviousRow < Grid.RowSelection Then
        FirstRow = PreviousRow
        LastRow = Grid.RowSelection
    Else
         FirstRow = Grid.RowSelection
        LastRow = PreviousRow
    End If
    
    For J = FirstRow To LastRow
    For i = 0 To 4
        MSFlexGrid1.Row = J
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbHighlight
        MSFlexGrid1.CellForeColor = vbWhite
    Next
Next

End If
'MSFlexGrid1.Enabled = True
'
MSFlexGrid1.Visible = True
MSFlexGrid1.SetFocus
DoEvents

'Now enable the toolbar buttons
EnableToolBarButtons
End Sub

Private Sub mTools_EmailScan_Click()
Dim sRet As String
Dim bLicenceStatus As Boolean
On Error Resume Next
'bLicenceStatus = CheckLicenceExpired
Timer1.Enabled = False
    frmStatus.lblTransferStatus = ""
    frmStatus.fMode = "Sending Messages in OutBox.."
    frmStatus.Show
    DoEvents
'    If bLicenceStatus Then frmAbout.Show vbModal
    ScanOutBoxForMessages
    If Not giEmailScanOption = SCAN_NONE Then
        frmStatus.lblTransferStatus = ""
        frmStatus.fMode = "Looking on the server for incoming messages.."
        ScanForMessages
    End If
Unload frmStatus
If giEmailScanInterval > 0 Then Timer1.Enabled = True

End Sub


Private Sub mToolsAddressBook_Click()
frmEditAddressBook.Show vbModal
End Sub
Private Sub mViewHeaders_Click()
frmDisplayMessageHeader.Show
End Sub

Private Sub mViewIMAPMailboxes_Click()
frmIMAPMailBoxList.Show
ConnectIMAP4
ListIMAPMailBoxes
DisconnectIMAP4
End Sub

Public Function MXRecord(sEmailAddress As String) As String
Dim EndTime    As Variant
Dim nTimeout As Variant
sMailServer = ""
MX1.WinsockLoaded = True
nTimeout = 10
If MX1.DNSServer = "" Then MX1.DNSServer = MailConnector.DNSServerName
MX1.EmailAddress = sEmailAddress
'-- Wait for the specified period of time
    '   for the connection to be made
EndTime = DateAdd("s", nTimeout, Now)
Do
    MX1.DoEvents
    If Now >= EndTime Then
            '-- Time's up. Exit with timeout Error
        MsgBox "Error obtaining MX Records"
        MX1.WinsockLoaded = False
        Exit Function
    End If
Loop Until Not sMailServer = ""
MXRecord = sMailServer
MX1.WinsockLoaded = False
End Function
Private Sub MX1_Response(RequestId As Integer, Domain As String, MailServer As String, Precedence As Integer, TimeToLive As Long, StatusCode As Integer, Description As String)
If Not sMailServer = "" Then Exit Sub
sMailServer = MailServer
End Sub

Private Sub POP1_Error(ErrorCode As Integer, Description As String)
MsgBox "POP error occured - " & Description, vbCritical
Beep
End Sub

Private Sub POP1_Header(Field As String, Value As String)
 Select Case UCase(Field)
    Case "SUBJECT": gMessageRecord.Subject = Value
    Case "FROM": gMessageRecord.From = Value
    Case "DATE": gMessageRecord.SentDate = Value
    Case "RECEIVED": gMessageRecord.Received = Value
    Case "TO": gMessageRecord.To = Value
    Case "CC": gMessageRecord.CC = Value
    Case "REPLY-TO": gMessageRecord.ReplyTo = Value
    Case "MESSAGE-ID": gMessageRecord.MessageID = Value
    Case "RETURN-PATH": gMessageRecord.ReturnPath = Value

End Select
'We also need this
gMessageRecord.Header = gMessageRecord.Header & Field & ": " & Value & vbCrLf
End Sub

Private Sub POP1_PITrail(Direction As Integer, Message As String)
On Error GoTo POPerror1
If Message = "Connected" Then frmStatus.lblTransferStatus = "Connected..."
Exit Sub
POPerror1:
    MsgBox Err.Description & " - Your connection lost...", vbApplicationModal + vbCritical, "Connection Error"
    Err.Clear
End Sub

Private Sub POP1_Transfer(BytesTransferred As Long, Text As String)
Dim ptrText As Long
Dim sText As String
    frmStatus.lblBytesTransferred.Caption = CStr(BytesTransferred)
    If gBuffer.hMem = 0 Then
        DoEvents
        gMessage = gMessage & Text & vbCrLf
    Else
        If Len(Text) > 8046 Then
            ShowStatus 1, "Buffer overflow error..."
            Beep
            Exit Sub
        End If
        sText = Text & vbCrLf
        gBuffer.BufferSize = CLng(Len(sText))
        Call agCopyData(ByVal sText, ByVal gBuffer.Address, gBuffer.BufferSize)
        gBuffer.Address = gBuffer.Address + CLng(Len(sText))
    End If
End Sub





Private Sub popupNewGroupFolder_Click()
mFile_NewFolderGroup_Click
End Sub

Private Sub popupNewSubFolder_Click()
CreateSubFolder
DisplayInBox
End Sub

Private Sub SMTP1_Error(ErrorCode As Integer, Description As String)
Dim s As String
s = s
End Sub

Private Sub SMTP1_Transfer(BytesTransferred As Long)
frmStatus.lblBytesTransferred.Caption = CStr(BytesTransferred)
DoEvents
End Sub

Private Sub SSPanel1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
MousePointer = vbNoDrop

End Sub

Private Sub SSPanel1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MousePointer = vbDefault
End Sub

Private Sub SSRibbon1_Click(Index As Integer, Value As Integer)
Static SplitterToggle As Boolean ' To simulate a pushbutton just reset the value to false
Dim iCurrentMessage As Integer ' and it will pop up
gCancelAction = False
'Dim piform(10) As New frmPI
If Value = True Then
    
    Select Case Index
        Case 0
            Me.MousePointer = vbHourglass
            
            gComposeMode = True
            CreateNewPIInstance
            Me.MousePointer = vbDefault
            'gpiform(2).Show
        Case 1
            DeleteMessage
        Case 2
            Me.MousePointer = vbHourglass
            gComposeMode = False
            iCurrentMessage = DisplayMessage
            ReplyToSender (iCurrentMessage)
            Me.MousePointer = vbDefault
        Case 3
        Case 4
            CreateSubFolder
            If Not gCancelAction Then frmMain.DisplayInBox
            
        Case 5
            Me.MousePointer = vbHourglass
            gComposeMode = False
            iCurrentMessage = DisplayMessage
            ForwardMessage (iCurrentMessage)
            Me.MousePointer = vbDefault
            
         Case 6
            Me.MousePointer = vbHourglass
            mTools_EmailScan_Click
            Me.MousePointer = vbDefault
            
        Case 7
           If Not PGP_SDKPresent Then
                MsgBox "PGP is not present, can't load FileSafe.  You need PGP Version 5 or Version 6.", vbApplicationModal + vbCritical, "File Safe Open Failure"
            Else
                frmFileSafe.Show
            End If
           
    End Select
End If
SSRibbon1(Index).Value = False
End Sub

Private Sub SSTree1_AfterLabelEdit(Index As Integer, Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)
Dim rs As Recordset
Dim lFolderIndex As Long
Dim n As SSNode

On Error GoTo EditError

If NewString = "" Then
    Beep
    Cancel = True
    Exit Sub
End If

'
'First Find the node name
'
lFolderIndex = SSTree1(Index).SelectedItem.Index
Set n = SSTree1(Index).Nodes.Item(lFolderIndex)

'
'If parent is nothing then we are on a node
'
If Not n.Parent Is Nothing Then
    If Index = 0 Then
        Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    Else
        Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
    End If
    
    rs.FindFirst "[Folder] =" & "'" & StripItemCount(n.Text) & "'"
    If Not rs.NoMatch Then
        If rs("Can Delete") Then
            rs.Edit
            rs("Folder") = StripItemCount(NewString)
            rs.Update
        Else
            Beep
            MsgBox "You can't rename a system folder.", vbApplicationModal + vbCritical, "Rename Error"
            Cancel = True
        End If
    Else
        Cancel = True
    End If
Else
    If Index = 0 Then
        Set rs = DB.OpenRecordset("Nodes", dbOpenDynaset)
    Else
        Set rs = DB.OpenRecordset("File Nodes", dbOpenDynaset)
    End If
    rs.FindFirst "[Node Name] =" & "'" & StripItemCount(n.Text) & "'"
    If Not rs.NoMatch Then
        If rs("Can Delete") Then
            rs.Edit
            rs("Node Name") = NewString
            rs.Update
        Else
            Beep
            MsgBox "You can't rename a system folder", vbApplicationModal + vbCritical, "Rename Error"
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If

rs.Close
Set rs = Nothing
Exit Sub
EditError:
    Cancel = True
    MsgBox "Could not edit the folder name:  " & Err.Description, vbApplicationModal + vbCritical
    Err.Clear
End Sub

Private Sub SSTree1_AfterNodeDelete(Index As Integer, Node As SSActiveTreeView.SSNode)
Dim rsNode As Recordset
Dim rsFolder As Recordset
Dim rsMessage As Recordset
Dim qd As QueryDef

If m_GroupNodeMarkedForDelete Then
    'We are going to delete a group folder here
  '  If Index = 0 Then
        Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
   ' Else
   '     Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
   ' End If
    rsNode.FindFirst "[Node Name] =" & "'" & Node.Text & "'"
    If Not rsNode.NoMatch Then
        rsNode.Delete
    Else
        Beep
        MsgBox "Folder is not in the database.  PI is continuing anyway."
    End If
Else
    'Okay we are going to delete a folder here
    'If Index = 0 Then
        Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
   ' Else
   '     Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
   ' End If
    rsFolder.FindFirst "[Folder] =" & "'" & StripItemCount(Node.Text) & "'"
    
    If Not rsFolder.NoMatch Then
       ' If Index = 0 Then
            Set qd = DB.QueryDefs("qdMessagesinFolder")
            qd.Parameters![FolderId] = rsFolder("Folder ID")
       ' Else
       '     Set qd = DB.QueryDefs("qdFilesinFolder")
       '     qd.Parameters![FolderId] = rsFolder("Folder ID")
      '  End If
        Set rsMessage = qd.OpenRecordset()
        Do While Not rsMessage.EOF
            rsMessage.Delete
            rsMessage.MoveNext
        Loop
        rsFolder.Delete
    Else
        Beep
        MsgBox "Folder is not in the database.  PI is continuing anyway."
    End If
End If

Set rsNode = Nothing
Set rsFolder = Nothing
Set rsMessage = Nothing
'If Index = 0 Then
   Dim s As String
    's = Node.Text
    'Node.Selected = True
    'InitialiseGrid ("From/To")
    'FillGrid
'Else
'    FillListView
'End If
End Sub


Private Sub SSTree1_BeforeLabelEdit(Index As Integer, Cancel As SSActiveTreeView.SSReturnBoolean)
Dim n As SSNode
Dim lFolderIndex As Long
Static sPreviousNode As String
'
'First Find the node name
'
lFolderIndex = SSTree1(Index).SelectedItem.Index
Set n = SSTree1(Index).Nodes.Item(lFolderIndex)
If Not sPreviousNode = n.Text Then
    sPreviousNode = n.Text
    Cancel = True
Else
    'Cancel = False
    sPreviousNode = ""
End If

End Sub

Private Sub SSTree1_BeforeNodeDelete(Index As Integer, Node As SSActiveTreeView.SSNode, Cancel As SSActiveTreeView.SSReturnBoolean, DispPromptMsg As SSActiveTreeView.SSReturnBoolean)
'Dim rsFolder As Recordset
Dim rsNode As Recordset
Dim rsFolder As Recordset
Dim rsMessage As Recordset
Dim qd As QueryDef
'Look for root or node
Dim n As SSNode
Dim lFolderIndex As Long

'
'First Find the node name
'
'lFolderIndex = SSTree1(Index).SelectedItem.Index
'Set n = SSTree1(Index).Nodes.Item(lFolderIndex)
'n.FirstSibling.


m_GroupNodeMarkedForDelete = False
'Tree(Index).Initialise
If Not Node.Children = 0 Then
    MsgBox "Delete sub-folders first", vbApplicationModal + vbCritical
    Cancel = True
    Exit Sub
End If

'Okay now search the database for the sub folder or group name
If Node.Parent Is Nothing Then
    'We are going to delete a group folder here
   ' If Index = 0 Then
        Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
    'Else
     '   Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
   ' End If
    rsNode.FindFirst "[Node Name] =" & "'" & StripItemCount(Node.Text) & "'"
    
   
    
    If Not rsNode.NoMatch Then
        If Not rsNode("Can Delete") Then
            Beep
            Cancel = True
            MsgBox "Can't delete Group Folder", vbQuestion + vbApplicationModal, "Group Folder"
        Else
           ' rsNode.Delete
           m_GroupNodeMarkedForDelete = True
        End If
    Else
            Beep
            MsgBox "Group Folder is not in the database.  PI is continuing anyway."
    End If
Else
    'Okay we are going to delete a folder here
    'If Index = 0 Then
        Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
   ' Else
   '     Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
   ' End If
    Dim s As String
    
   's = rsFolder("Folder")
   
    rsFolder.FindFirst "[Folder] =" & "'" & StripItemCount(Node.Text) & "'"
    If Not rsFolder.NoMatch Then
        If Not rsFolder("Can Delete") Then
            Beep
            Cancel = True
            MsgBox "Can't delete a system folder", vbQuestion + vbApplicationModal, "System Folder"
        Else
            m_GroupNodeMarkedForDelete = False
          ' Set qd = DB.QueryDefs("qdMessagesinFolder")
           ' qd.Parameters![FolderId] = rsFolder("Folder ID")

           ' Set rsMessage = qd.OpenRecordset()
           ' Do While Not rsMessage.EOF
           '     rsMessage.Delete
           '     rsMessage.MoveNext
           ' Loop
           ' rsFolder.Delete
            End If
    Else
        Beep
        MsgBox "Folder is not in the database.  PI is continuing anyway."
    End If
End If

'rsNode.Close
Set rsNode = Nothing
Set rsFolder = Nothing
End Sub


Private Sub SSTree1_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
Dim ssNodeTmp As SSNode
Dim ssNodeX As SSNode
Dim Sqlq As String
Dim i As Integer
Dim lWidth As Long
Dim Heading As String
Dim NewFolderID As Long
Dim rs As Recordset
Dim StoredFileName As String
Dim StartRow As Long
Dim EndRow As Long

On Error Resume Next
MousePointer = vbHourglass
DoEvents
Set ssNodeTmp = SSTree1(Index).HitTest(x, y)
If ssNodeTmp Is Nothing Then Exit Sub
'Set SSTree1(Index).SelectedItem = ssNodeTmp
If Source.name = "MSFlexGrid1" Or Source.name = "lvFileListView" Then
    'SSTree1.DragIcon = ImgList(5).Picture 'LoadPicture("d:\pi32\icons\drop1pg.ico") 'ImageList1(0) 'drop icon
    'gSelectedFolder = Node
    If Index = 0 Then
        Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    Else
        Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
    End If
    rs.FindFirst "Folder =" & "'" & StripItemCount(ssNodeTmp.Text) & "'"
    If rs.NoMatch Then
        rs.Close
        Set rs = Nothing
        Beep
        Exit Sub
    Else
        NewFolderID = rs("Folder ID")
    End If
    rs.Close
'Update

    If Index = 0 Then
        Set rs = DB.OpenRecordset(Grid.SelectedQuery, dbOpenDynaset)
        If rs.EOF Then
            Beep
            Set rs = Nothing
            Exit Sub
        Else
            rs.MoveFirst
        End If
    If Grid.RowSelection >= Grid.SelectedRow Then
        StartRow = Grid.SelectedRow
        EndRow = Grid.RowSelection
    Else
        EndRow = Grid.SelectedRow
        StartRow = Grid.RowSelection
    End If


For i = 1 To StartRow - 1
    rs.MoveNext
Next

For i = StartRow To EndRow
        rs.Edit
        rs("Folder ID") = NewFolderID
        rs.Update
        rs.MoveNext
Next
rs.Close
   
    'Update FlexGrid
        
    MSFlexGrid1.Row = StartRow
    For i = StartRow To EndRow
         DoEvents
        If MSFlexGrid1.Rows = 2 Then
            MSFlexGrid1.Clear
        Else
            MSFlexGrid1.RemoveItem MSFlexGrid1.Row
        End If
    Next
        'Now select next item sof the user has a reference
        'Now highlight selected cell
   
        If Grid.SelectedRow <= MSFlexGrid1.Rows - 1 Then
            MSFlexGrid1.Row = Grid.SelectedRow
        End If
        'Highlight the selected row
        For i = 0 To 4
            MSFlexGrid1.Col = i
            MSFlexGrid1.CellBackColor = vbHighlight
            MSFlexGrid1.CellForeColor = vbWhite
        Next
    Else
       
    End If
            
End If
Set rs = Nothing
SSTree1(Index).SelectedItem = SSTree1(Index).Nodes(CInt(SSTree1(Index).Tag))
If Index = 0 Then
   ' FillGrid
    MSFlexGrid1.Drag vbEndDrag
End If
AddFolderParameters (Index)
MSFlexGrid1.SetFocus
gDragCommenced = False
ssNodeTmp.Selected = False
Call Form_Resize
MousePointer = vbDefault
ShowStatus 1, "Completed successfully.."
'AddFolderParameters (Index)
End Sub

Private Sub SSTree1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
Dim ssNodeTmp As SSNode

If Index = 1 Then
    If Not Source.name = "lvFileListView" Then
        SSTree1(Index).Drag 0
        Exit Sub
    End If
End If
If Index = 0 Then
    If Not Source.name = "MSFlexGrid1" Then
        SSTree1(Index).Drag 0
        Exit Sub
    End If
End If
SSTree1(Index).SetFocus
Set ssNodeTmp = SSTree1(Index).HitTest(x, y)
If Not ssNodeTmp Is Nothing Then
    Set SSTree1(Index).SelectedItem = ssNodeTmp
    Set SSTree1(Index).DropHighlight = Nothing
    'ssNodeTmp.Font.bold = True
End If

End Sub

Private Sub SSTree1_Expand(Index As Integer, Node As SSActiveTreeView.SSNode)
'Dim rs As Recordset

If Not Node.Parent Is Nothing Then
        InitialiseGrid ("From/To")
        FillGrid
End If
End Sub


Private Sub SSTree1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As String
 If Button = 2 Then   ' Check if right mouse button
                       ' was clicked.
      PopupMenu mnuTreeListContext   ' Display the File menu as a
                        ' pop-up menu.
End If

'If Button = 2 And Not Tree(Index).NodeName = "" Then
  '  Tree(Index).CancelClick = True
'End If
End Sub

Private Sub SSTree1_NodeClick(Index As Integer, Node As SSActiveTreeView.SSNode)
Dim rs As Recordset
Dim lFolderIndex As Long
Dim n As SSNode

On Error GoTo NodeError

If Index = 0 Then Grid.Initialise

If Node.Parent Is Nothing Then Exit Sub
 
 'Need to do this because expand is not firing
 InitialiseGrid ("From/To")
 FillGrid
 MousePointer = vbDefault
 AddFolderParameters (Index)
 'MSFlexGrid1.
Exit Sub
NodeError:
    'PIForm(gActivePIInstance).ShowStatus 1, "Error: " & Err.Description
    Err.Clear
    MousePointer = vbDefault
End Sub

Private Sub SSTree1_OLEDragOver(Index As Integer, Data As SSActiveTreeView.SSDataObject, Effect As SSActiveTreeView.SSReturnLong, Button As Integer, Shift As Integer, x As Single, y As Single, State As SSActiveTreeView.SSReturnShort)

Set SSTree1(Index).DropHighlight = SSTree1(Index).HitTest(x, y)

    If Not SSTree1(Index).DropHighlight Is Nothing Then
        'SSTree1(Index).Nodes.Item (2)
        If SSTree1(Index).DropHighlight.Level < 2 Then
            Effect = ssOLEDropEffectNone
              Set SSTree1(Index).DropHighlight = Nothing
            Exit Sub
        End If

        Effect = ssOLEDropEffectCopy
    End If
  
End Sub

Private Sub SSTree1_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer
i = Index
End Sub

Private Sub Timer1_Timer()
'If Index = TIMER_DRAG_DROP Then
   ' Timer1(TIMER_DRAG_DROP).Enabled = False
    'MousePointer = vbDefault
'End If
'If Index = 1 Then
    giTimerCounter = giTimerCounter - 1
    If giTimerCounter = 0 Then
        Timer1.Enabled = False
        mTools_EmailScan_Click
        giTimerCounter = giEmailScanInterval
        Timer1.Enabled = True
        'MSFlexGrid1.Enabled = True
    End If
'End If
End Sub



Public Sub InitialiseGrid(Optional Direction As String)
Dim Headings As String
Dim lWidth As Long

'Grid.Initialise
If Direction = "" Then Direction = "From/To"
MSFlexGrid1.Clear
MSFlexGrid1.Refresh
Headings = "^Read|^Att|<" & Direction & "|<Subject|<Date Sent/Received"
MSFlexGrid1.Cols = 5
MSFlexGrid1.Rows = 1
'MSFlexGrid1.RowSel = 1
'MSFlexGrid1.CellBackColor = vbBlue
MSFlexGrid1.FormatString = Headings
Call Form_Resize
End Sub

Public Function ConnectPOP() As Long
    On Error GoTo ConnectError
    
    If Not MailConnector.MailServerName = "" Then
        If Not POP1.WinsockLoaded Then POP1.WinsockLoaded = True
        POP1.MailServer = MailConnector.MailServerName
        POP1.User = MailConnector.AccountName
        POP1.Password = MailConnector.AccountPassword
        If Not MailConnector.POPPort = 0 Then POP1.MailPort = MailConnector.POPPort
        POP1.Action = 1
        ConnectPOP = POP1.MessageCount
    Else
        MsgBox ("A POP mail server hasn't been specified.")
    End If
    Exit Function

ConnectError:
    Select Case Err
        Case 20172
            '---------------------------------------------
            'Invalid Password.
            '---------------------------------------------
            MsgBox "Password you specified is not valid."
        Case 25066
            MsgBox "This is an ISP connection error.  Apparently, you have become disconnected from your Internet Service Provider.  Please connect to your ISP and retry."
        
        Case 26005
             End Select
   ' MailConnector.AccountPassword = "ERROR"
    Err.Clear
End Function
Sub GetPOPMessage(msgNumber As Integer)
Dim MsgSize As Long
    On Error GoTo GetPOPError
    gMessageRecord.Header = ""
    POP1.MessageNumber = msgNumber
    gMessageRecord.MessageSize = POP1.MessageSize
    gMessageRecord.MessageNumber = msgNumber
    POP1.MaxLines = 0
    
    'Allocate the maximum -
    MsgSize = CLng(POP1.MessageSize) * 10
    If AllocateMemory(MsgSize) Then
        POP1.Action = a_Retrieve
        'This is the actual size of the message
        'MsgSize = gBuffer.BufferSize
        gMessage = String(MsgSize, Chr(0))
        Call agCopyData(ByVal gBuffer.StartAddr, ByVal gMessage, MsgSize)
        
        If Not FreeMemory Then
                Err.Raise 99999, , "Can't free memory - you should terminate the application and restart."
        End If
    Else
        Err.Raise 99999, , "Can't allocate sufficient memory!"
    End If
    Exit Sub

GetPOPError:
    MsgBox "Error returned by POP: " & Err.Description & " (GetPOPMessage)", vbApplicationModal + vbCritical, "POP Error"
    POP1.Action = 0
    Err.Clear
    Beep
End Sub

Public Function GetPOPTop(MsgNum As Integer) As Boolean
   
    On Error GoTo POPTopError
    gMessage = ""
    gMessageRecord.EndTransfer = False
    POP1.MaxLines = 5
    POP1.MessageNumber = MsgNum
    POP1.Action = 3
    gMessageRecord.MessageSize = POP1.MessageSize
    GetPOPTop = True
    
Exit Function

POPTopError:
    Select Case Err
    Case 20172
            '---------------------------------------------
            'Invalid Password.
            '---------------------------------------------
            Err.Clear
            Unload frmStatus
            MsgBox "Password you specified is not valid.", vbApplicationModal + vbCritical, "Password Erorr"
            frmMailServerOptions.Show vbModal
            Beep
    Case 25058
            '---------------------------------------------
            'Lost the socket.
            '---------------------------------------------
            MsgBox "The connection was lost.", vbApplicationModal + vbCritical, "Connection Erorr"
            Beep
    Case Else
            MsgBox "The error: " & Err.Description & " occurred in (GetPOPTop)", vbApplicationModal + vbCritical, "Connection Erorr"
        Beep
    End Select
    POP1.Action = a_Idle
    Err.Clear
    GetPOPTop = False
End Function
Function ScanTopHeaders(search As String) As String
       
    'search - a string to search for within the header
    'ScanTopHeaders - returns a string containing found message numbers delimited
    'by spaces
    Dim NumMessages As Integer
    Dim tmpstr As String
    Dim J As Integer
    
    NumMessages = 0
    tmpstr = ""
    For J = 1 To MailConnector.NumMessages
        DoEvents
        If search = "" Then
            'This is a simple fix to get all messages
            NumMessages = NumMessages + 1
            tmpstr = tmpstr & Format(J) & " "
        Else
            If GetPOPTop(J) Then
                DoEvents
                gMessage = Mid$(gMessage, 1, Len(gMessage))
                If InStr(1, gMessage, search) Then
                    NumMessages = NumMessages + 1
                    tmpstr = tmpstr & Format(J) & " "
                End If
            End If
        End If
    Next
    ScanTopHeaders = RTrim$(tmpstr)
End Function

Sub ParseFoundMessages(TheMessages As String, TheDeleteMessages As String)
       
    Dim TempMessages As String
    Dim CurrentMessage As String
    Dim CurrentDeleteMessage As String
    Dim SectionName As String
    Dim res As Long
    Dim i As Integer
    Dim J As Integer
    Dim bzap As Boolean
    Dim tmpstr As String
        
    Dim lFolderIndex As Long
    Dim lFolderID As Long
    Dim n As SSNode

    On Error GoTo ParseError
'First Find the node name
'
    lFolderIndex = SSTree1(0).SelectedItem.Index
    Set n = SSTree1(0).Nodes.Item(lFolderIndex)
        
        
    MousePointer = vbHourglass
    DoEvents
    If TheMessages <> "" Then
      
        TempMessages = TheMessages
        On Error GoTo ParseError
        'see if messages should be deleted from the server after download
        SectionName = "Options"
        tmpstr = ReadProfile(SectionName, "ServerDelete")
        If tmpstr = "False" Then
            bzap = False
        Else
            bzap = True
        End If
        'parse through the string, extracting each message number
        i = 1
        J = 1
        Do Until i = 0
            i = InStr(1, TempMessages, " ")
            J = InStr(1, TheDeleteMessages, " ")
            If i > 1 Then
                CurrentMessage = Mid$(TempMessages, 1, i - 1)
            Else
                CurrentMessage = TempMessages
            End If
            If J > 1 Then
                CurrentDeleteMessage = Mid$(TheDeleteMessages, 1, J - 1)
            Else
                CurrentDeleteMessage = TheDeleteMessages
            End If
            
            DoEvents
                
            If Not CurrentMessage = CurrentDeleteMessage Then
                If MailConnector.ConnectUsing = CONNECT_POP3 Then
                    GetPOPMessage (CurrentMessage)
                Else
                    GetIMAPMessage (CurrentMessage)
                End If
                'write it to database
                frmStatus.lblTransferStatus.Caption = "Saving message number " & CurrentMessage & "."
                ShowStatus 1, "Saving message number " & CurrentMessage & "."
                WriteMessageRecord
                If StripItemCount(UCase(n.Text)) = UCase("Inbox") Then FillGrid
            
            'now delete all messages from the server from the server - if option is set
            Else
                frmStatus.lblTransferStatus.Caption = "Deleting message number " & CurrentMessage & "."
                ShowStatus 1, "Deleting message number " & CurrentMessage & "."
            End If
            If bzap = True Then
            If MailConnector.ConnectUsing = CONNECT_POP3 Then
                DeletePOPMessage (CurrentMessage)
            Else
                DeleteIMAPMessage (CurrentMessage)
            End If
                
            End If
            'bump the string
            TempMessages = Mid$(TempMessages, i + 1, Len(TempMessages) - i)
        Loop
       ' Beep   'don't show  frmMailBox.Show just beep to let them know
    End If
    MousePointer = vbDefault
    ShowStatus 1, ""
    Exit Sub

ParseError:
    MousePointer = vbDefault
    MsgBox Err.Description & " (ParseFoundMessages)"
    ShowStatus 1, ""
    Err.Clear
End Sub

Private Sub DeleteMessage()
Dim i As Integer
Dim qd As QueryDef
Dim StartRow As Integer
Dim EndRow As Integer
Dim DeletedFolderID As Long
Dim rsAttachment As Recordset
Dim iResponse As Integer
Static InHere As Boolean
Dim Sqlq As String
Dim rs As Recordset
Dim lFolderIndex As Long
Dim n As SSNode
 
 If InHere Then Exit Sub
'
'First Find the node name
'
lFolderIndex = SSTree1(0).SelectedItem.Index
Set n = SSTree1(0).Nodes.Item(lFolderIndex)


InHere = True
On Error GoTo BadKeyEntry
Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
rs.FindFirst "Folder =" & "'" & "Deleted Items" & "'"
If rs.NoMatch Then
    Err.Raise 1002, "System Folder has been deleted"
Else
    DeletedFolderID = rs("Folder ID")
End If
rs.Close

Set rs = DB.OpenRecordset(Grid.SelectedQuery, dbOpenDynaset)
If rs.EOF Then
    Beep
    InHere = False
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.MoveFirst
End If


If Grid.RowSelection >= Grid.SelectedRow Then
    StartRow = Grid.SelectedRow
    EndRow = Grid.RowSelection
Else
    EndRow = Grid.SelectedRow
    StartRow = Grid.RowSelection
End If


For i = 1 To StartRow - 1
    rs.MoveNext
Next
'Don't ask - just delete
'Let's keep going on an error
On Error Resume Next
For i = StartRow To EndRow
        If StripItemCount(n.Text) = "Deleted Items" Then
            Kill App.Path & "\mailbox\" & rs("MIME Message")
            rs.Delete
            rs.MoveNext
        Else
            'Place in deleted folder
            rs.Edit
            rs("Folder ID") = DeletedFolderID
            rs.Update
            rs.MoveNext
        End If
Next
rs.Close
Set rs = Nothing
MSFlexGrid1.Row = StartRow
For i = StartRow To EndRow
    If MSFlexGrid1.Rows = 2 Then
        MSFlexGrid1.Clear
    Else
        MSFlexGrid1.RemoveItem MSFlexGrid1.Row
    End If
Next
'Grid.RowSelection = MSFlexGrid1.Row
InHere = False
AddFolderParameters (0)

Exit Sub

BadKeyEntry:
    MsgBox Err.Description & " - There was an error deleting the file."
    Beep
    Err.Clear
    InHere = False
    FillGrid
End Sub

Private Sub FillGrid()
Dim rs As Recordset
Dim pos1 As Integer
Dim pos2 As Integer
Dim Sqlq As String
Dim RowIndex As Integer
Dim ItemString As String
Dim FromString As String
Dim ToString As String
Dim i As Integer
Dim sName As String
Dim lFolderIndex As Long
Dim lFolderID As Long
Dim n As SSNode
Dim HideUpdate As Boolean

On Error GoTo FillGridError
'
'First Find the node name
'
If (SSTree1(0).SelectedItem Is Nothing) Then Exit Sub
lFolderIndex = SSTree1(0).SelectedItem.Index
Set n = SSTree1(0).Nodes.Item(lFolderIndex)
'Grid.SelectedRow
If n.Index < 3 Then Exit Sub
'
'Find the Node Id
'
'If Index = 0 Then
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
'Else
  ' Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
'End If
If Not n.Text = "" Then
    rs.FindFirst "Folder =" & "'" & StripItemCount(n.Text) & "'"
    If rs.NoMatch Then
        MsgBox "Expand: Can't find folder- Database error", vbApplicationModal + vbCritical, "Expand"
        Exit Sub
    Else
        lFolderID = rs("Folder ID")
    End If
End If

rs.Close
Set rs = Nothing

'
'Now fill the grid
'
Me.MousePointer = vbHourglass

Select Case Grid.SelectedColToSort
    Case 0
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Folder ID] = " & lFolderID & ") " 'And (Messages.Read = True) "
        Sqlq = Sqlq & "ORDER BY Messages.[Message Read] DESC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
    Case 1
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Folder ID] = " & lFolderID & ") " 'And (Messages.Attachment = True) "
        Sqlq = Sqlq & "ORDER BY Messages.Attachment  ASC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
    Case 2
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Folder ID] = " & lFolderID & ") " 'And (Messages.Attachment = True) "
        Sqlq = Sqlq & "ORDER BY Messages.From  ASC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
    Case 3
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Folder ID] = " & lFolderID & ")"
        Sqlq = Sqlq & "ORDER BY Messages.Subject ASC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
   ' Case 5
    Case 4
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Folder ID] = " & lFolderID & ")"
        Sqlq = Sqlq & "ORDER BY Messages.[Date Sent] DESC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
    Case 5
    ' Else
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Folder ID] = " & lFolderID & ")"
        Sqlq = Sqlq & "ORDER BY Messages.[Date Received] DESC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
        'Me.MousePointer = vbDefault
        'Exit Sub
End Select

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Me.MousePointer = vbDefault
    ItemString = "" & vbTab & "" & vbTab & "There are no items to display"
    MSFlexGrid1.Clear
    MSFlexGrid1.AddItem ItemString, 1
    Exit Sub
End If
'Now fill the message area
'First check which mail box we are ing

Select Case UCase(StripItemCount(n.Text))
    Case UCase("Sent Items"), UCase("Outbox"), UCase("Drafts")
        InitialiseGrid ("To")
    
    Case UCase("Inbox")
        InitialiseGrid ("From")
        
    Case UCase("Deleted Items")
        InitialiseGrid ("From/To")
    
    Case Else
        InitialiseGrid ("From")
        
End Select
    

RowIndex = 1
Grid.SelectedQuery = Sqlq



HideUpdate = IIf(rs.RecordCount > 1, True, False)
If HideUpdate Then MSFlexGrid1.Visible = False

While Not rs.EOF
    MSFlexGrid1.Col = 0
    Select Case UCase(StripItemCount(n.Text))
     
     Case UCase("Inbox")
        sName = StripFullName(rs("From"))
        sName = IIf(sName = "", StripEMailAddress(rs("From")), sName)
        ItemString = "" & vbTab & "" & vbTab & sName & vbTab & rs("Subject") & vbTab & Format(rs("Date Received"), "ddd, ddddd ttttt") '"ddd d/mm/yy h:m ")
     
     Case UCase("Drafts")
        
       ' sName = StripFullName(rs("To"))
        sName = StripFullName(rs("To"))
        sName = IIf(sName = "", StripEMailAddress(rs("To")), sName)
        ItemString = "" & vbTab & "" & vbTab & sName & vbTab & rs("Subject") & vbTab & Format(rs("Date Received"), "ddd, ddddd ttttt") '"ddd d/mm/yy h:m ")
     
     Case UCase("Sent Items")
        sName = StripFullName(rs("To"))
        sName = IIf(sName = "", StripEMailAddress(rs("To")), sName)
        ItemString = "" & vbTab & "" & vbTab & sName & vbTab & rs("Subject") & vbTab & Format(rs("Date Sent"), "ddd, ddddd ttttt") '"ddd d/mm/yy h:m ")
      
    Case Else
        sName = StripFullName(rs("From"))
        sName = IIf(sName = "", StripEMailAddress(rs("To")), sName)
        ItemString = "" & vbTab & "" & vbTab & sName & vbTab & rs("Subject") & vbTab & Format(rs("Date Received"), "ddd, ddddd,ttttt") '"ddd d/mm/yy h:m ")
    
    End Select
    
    MSFlexGrid1.AddItem ItemString, RowIndex
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = RowIndex
    If rs("Message Read") Then
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Select Case rs("Message Status")
            Case PGPAnalyze_Unknown
                Set MSFlexGrid1.CellPicture = ImageList1.Overlay("Open Envelope", "Open Eye")
            Case PGPAnalyze_Encrypted
                Set MSFlexGrid1.CellPicture = ImageList1.Overlay("Open Envelope", "Key")
            Case Else
                Set MSFlexGrid1.CellPicture = ImageList1.Overlay("Mask", "Open Envelope")
        End Select
        With MSFlexGrid1
            For i = 0 To 4
                .Col = i
                .CellFontBold = False
            Next
        End With
    Else
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Select Case rs("Message Status")
            Case PGPAnalyze_Unknown
                Set MSFlexGrid1.CellPicture = ImageList1.Overlay("Closed Envelope", "Open Eye")
            Case PGPAnalyze_Encrypted
                Set MSFlexGrid1.CellPicture = ImageList1.Overlay("Closed Envelope", "Key")
            Case Else
                Set MSFlexGrid1.CellPicture = ImageList1.Overlay("Mask", "Closed Envelope")
        End Select
        With MSFlexGrid1
            For i = 0 To 4
                .Col = i
                .CellFontBold = True
            Next
        End With
    End If
    MSFlexGrid1.Col = 1
    If rs("Attachment") Then
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(Attachment).Picture
    End If
    RowIndex = RowIndex + 1
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
Call Form_Resize

'=========Don't select in Flexgrid as it caused a dual selection
'MSFlexGrid1.Row = Grid.SelectedRow
'MSFlexGrid1.Col = 0
'MSFlexGrid1.ColSel = 4
'Now select current Row and highlight it
'MSFlexGrid1.Visible = False
'For i = 0 To 4
 '   MSFlexGrid1.Col = i
 '   MSFlexGrid1.CellBackColor = vbHighlight
 '   MSFlexGrid1.CellForeColor = vbWhite
'Next
MSFlexGrid1.Visible = True
AddFolderParameters (0)
Me.MousePointer = vbDefault
Exit Sub

FillGridError:
    MsgBox "Error returned filling the message grid: " & Err.Description, vbCritical + vbApplicationModal, "Internal Error"
    Err.Clear
    Me.MousePointer = vbDefault
    MSFlexGrid1.Visible = True
End Sub


Private Sub RestoreMainSettings()
Dim SectionName As String
Dim sRes As String

'Get Mail Server Options
SectionName = "MailServerOptions"
If ReadProfile(SectionName, "MessageScanInverval") = "" Then
    giEmailScanInterval = 10
Else
    giEmailScanInterval = CInt(ReadProfile(SectionName, "MessageScanInverval"))
End If
giTimerCounter = giEmailScanInterval
If giEmailScanInterval > 0 Then Timer1.Enabled = True

If ReadProfile(SectionName, "EmailScanOption") = "" Then
    giEmailScanOption = SCAN_PGP_ONLY
Else
    giEmailScanOption = CInt(ReadProfile(SectionName, "EmailScanOption"))
End If

SectionName = "Options"
sRes = ReadProfile(SectionName, "ConnectUsing")
Select Case sRes
    Case "", "CONNECT_POP3"
        MailConnector.ConnectUsing = CONNECT_POP3
        mViewIMAPMailboxes.Enabled = False
    Case "CONNECT_IMAP4"
        MailConnector.ConnectUsing = CONNECT_IMAP4
        mViewIMAPMailboxes.Enabled = True
End Select

MailConnector.EmailAddress = ReadProfile(SectionName, "EmailAddress")
MailConnector.ReplyEmailAddress = ReadProfile(SectionName, "ReplyEmailAddress")
MailConnector.RealName = ReadProfile(SectionName, "RealName")
MailConnector.SMTPServerName = ReadProfile(SectionName, "SMTPServerName")
MailConnector.SMTPPort = ReadProfile(SectionName, "SMTPPort")
MailConnector.AuthenticationRequired = ReadProfile(SectionName, "AuthenticationRequired")
MailConnector.MailServerName = ReadProfile(SectionName, "MailServerName")
MailConnector.NNTPServerName = ReadProfile(SectionName, "NNTPServerName")
MailConnector.DNSServerName = ReadProfile(SectionName, "DNSServerName")
MailConnector.AccountName = ReadProfile(SectionName, "AccountName")
MailConnector.AccountPassword = GetPasswordFromDatabase
MailConnector.POPPort = Val(ReadProfile(SectionName, "MailPort"))

SectionName = "Options"
gPGPVersion = ReadProfile(SectionName, "PGPStatus")
'gPGPVersion = sPGPStatus

End Sub

Public Sub ScanForMessages()
Dim i As Integer
Dim sRes As String
Dim SectionName As String
    On Error GoTo ScanError
    MousePointer = vbHourglass
    DoEvents
    If CheckConnection Then
        frmStatus.lblTransferStatus.Caption = "Connecting to mail server."
        If MailConnector.ConnectUsing = CONNECT_POP3 Then
            MailConnector.NumMessages = ConnectPOP ' Connect and get number of messages
            MailConnector.ServerState = POPSTATE
        Else
            MailConnector.NumMessages = ConnectIMAP4 ' Connect and get number of messages
            MailConnector.ServerState = IMAPSTATE ' These are important if the user cancels the status
        End If
        If Not MailConnector.AccountPassword = "ERROR" Then
           'password okay
            DoEvents
            If Not MailConnector.NumMessages = 0 Then
                    frmStatus.lblTransferStatus.Caption = "Retrieving " & MailConnector.NumMessages & " messages from the mail server."
                    DoEvents
                    
                    Select Case giEmailScanOption
                        Case SCAN_PGP_ONLY
                            gFoundMessages = ScanTopHeaders("-----BEGIN PGP MESSAGE-----")
                            If gFoundMessages <> "" Then
                                SectionName = "Options"
                                sRes = ReadProfile(SectionName, "ServerPreviewMessages")
                                If sRes = "True" Then
                                    frmStatus.Hide
                                    DoEvents
                                    gMessagesToBeDeleted = ""
                                    frmPreviewMessages.gszNumMessages = MailConnector.NumMessages
                                    frmPreviewMessages.Show vbModal
                                    frmStatus.Show
                                End If
                                Call ParseFoundMessages(gFoundMessages, "")
                            End If
                        Case SCAN_ALL
                            gFoundMessages = ScanTopHeaders("")
                            If gFoundMessages <> "" Then
                                'This returns list of messages in gFoundMessages - Delete all else
                                SectionName = "Options"
                                sRes = ReadProfile(SectionName, "ServerPreviewMessages")
                                If sRes = "True" Then
                                    frmStatus.Hide
                                    DoEvents
                                    gMessagesToBeDeleted = ""
                                    frmPreviewMessages.gszNumMessages = MailConnector.NumMessages
                                    frmPreviewMessages.Show vbModal
                                    frmStatus.Show
                                End If
                                Call ParseFoundMessages(gFoundMessages, gMessagesToBeDeleted)
                            End If
                    End Select
            End If
            '---------------------------------------------
            'need this to enter the update state
            'we be done, so...
            '---------------------------------------------
            If MailConnector.ConnectUsing = CONNECT_POP3 Then
                DisconnectPOP
            Else
                DisconnectIMAP4
            End If
        Else
           'password NOT okay
           If MailConnector.ConnectUsing = CONNECT_POP3 Then
                DisconnectPOP
            Else
                DisconnectIMAP4
            End If
        End If
    End If
    MousePointer = vbDefault
    Exit Sub
ScanError:
    MousePointer = vbDefault
    Unload frmStatus
    'ShowStatus("")
    Select Case Err.Number
        Case 20113
            '---------------------------------------------
            'Last action tied up network, so disconnect
            '---------------------------------------------
        Case 25066
            '---------------------------------------------
            '[10065] No route to host.
            '---------------------------------------------
            
        Case Else
    End Select
   ' Dim s As String
    's = Err.Description
    If MailConnector.ConnectUsing = CONNECT_POP3 Then
        DisconnectPOP
    Else
        DisconnectIMAP4
    End If
    Err.Clear
End Sub

Public Sub ReplyToSender(iCurrentMessage As Integer)
If iCurrentMessage = 0 Then Exit Sub
PIForm(iCurrentMessage).WebBrowser1.Visible = False
PIForm(iCurrentMessage).MessageArea.Visible = True
PIForm(iCurrentMessage).EnableMessageFields

PIForm(iCurrentMessage).btnTo.Caption = "To"
PIForm(iCurrentMessage).txtTo.SelStart = 0
PIForm(iCurrentMessage).txtTo.Text = PIForm(iCurrentMessage).lblFrom(1).Caption
PIForm(iCurrentMessage).txtCC.Text = ""
PIForm(iCurrentMessage).cmbRemailerSelect.Enabled = True
PIForm(iCurrentMessage).MessageArea.SelStart = 0
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = "------------"
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelBold = True
PIForm(iCurrentMessage).MessageArea.SelText = "From: " & PIForm(iCurrentMessage).lblFrom(1)
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = "Subject: " & PIForm(iCurrentMessage).txtsubject
PIForm(iCurrentMessage).MessageArea.SelBold = False
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelStart = 0
PIForm(iCurrentMessage).AddressList.Initialise
'PIForm(iCurrentMessage).MessageArea.SetFocus

End Sub

Public Sub DisplayInBox()
Dim rs As Recordset
Dim objNode As SSNode
Dim i As Integer
'Dim lFolderIndex As Long
'Dim n As SSNode

'lFolderIndex = SSTree1(TreeIndex).SelectedItem.Index
'Set n = SSTree1(TreeIndex).Nodes.Item(lFolderIndex)
'Okay find the inbox object
SSTree1(0).Nodes.Item(6).Selected = True
SSTree1(0).Nodes.Item(6).Expanded = True

For i = 1 To SSTree1(0).Nodes.Count
    If Not InStr(1, UCase(SSTree1(0).Nodes(i)), UCase("InBox")) = 0 Then Exit For
Next i
'This is the Inbox Node Object
Set objNode = SSTree1(0).Nodes(i)
objNode.Selected = True

Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
rs.FindFirst "Folder =" & "'" & "InBox" & "'"
If rs.NoMatch Then
   MsgBox "In Display In Box : Can't find InBox - Database may be corrupted", vbApplicationModal + vbCritical, "Display In Box"
   ' Set rs = DB.OpenRecordset("Nodes", dbOpenDynaset)
   ' rs.FindFirst "[Node Name] =" & "'" & "InBox" & "'"
   ' If Not rs.NoMatch Then
    '    Tree(giTreeIndex).NodeName = rs("Node Name")
   ' End If
    'Tree(giTreeIndex).FolderId = 0
'Else
    'Tree(giTreeIndex).FolderId = rs("Folder ID")
End If

rs.Close
Set rs = Nothing
InitialiseGrid ("From/To")
FillGrid
End Sub

Public Sub CreateSubFolder()
Dim rsNode As Recordset
Dim rsFolder As Recordset
Dim FolderName As String
Dim n As SSNode
Dim lFolderIndex As Long
On Error GoTo SubFolderError

lFolderIndex = SSTree1(0).SelectedItem.Index
Set n = SSTree1(0).Nodes.Item(lFolderIndex)

frmFolderName.Caption = "Create New Sub Folder"
frmFolderName.lblNamePrompt = "Enter the name of the Sub Folder"
frmFolderName.Show vbModal
FolderName = frmFolderName.FolderName
Set frmFolderName = Nothing
 
'
'First Find the node name
'
If FolderName = "" Then
    gCancelAction = True
    Exit Sub
End If
'FolderName = Tree(Index).CreateFolderName

'lFolderIndex = SSTree1(Index).SelectedItem.Index
'Set n = SSTree1(Index).Nodes.Item(lFolderIndex)

'If Index = 0 Then
    Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
'Else
 '   Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
'End If
rsNode.FindFirst "[Node Name] =" & "'" & n.Text & "'"
If Not rsNode.NoMatch Then
   ' If Index = 0 Then
        Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
    'Else
    '    Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
    'End If
    rsFolder.AddNew
    rsFolder("Folder") = FolderName
    rsFolder("Can Delete") = True
    rsFolder("Node ID") = rsNode("Node ID")
    rsFolder.Update
    rsFolder.Close
    Set rsFolder = Nothing
Else
    Beep
    Err.Raise 1002, , "Can't add a sub folder to a folder!"
End If

rsNode.Close
Set rsNode = Nothing
BuildTree (0)
Exit Sub
SubFolderError:
    MsgBox "There was an error creating the folder: " & Err.Description, vbCritical + vbApplicationModal
    Err.Clear
    gCancelAction = True
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

Public Sub ScanOutBoxForMessages()
Dim rs As Recordset
Dim OutBoxFolderId As Long
Dim SentFolderId As Long
Dim qd As QueryDef
Dim Count As Integer
Dim bUseMXRecords As Boolean
Dim sRes As String
'Dim SectionNames As String

    On Error GoTo ScanMessageError
    frmStatus.lblTransferStatus.Caption = "Looking for messages in your Outbox..."
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "[Folder] =" & "'" & "Outbox" & "'"
    If rs.NoMatch Then
        Err.Raise 10002, , "Outbox folder missing from database."
    End If
    OutBoxFolderId = rs("folder id")
    rs.Close
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "[Folder] =" & "'" & "Sent Items" & "'"
    If rs.NoMatch Then
        Err.Raise 10002, , "Sent Items folder missing from database."
    End If
    SentFolderId = rs("folder id")
    rs.Close
    
    Set qd = DB.QueryDefs("qd_MessagesInOutBox")
    qd.Parameters![OutBoxID] = OutBoxFolderId
    Set rs = qd.OpenRecordset()
    If Not rs.EOF Then
        If CheckConnection Then
            'SMTP1.Hello = "itech.net.au"
            'SMTP1.MailServer = "nt4adl.itech.net.au" '"mc2.law13.hotmail.com"
            'SMTP1.MailServer = MXRecord("alex.cameron@itech.net.au")
           'Dim s As String
           's = MXRecord(rs("To"))
        'If Direct option then....
        frmStatus.lblTransferStatus.Caption = "Connecting to remote mail server."
        'SectionName = "Options"
        sRes = ReadProfile("Options", "DeliverSMTPMessagesDirect")
        If sRes = "True" Then bUseMXRecords = True
        If Not bUseMXRecords Then ConnectSMTP
        Count = 1 'First Message
        Do While Not rs.EOF
            If bUseMXRecords Then
                SMTP1.Action = a_Disconnect
                MailConnector.SMTPServerName = MXRecord(rs("To"))
                ConnectSMTP
            End If
            frmStatus.lblTransferStatus.Caption = "Sending message  " & Count & "  of  " & rs.RecordCount & "."
            DoEvents
            SMTP1.Action = a_ResetHeaders
            SMTP1.To = rs("To") '"alex@itech.net.au,pisupport@itech.net.au" '
            SMTP1.From = rs("From")
            SMTP1.Subject = rs("Subject")
            SMTP1.ReplyTo = MailConnector.ReplyEmailAddress
            SMTP1.Date = Format(rs("Date Sent"), "ddd, ddddd, ttttt") ' "dddd, dd mmm, yyyy hh:mm ")
            SMTP1.CC = rs("CC")
            SMTP1.OtherHeaders = "X-mailer: Private Idhao Email http://www.itech.net.au/pi"
            If MailConnector.EmailAddress = "" Then
                MailConnector.EmailAddress = "Anonymous@mail"
            End If
            If Not IsNull(rs("MIME Message Header")) Then
                SMTP1.OtherHeaders = SMTP1.OtherHeaders & vbCrLf & rs("MIME Message Header")
               ' Clipboard.SetText rs("MIME Message Header")
            End If
            If Not IsNull(rs("MIME Message")) Then
                SMTP1.AttachedFile = App.Path & "\mailbox\" & rs("MIME Message")
            End If
            
                       
            SMTP1.Action = a_Send
            rs.Edit
            rs("Folder ID") = SentFolderId
            rs("Message Sent") = True
            rs.Update
            rs.MoveNext
            Count = Count + 1
            FillGrid
        Loop
        SMTP1.Action = a_Disconnect
    End If



    End If
        rs.Close
    Exit Sub
ScanMessageError:
    Unload frmStatus
    MsgBox "Could not send messages.  There was an SMTP Error.  Error returned was: " & Err.Description, vbApplicationModal + vbCritical, "Send Message Error"
    SMTP1.Action = a_Disconnect
    Err.Clear
End Sub
Public Sub ConnectSMTP()

Dim b64 As Base64Class
Set b64 = New Base64Class
    
    SMTP1.Action = a_ResetHeaders
    If SMTP1.WinsockLoaded = False Then SMTP1.WinsockLoaded = True
    If MailConnector.SMTPServerName = MailConnector.MailServerName Then
        SMTP1.MailServer = MailConnector.MailServerName
        SMTP1.MailPort = MailConnector.SMTPPort
    Else
        SMTP1.MailServer = IIf(MailConnector.SMTPServerName = "", MailConnector.MailServerName, MailConnector.SMTPServerName)
      SMTP1.MailPort = MailConnector.SMTPPort
    End If
    SMTP1.Action = a_Connect
    If MailConnector.AuthenticationRequired Then
        SMTP1.Command = "AUTH LOGIN" '
        SMTP1.Command = b64.EncodeString(MailConnector.AccountName)
        SMTP1.Command = b64.EncodeString(MailConnector.AccountPassword)
    End If
    Set b64 = Nothing
   ' SMTP1.Action = a_Connect
End Sub
Private Function DisplayMessage() As Integer
Dim FileName As String
Dim FileNum As Integer
Dim StartSection As Integer
Dim i As Integer
Dim J As Integer
Dim Sqlq As String
Dim rs As Recordset
Dim rsFolder As Recordset
Dim lListItem As ListItem
Dim AttachmentFileName As String
Dim MessageInstance As Integer

On Error GoTo BadMessageDisplay
Me.MousePointer = vbHourglass

'If InstanceNumber > MAX_OPEN_MESSAGES Then
  '  Beep
  '  Exit Sub
'End If
'To use this we need to keep track of the forms, ie which one is active etc
'Load piform(MessageInstance)
'InstanceNumber = InstanceNumber + 1
MessageInstance = CreateNewPIInstance
If MessageInstance = 0 Then Exit Function
'Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
'Sqlq = Sqlq & "(Messages.[Folder ID] = " & Tree.FolderId & ")"
'Sqlq = Sqlq & "ORDER BY Messages.Date;"
Set rs = DB.OpenRecordset(Grid.SelectedQuery, dbOpenDynaset)
If rs.EOF Then
    Beep
    rs.Close
    Set rs = Nothing
    Exit Function
Else
    rs.MoveFirst
End If

'Okay clear the message area

'piform(MessageInstance).MessageArea.Text = " "
PIForm(MessageInstance).MessageArea.Text = ""
For i = 1 To frmMain.MSFlexGrid1.Row - 1
    rs.MoveNext
Next

'To get here it must be the right row
'FileName = rs("Message File Name")
'First see if it is in the drafts folder
'Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
'rsFolder.FindFirst "[Folder] =" & "'" & "Drafts" & "'"
'If rsFolder.NoMatch Then
'        Err.Raise 1002, , "Draft folder missing from database."
'End If
'If rs("Folder id") = rsFolder("folder id") Then
'Keep the Message Id as reference for saving in drafts folder
PIForm(MessageInstance).MessageID = rs("Message ID")
'End If
'rsFolder.Close
FileName = rs("MIME Message")
PIForm(MessageInstance).lvwAttachments.ListItems.Clear
PIForm(MessageInstance).txtsubject.Text = rs("Subject")
PIForm(MessageInstance).txtTo.Text = rs("To")
PIForm(MessageInstance).lblFrom(1) = rs("From")
PIForm(MessageInstance).txtCC.Text = rs("CC")
FileName = App.Path & "\mailbox\" & FileName
gMessageRecord.Header = IIf(IsNull(rs("MIME Message Header")), "", rs("MIME Message Header"))
gMessage = GetFileText(FileName) 'rs("MIME Message")
rs.Edit
rs("Message Read") = True
rs.Update
rs.Close
If Not InStr(1, gMessageRecord.Header, "boundary=") = 0 Then
  '  PIForm(MessageInstance).MessageArea.Text = gMessageRecord.Header & vbCrLf & vbCrLf & gMessage
   PIForm(MessageInstance).MessageArea.Text = gMessageRecord.Header & gMessage
    'Decode the message and add attachments into listview - if any.
    PIForm(MessageInstance).DecodeMessage
Else
    PIForm(MessageInstance).MessageArea.Text = gMessage
End If
'Okay houskeeping
PIForm(MessageInstance).bDisplayMessageMode = True
PIForm(MessageInstance).Show
Me.MousePointer = vbDefault
DisplayMessage = MessageInstance
DoEvents
'Dim mdoc As htmldocument
PIForm(MessageInstance).MessageArea.SelStart = 0
PIForm(MessageInstance).MessageArea.SelLength = 1024
'If Not InStr(1, LCase(PIForm(MessageInstance).MessageArea.SelText), "<html>") = 0 Then
If Not InStr(1, LCase(PIForm(MessageInstance).AttachmentFileName), ".htm") = 0 Then
  PIForm(MessageInstance).WebBrowser1.Navigate (PIForm(MessageInstance).AttachmentFileName)
  PIForm(MessageInstance).MessageArea.Visible = False
  PIForm(MessageInstance).WebBrowser1.Visible = True
End If
PIForm(MessageInstance).MessageArea.SelLength = 0
Exit Function
BadMessageDisplay:
    Me.MousePointer = vbDefault
    MsgBox Err.Description & " Can't find message", vbApplicationModal + vbCritical, App.Title
    Err.Clear
    DisplayMessage = 0
End Function

Public Function CheckLicenceExpired() As Boolean
'Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
Dim rs As Recordset
Dim sDate As String
Dim pDate As String
Dim eDate As String 'Expiry Date
Dim iInstallDate As String 'Install Date
Dim iInstallDateCode As String
Dim iBuildDate As String  '2nd Check on date
Dim SectionName As String
Dim dirLen As String
Dim Windir As String * 256
Dim retStatus As String
Dim sRandomString As String
Dim NowDate As String
Dim i As Integer
Dim MAGIC_NUMBER As Double
Dim InstallDateOffset As Integer
Dim CurrentDateOffset As Integer
'Dim SerialNumberOffset As Integer

MAGIC_NUMBER = 573627385 '567643565
InstallDateOffset = 1
CurrentDateOffset = 2
CheckLicenceExpired = False
'Exit Function

On Error Resume Next
'
' If a registered copy then exit straight away
'
If gFullRelease = 1 Then Exit Function
 
 '
'Open the data base - Username will hode the trial period data
    Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
    'rs.Edit
   ' rs("Expired") = True
   ' rs.Update
    'rs.Close


'dirLen = GetSystemDirectory(Windir, 255)
'SetFileAttributes Mid(Windir, 1, dirLen) & "\pipgp_524.pil", FILE_ATTRIBUTE_NORMAL
'On Error Resume Next


On Error GoTo BadLicence
'This is the date of the exe build and nothing can be installed befor this

iBuildDate = "11,03,2002" ' Note this should not have a month of 12
iInstallDate = Format(Now(), "dd,mm,yyyy")
'This is the official install date

'If he wound back the clock don't do anything
If DateDiff("d", CDate(Format(iBuildDate, "dd,mm,yyyy")), _
            CDate(Format(iInstallDate, "dd,mm,yyyy"))) < 0 Then
            CheckLicenceExpired = True
   ' Exit Function
End If
If DateDiff("d", CDate(Format(iBuildDate, "dd,mm,yyyy")), _
        CDate(Format(Now(), "dd,mm,yyyy"))) > 90 Then  'Three months the expire
            CheckLicenceExpired = True
End If

iInstallDate = DatePart("yyyy", Now())
iInstallDate = iInstallDate & IIf(Len(DatePart("m", Now())) > 1, DatePart("m", Now()), "0" & DatePart("m", Now()))
iInstallDate = iInstallDate & IIf(Len(DatePart("d", Now())) > 1, DatePart("d", Now()), "0" & DatePart("d", Now()))


'This bit will be more complicated late - write to dll etc
Dim lStatus As Long
lStatus = IIf(IsNull(rs("Username")), 0, 1) ', SetFileAttributes(Mid(Windir, 1, dirLen) & "\pipgp_529.pil", FILE_ATTRIBUTE_NORMAL)

'sRandomString = rs("UserName") 'GetFileText(Mid(Windir, 1, dirLen) & "\pipgp_529.pil")

'Installdate is 4 bytes long
'sDate = Mid(sRandomString, (InstallDateOffset + 4) * 3 + 1, 9)

If lStatus = 0 Then 'And CheckLicenceExpired = False Then
    'New install
    sDate = Format(CStr((CLng(iInstallDate) + MAGIC_NUMBER) Xor 11011101), "00000000#")
    iInstallDate = Format(CStr((CLng(iInstallDate) + MAGIC_NUMBER) Xor 11011101), "00000000#")
    sRandomString = ""
    For i = 1 To 20
        Randomize (i + 1)
        sRandomString = sRandomString + Format(CStr(Fix(1000 * Rnd)), "00#")
        If i = CurrentDateOffset Then sRandomString = sRandomString + sDate
        If i = InstallDateOffset Then sRandomString = sRandomString + iInstallDate
    Next
    rs.Edit
    rs("Username") = sRandomString
    rs.Update
    'rs.Close
    
   ' If retStatus Then SetFileAttributes Mid(Windir, 1, dirLen) & "\pipgp_529.pil", FILE_ATTRIBUTE_HIDDEN
Else

    'First check if licence expired
    sRandomString = rs("Username")
    'Installdate is 4 bytes long
    sDate = Mid(sRandomString, (InstallDateOffset + 4) * 3 + 1, 9)
    iInstallDateCode = Mid(sRandomString, InstallDateOffset * 3 + 1, 9)
    iInstallDate = CStr(((CLng(iInstallDateCode)) Xor (11011101)) - MAGIC_NUMBER)
    iInstallDate = Mid(iInstallDate, 7, 2) & "," & Mid(iInstallDate, 5, 2) & "," & Mid(iInstallDate, 1, 4)
    
    sDate = CStr(((CLng(sDate)) Xor (11011101)) - MAGIC_NUMBER)
    sDate = Mid(sDate, 7, 2) & "," & Mid(sDate, 5, 2) & "," & Mid(sDate, 1, 4)
    If Not IsDate(sDate) Then
        CheckLicenceExpired = True
        'Exit Function
    End If

    On Error Resume Next
    If Abs(CInt(DateDiff("d", CDate(Format(iInstallDate, "dd,mm,yyyy")), Format(sDate, "dd,mm,yyyy")))) > 15 Then
        CheckLicenceExpired = True
    End If
    'Okay bump the date
        sRandomString = ""
        NowDate = DatePart("yyyy", Now())
        NowDate = NowDate & IIf(Len(DatePart("m", Now())) > 1, DatePart("m", Now()), "0" & DatePart("m", Now()))
        NowDate = NowDate & IIf(Len(DatePart("d", Now())) > 1, DatePart("d", Now()), "0" & DatePart("d", Now()))
 
       For i = 1 To 20
            Randomize (i + 1)
            sRandomString = sRandomString + Format(CStr(Fix(1000 * Rnd)), "00#")
            If i = InstallDateOffset Then sRandomString = sRandomString + iInstallDateCode
            If i = CurrentDateOffset Then sRandomString = sRandomString + Format(CStr((CLng(NowDate) + MAGIC_NUMBER) Xor 11011101), "00000000#")

        Next
        rs.Edit
        rs("Username") = sRandomString
        rs.Update
       ' rs.Close
        'retStatus = PutFileText(Mid(Windir, 1, dirLen) & "\pipgp_529.pil", sRandomString) 'Install Date
        'If retStatus Then SetFileAttributes Mid(Windir, 1, dirLen) & "\pipgp_529.pil", FILE_ATTRIBUTE_HIDDEN
'MsgBox "Debug 2", vbApplicationModal
End If
'Set a dummy write to this old file
On Error Resume Next
dirLen = GetSystemDirectory(Windir, 255)
SetFileAttributes Mid(Windir, 1, dirLen) & "\pipgp_529.pil", FILE_ATTRIBUTE_NORMAL
retStatus = PutFileText(Mid(Windir, 1, dirLen) & "\pipgp_529.pil", sRandomString)
'This is a bad serial number - if we need to use it
sRandomString = "EF8329ABF923968"
If CheckLicenceExpired Then
   ' Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
    rs.Edit
    rs("SerialNumber") = sRandomString
    rs.Update
End If
rs.Close
Exit Function
BadLicence:
    rs.Close
    CheckLicenceExpired = True
    Err.Clear
End Function

Public Function ProgramIsAlreadyRunning() As Boolean
If App.PrevInstance = True Then
    ProgramIsAlreadyRunning = True
Else
    ProgramIsAlreadyRunning = False
End If


End Function

Public Function ConnectIMAP4() As Long
    On Error GoTo ConnectError
    'If Not MailConnector.ServerConnected Then
        If Not MailConnector.MailServerName = "" Then
            IMAP1.MailServer = MailConnector.MailServerName
            IMAP1.User = MailConnector.AccountName
            IMAP1.Password = MailConnector.AccountPassword
            If Not MailConnector.POPPort = 0 Then IMAP1.POPPort = MailConnector.POPPort
            If Not IMAP1.WinsockLoaded Then IMAP1.WinsockLoaded = True
            IMAP1.Action = 2 'a_Connect
            'IMAP1.Action = a_SelectMailbox
           ' IMAP1.Mailbox = "IN*"
            'IMAP1.Action = 15 'a_ListMailboxes
            IMAP1.Mailbox = """INBOX""" 'Default Mailbox
            IMAP1.Action = 4 'a_SelectMailbox
            If Not IMAP1.MessageCount = 0 Then
                IMAP1.MessageSet = "1:" & IMAP1.MessageCount
                'IMAP1.Action = a_GetMessageHeaders
                'Text1 = IMAP1.MessageHeaders
            End If
            ConnectIMAP4 = IMAP1.MessageCount
            MailConnector.ServerConnected = True
        Else
            MsgBox ("A IMAP4 mail server hasn't been specified.")
        End If
    'End If
    Exit Function

ConnectError:
    Select Case Err
        Case 20312 'Protocol Error
            '---------------------------------------------
            'Invalid Password.
            '---------------------------------------------
            MsgBox "A protocol error was discovered. Error returned was: " & Err.Description, vbCritical + vbApplicationModal, "IMAP Error"
        Case Else
            MsgBox "An error discovered.  Error returned was: " & Err.Description, vbApplicationModal + vbCritical, "IMAP Error"
    End Select
   ' MailConnector.AccountPassword = "ERROR"
    Err.Clear
End Function

Public Sub DisconnectPOP()
POP1.Action = a_Disconnect
POP1.WinsockLoaded = False
MailConnector.ServerState = 0
End Sub

Public Sub DisconnectIMAP4()
IMAP1.Action = 14 'a_CloseMailbox
IMAP1.Action = 3 'a_Disconnect
IMAP1.WinsockLoaded = False
MailConnector.ServerState = 0
End Sub

Public Sub GetIMAPMessage(MessageNumber As String)
Dim MsgSize As Long
Dim i As Integer

    On Error GoTo GetIMAPError
    '
    'Now get header info
    '
    IMAP1.MessageSet = MessageNumber
    IMAP1.Action = a_GetMessageInfo
    gMessageRecord.Subject = IMAP1.MessageSubject
    gMessageRecord.From = IMAP1.MessageFrom
    gMessageRecord.SentDate = IMAP1.MessageDate
    gMessageRecord.Received = IMAP1.MessageDeliveryTime
    gMessageRecord.ReplyTo = IMAP1.MessageReplyTo
    gMessageRecord.MessageID = IMAP1.MessageID
    i = 1
    gMessageRecord.To = IMAP1.MessageTo(0)
    Do While Not IMAP1.MessageTo(i) = ""
        gMessageRecord.To = gMessageRecord.To & "," & IMAP1.MessageTo(i)  ' Need to list all addressees...
        i = i + 1
    Loop
    gMessageRecord.CC = IMAP1.MessageCc(0)
    i = 1
    Do While Not IMAP1.MessageCc(i) = ""
        gMessageRecord.CC = gMessageRecord.CC & "," & IMAP1.MessageCc(i)   ' Need to list all addressees...
        i = i + 1
    Loop
    
    IMAP1.Action = a_GetMessageHeaders
    gMessageRecord.Header = IMAP1.MessageHeaders
    
    'Okay fire the start transfer event
    
    MsgSize = CLng((IMAP1.MessageSize) * 1.5)
    If AllocateMemory(MsgSize) Then
        IMAP1.Action = a_GetMessageText
        gMessage = String(MsgSize, Chr(0))
        Call agCopyData(ByVal gBuffer.StartAddr, ByVal gMessage, MsgSize)
        If Not FreeMemory Then
                Err.Raise 99999, , "Can't free memory..you should terminate the application and restart."
                Beep
        End If
    Else
        Err.Raise 99999, , "Can't allocate sufficient memory!"
    End If
    Exit Sub

GetIMAPError:
    Beep
    MsgBox "Error returned by IMAP Server: " & Err.Description & " (GetIMAPMessage)", vbApplicationModal + vbCritical, "POP Error"
    DisconnectIMAP4
    Err.Clear
End Sub

Public Sub DeletePOPMessage(MessageNumber As String)
POP1.MessageNumber = (Val(MessageNumber))
POP1.Action = a_Delete
End Sub

Public Sub DeleteIMAPMessage(MessageNumber As String)
IMAP1.MessageSet = (Val(MessageNumber))
IMAP1.MessageFlags = "\Deleted"
IMAP1.Action = 24 'a_SetMessageFlags
IMAP1.Action = 13 'a_ExpungeMailbox
End Sub

Public Sub ForwardMessage(iCurrentMessage As Integer)
If iCurrentMessage = 0 Then Exit Sub
PIForm(iCurrentMessage).WebBrowser1.Visible = False
PIForm(iCurrentMessage).MessageArea.Visible = True
PIForm(iCurrentMessage).EnableMessageFields
'PIForm(iCurrentMessage).txtSubject.Text = rs("Subject")

PIForm(iCurrentMessage).cmbRemailerSelect.Enabled = True
PIForm(iCurrentMessage).btnTo.Caption = "To"
PIForm(iCurrentMessage).txtCC.Text = ""
PIForm(iCurrentMessage).MessageArea.SelStart = 0
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = "------------"
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelBold = True
PIForm(iCurrentMessage).MessageArea.SelText = "From: " & PIForm(iCurrentMessage).lblFrom(1)
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = "Subject: " & PIForm(iCurrentMessage).txtsubject
PIForm(iCurrentMessage).MessageArea.SelBold = False
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelText = vbNewLine
PIForm(iCurrentMessage).MessageArea.SelStart = 0
PIForm(iCurrentMessage).txtTo.Text = ""
PIForm(iCurrentMessage).txtCC.Text = ""
PIForm(iCurrentMessage).AddressList.Initialise
End Sub

Public Sub ListIMAPMailBoxes()
On Error Resume Next
IMAP1.Mailbox = """*"""
IMAP1.Action = a_ListMailboxes
End Sub

Public Sub DisableToolBarButtons()
SSRibbon1.Item(1).Enabled = False
SSRibbon1.Item(2).Enabled = False
SSRibbon1.Item(5).Enabled = False
End Sub

Public Sub EnableToolBarButtons()
SSRibbon1.Item(1).Enabled = True
SSRibbon1.Item(2).Enabled = True
SSRibbon1.Item(5).Enabled = True
End Sub

Public Function CheckLicenceExpiredinDatabase() As Boolean
Dim rs As Recordset
On Error Resume Next
Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
If rs("Expired") = True Then CheckLicenceExpiredinDatabase = True
End Function

Private Sub ConfigureDisplayControls(iCurrentMessage As Integer)
PIForm(iCurrentMessage).SSRibbon1(0).Enabled = False ' don't allow them to send
PIForm(iCurrentMessage).SSRibbon1(6).Enabled = False ' don't allow them to send
PIForm(iCurrentMessage).SSRibbon1(9).Enabled = False ' don't allow them to send
PIForm(iCurrentMessage).SSRibbon1(4).Enabled = False ' don't allow them to send
PIForm(iCurrentMessage).SSRibbon1(7).Enabled = False ' don't allow them to send
PIForm(iCurrentMessage).SSRibbon1(10).Enabled = True ' allow reply to sendere
PIForm(iCurrentMessage).SSRibbon1(12).Enabled = True ' allow forware to sendere

    PIForm(iCurrentMessage).cmbRemailerSelect.Enabled = False
    
    PIForm(iCurrentMessage).txtToAddresses.BackColor = PIForm(iCurrentMessage).BackColor
    PIForm(iCurrentMessage).txtCCAddresses.BackColor = PIForm(iCurrentMessage).BackColor

    PIForm(iCurrentMessage).btnTo.Visible = False
    PIForm(iCurrentMessage).txtTo.Visible = False
    PIForm(iCurrentMessage).txtToAddresses.Visible = True
    PIForm(iCurrentMessage).lblTo.Visible = True
    
    PIForm(iCurrentMessage).lblTo.Left = PIForm(iCurrentMessage).btnTo.Left
    PIForm(iCurrentMessage).txtToAddresses = PIForm(iCurrentMessage).txtTo.Text
    PIForm(iCurrentMessage).txtToAddresses.Left = PIForm(iCurrentMessage).txtTo.Left
    PIForm(iCurrentMessage).txtToAddresses.Width = PIForm(iCurrentMessage).txtTo.Width
    

    PIForm(iCurrentMessage).btnCC.Visible = False
    PIForm(iCurrentMessage).txtCC.Visible = False
    PIForm(iCurrentMessage).lblcc.Visible = True
    PIForm(iCurrentMessage).txtCCAddresses.Visible = True
   
    
    PIForm(iCurrentMessage).lblcc.Left = PIForm(iCurrentMessage).btnCC.Left
    PIForm(iCurrentMessage).txtCCAddresses = PIForm(iCurrentMessage).txtCC.Text
    PIForm(iCurrentMessage).txtCCAddresses.Left = PIForm(iCurrentMessage).txtCC.Left
   ' PIForm(iCurrentMessage).lblCC.Height = PIForm(iCurrentMessage).txtCC.Height
    PIForm(iCurrentMessage).txtCCAddresses.Width = PIForm(iCurrentMessage).txtCC.Width
   ' PIForm(iCurrentMessage).lblCC.Top = PIForm(iCurrentMessage).txtCC.Top
    
    PIForm(iCurrentMessage).lblSubject = PIForm(iCurrentMessage).txtsubject
    PIForm(iCurrentMessage).lblSubject.Left = PIForm(iCurrentMessage).txtsubject.Left
    'PIForm(iCurrentMessage).lblSubject.Height = PIForm(iCurrentMessage).txtsubject.Height
    PIForm(iCurrentMessage).lblSubject.Width = PIForm(iCurrentMessage).txtsubject.Width
    'PIForm(iCurrentMessage).lblSubject.Top = PIForm(iCurrentMessage).txtsubject.Top
    PIForm(iCurrentMessage).lblSubject.Visible = True
    PIForm(iCurrentMessage).txtsubject.Visible = False
End Sub



Public Function StripItemCount(sName As String) As String
Dim i As Integer

If InStrRev(sName, "(") = 0 Then
    StripItemCount = sName
Else
    i = InStr(Len(sName) - 5, sName, "(")
    If i = 0 Then
        StripItemCount = sName
        Exit Function
    End If
    StripItemCount = Mid(sName, 1, i - 2)
End If
End Function

Private Sub AddFolderParameters(TreeIndex)
Dim rsItems As Recordset
Dim rsFolder As Recordset
Dim qd As QueryDef
Dim qdItems As QueryDef
Dim lFolderIndex As Long
Dim n As SSNode

  
'
'First Find the node name
'
    On Error GoTo Exitsub:
    lFolderIndex = SSTree1(TreeIndex).SelectedItem.Index
    Set n = SSTree1(TreeIndex).Nodes.Item(lFolderIndex)
'n.Text = "fdsfasdffsd24242dssfs"

'
'Now find the folder ID in the database
'
  On Error GoTo ErrorAddingParameter
    If TreeIndex = 0 Then
        Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
        rsFolder.FindFirst "[Folder] =" & "'" & StripItemCount(n.Text) & "'"
    Else
        Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
        rsFolder.FindFirst "[Folder] =" & "'" & StripItemCount(n.Text) & "'"
    End If
    If rsFolder.NoMatch Then
        Exit Sub
    End If
    'If TreeIndex = 0 Then
   '     Set qd = DB.QueryDefs("qdSubFolders")
   ' Else
   '     Set qd = DB.QueryDefs("qdFileSubFolders")
   ' End If
   ' qd.Parameters![NodeID] = Tree(TreeIndex).NodeID 'rs("Node ID")
   ' Set rsFolder = qd.OpenRecordset()
'
'We have the Folder ID now
'
   ' sFolderName = rsFolder("Folder")
   ' MsgBox "Debug 1", vbApplicationModal + vbCritical
    If TreeIndex = 0 Then
     ' MsgBox "Debug 2", vbApplicationModal + vbCritical
        Set qdItems = DB.QueryDefs("qdNumberofMessagesinFolder")
      '    MsgBox "Debug 3", vbApplicationModal + vbCritical
        qdItems.Parameters![FolderId] = rsFolder("Folder ID")
       '   MsgBox "Debug 4", vbApplicationModal + vbCritical
        Set rsItems = qdItems.OpenRecordset()
       '   MsgBox "Debug 5", vbApplicationModal + vbCritical
        'If Not rsItems.EOF Then rsItems.MoveFirst
        n.Text = StripItemCount(n.Text) & " (" & IIf(rsItems.EOF, 0, rsItems("Number of Messages")) & ")"
       '   MsgBox "Debug 6", vbApplicationModal + vbCritical
        'If Not rsItems.RecordCount = 0 Then sFolderName = sFolderName & " (" & rsItems("[Number of Messages]") & ")"
    Else
    '  MsgBox "Debug 7", vbApplicationModal + vbCritical
        Set qdItems = DB.QueryDefs("qdNumberofFilesinFolder")
        '  MsgBox "Debug 8", vbApplicationModal + vbCritical
        '    MsgBox "Debug 9", vbApplicationModal + vbCritical
        qdItems.Parameters![FolderId] = rsFolder("Folder ID")
        '  MsgBox "Debug 10", vbApplicationModal + vbCritical
        Set rsItems = qdItems.OpenRecordset()
      '    MsgBox "Debug 11", vbApplicationModal + vbCritical
       ' If Not rsItems.EOF Then rsItems.MoveFirst
        n.Text = StripItemCount(n.Text) & " (" & IIf(rsItems.EOF, 0, rsItems("Number of Files")) & ")"
      '    MsgBox "Debug 12", vbApplicationModal + vbCritical
        'If rsItems.EOF Then Exit Sub
        'If Not rsItems.RecordCount = 0 Then sFolderName = sFolderName & " (" & rsItems("[Number of Files]") & ")"
    End If
    'Set n = SSTree1(Index).Nodes.Add(sNodeName, ssatChild, , sFolderName, "closed", "open")
         
    'Set n = Nothing

'i = SSTree1(Index).SelectedItem.Index
'Set n = SSTree1(Index).Nodes.Item(i) '..Add(, , , "Private Idaho File Safe", "EncryptedEnvelope", "EncryptedEnvelope")
'If Not rsItems.EOF Then rsItems.MoveFirst
'n.Text = StripItemCount(n.Text) & " (" & rsItems("Number of Files") & ")"
'n.Font.Bold = True
Set n = Nothing
Exitsub:
Err.Clear
Exit Sub

ErrorAddingParameter:
    MsgBox "Error adding folder parameter: " & Err.Description, vbApplicationModal + vbCritical, "Add Folder Parameters"
    Err.Clear
End Sub

Private Function GetSelectedFolderID(Index As Integer) As Long
Dim rs As Recordset
Dim lFolderIndex As Long
Dim n As SSNode

On Error Resume Next
'
'First Find the node name
'
    lFolderIndex = SSTree1(Index).SelectedItem.Index
    Set n = SSTree1(Index).Nodes.Item(lFolderIndex)

If Index = 0 Then
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
Else
    Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
End If
rs.FindFirst "Folder =" & "'" & StripItemCount(n.Text) & "'"
If rs.NoMatch Then
   GetSelectedFolderID = 0
   ' If Index = 0 Then
  '      Set rs = DB.OpenRecordset("Nodes", dbOpenDynaset)
  '  Else
   '     Set rs = DB.OpenRecordset("File Nodes", dbOpenDynaset)
  '  End If
  '  rs.FindFirst "[Node Name] =" & "'" & Node & "'"
  '  If Not rs.NoMatch Then
  '      Tree(Index).NodeName = rs("Node Name")
  '  End If
  '  Tree(Index).FolderId = 0
Else
    GetSelectedFolderID = rs("Folder ID")
End If

rs.Close
Set rs = Nothing
End Function

'Create instance and return index

Private Function CreateNewPIInstance() As Integer
gFormInstance = gFormInstance + 1
If gFormInstance > MAX_NUMBER_OF_PI_INSTANCES Then
    CreateNewPIInstance = 0 'Error
    Exit Function
End If
Set PIForm(gFormInstance) = New frmPI
PIForm(gFormInstance).Show
CreateNewPIInstance = gFormInstance
End Function
