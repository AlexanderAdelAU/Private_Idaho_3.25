VERSION 5.00
Object = "{33333243-F789-11CE-86F8-0020AFD8C6DB}#3.0#0"; "POP33.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMain1 
   Caption         =   "Private Idaho Version 3.2"
   ClientHeight    =   6195
   ClientLeft      =   2835
   ClientTop       =   4215
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6195
   ScaleWidth      =   11355
   Begin POPLib.POP POP1 
      Left            =   2970
      OleObjectBlob   =   "main1.frx":0000
      Top             =   5700
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   30000
      Left            =   6780
      Top             =   5730
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   6300
      Top             =   5730
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5115
      Left            =   300
      TabIndex        =   1
      Top             =   540
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   9022
      _Version        =   131073
      ClipControls    =   -1  'True
      PaneTree        =   "main1.frx":002C
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "main1.frx":007E
         Height          =   5055
         Left            =   2640
         TabIndex        =   2
         Top             =   30
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   8916
         _Version        =   65541
         Rows            =   7
         Cols            =   6
         SelectionMode   =   1
         AllowUserResizing=   1
         OLEDropMode     =   1
      End
      Begin SSActiveTreeView.SSTree SSTree1 
         Height          =   5055
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   8916
         _Version        =   65536
         ImagesMaskColor =   16777215
         LineStyle       =   1
         IndentationStyle=   1
         Indentation     =   315
         ImageCount      =   3
         ImageListIndex  =   0
         AllowDelete     =   -1  'True
         HideSelection   =   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         TabStops        =   "32,63,96,128,200"
         ImageList       =   "ImageList1"
         Image(0).Index  =   0
         Image(0).Picture=   "main1.frx":0093
         Image(0).Key    =   "closed"
         Image(1).Index  =   1
         Image(1).Picture=   "main1.frx":03E5
         Image(1).Key    =   "open"
         Image(2).Index  =   2
         Image(2).Picture=   "main1.frx":053F
         Image(2).Key    =   "leaf"
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
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   767
      _Version        =   131073
      PictureBackgroundStyle=   2
      PictureBackground=   "main1.frx":0699
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131073
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main1.frx":10FC3
         ButtonStyle     =   3
      End
      Begin VB.Image ImgList 
         Height          =   480
         Index           =   5
         Left            =   6150
         Picture         =   "main1.frx":110FD
         Top             =   -30
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgList 
         Height          =   480
         Index           =   4
         Left            =   6600
         Picture         =   "main1.frx":11407
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   1
         Left            =   570
         TabIndex        =   4
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131073
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "main1.frx":11711
         ButtonStyle     =   3
      End
      Begin ComctlLib.ImageList ImageList1 
         Index           =   0
         Left            =   8970
         Top             =   -180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   27
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":1184B
               Key             =   "closed"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":11B9D
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":11EEF
               Key             =   "leaf"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":12241
               Key             =   "happy"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":1255B
               Key             =   "apathy"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":12875
               Key             =   "USA"
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":12B8F
               Key             =   "Brazil"
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":12EA9
               Key             =   "Canada"
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":131C3
               Key             =   "France"
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":134DD
               Key             =   "Italy"
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":137F7
               Key             =   "Japan"
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":13B11
               Key             =   "Spain"
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":13E2B
               Key             =   "Sweden"
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":14145
               Key             =   "Switzerland"
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":1445F
               Key             =   "UnitedKingdom"
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":14779
               Key             =   "Austria"
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":14A93
               Key             =   "bulbon"
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":14DAD
               Key             =   "bulboff"
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":150C7
               Key             =   "question"
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":153E1
               Key             =   "xmas"
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":156FB
               Key             =   "openlock"
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":15A15
               Key             =   "closedlock"
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":15D2F
               Key             =   "rad"
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":16049
               Key             =   "exclamation"
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":16363
               Key             =   "BrokenEnvelope"
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":1667D
               Key             =   "DragIcon"
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "main1.frx":16997
               Key             =   "DropIcon"
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgList 
         Height          =   195
         Index           =   3
         Left            =   8490
         Picture         =   "main1.frx":16CB1
         Top             =   90
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   0
         Left            =   7110
         Picture         =   "main1.frx":16D9B
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   1
         Left            =   7470
         Picture         =   "main1.frx":172A9
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgList 
         Height          =   240
         Index           =   2
         Left            =   7890
         Picture         =   "main1.frx":177FD
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
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
      Begin VB.Menu mFile_NewFolder 
         Caption         =   "New Folder"
      End
      Begin VB.Menu Filebreak2 
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
   End
   Begin VB.Menu mTools 
      Caption         =   "Tools"
      Begin VB.Menu mTools_EmailScan 
         Caption         =   "Scan for PGP Messages"
      End
      Begin VB.Menu mFile_MessageScanInterval 
         Caption         =   "Message Scan Interval"
      End
   End
End
Attribute VB_Name = "frmMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Folders(100) As String

Private Const CLOSED_ENVELOPE As Integer = 0
Private Const OPEN_ENVELOPE As Integer = 1
Private Const ATTACHMENT As Integer = 2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Const TIMER_DRAG_DROP As Integer = 0
Private Const TIMER_MESSAGE_SCAN As Integer = 1
Public sFolderName As String


'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub BuildTree()
    Dim i As Integer
    Dim FolderCount As Integer
    Dim Key As String
    Dim start As Long
    Dim n As SSNode
    Dim rs As Recordset
    'Root folders
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    FolderCount = 0
    Do While Not rs.EOF
        Folders(FolderCount) = rs("Folder")
        rs.MoveNext
        FolderCount = FolderCount + 1
    Loop
    rs.Close
    Set rs = Nothing
    'Show Picture property
    SSTree1.Nodes.Clear
    Set n = SSTree1.Nodes.Add(, , , Folders(0), "BrokenEnvelope")
    n.Font.bold = True
    Set n = Nothing
    SSTree1.Nodes.Add 1, , "Pic", "Mail Folders", 1, 2
           
    For i = 1 To FolderCount - 1
        Set n = SSTree1.Nodes.Add("Pic", ssatChild, , Folders(i), "closed", "open")
       ' n.Picture = "rad"
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
        Set n = Nothing
    Next
    GoTo temp
        Set n = SSTree1.Nodes.Add("Pic", ssatChild, , Folders(4), "open", "closed")
       ' n.Picture = "rad"
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
        Set n = Nothing
        
        Set n = SSTree1.Nodes.Add("Pic", ssatChild, , "Deleted Items", "open", "closed")
        'n.Picture = "rad"
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
        Set n = Nothing
      
    SSTree1.Nodes.Add , , "Tips", "ActiveTreeView supports three kinds of ToolTips."
    SSTree1.Nodes.Add "Tips", ssatChild, "NodeTips", "NodeTips:"
        Set n = SSTree1.Nodes.Add("NodeTips", ssatChild, , "If a node is only partially visible then hold the mouse cursor over it, and you'll see a NodeTip pop up.  The Delay is dictated by the NodeTipDelay property.")
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
        Set n = Nothing
     SSTree1.Nodes.Add "Tips", ssatChild, "LineTips", "LineTips:"
        Set n = SSTree1.Nodes.Add("LineTips", ssatChild, , "If a node's is not visible but its treeline is then hold the mouse cursor over the line and the parent's name will pop up.")
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
        Set n = Nothing
    SSTree1.Nodes.Add "Tips", ssatChild, "ScrollTips", "ScrollTips:"
        Set n = SSTree1.Nodes.Add("ScrollTips", ssatChild, , "Set the ScrollStyle property to 'Deferred with ScrollTips' and then scroll vertically.  Tips will pop up to tell you which node would be at the tip and disappear when scrolling is complete.")
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
        Set n = Nothing

    'Scrolling
    SSTree1.Nodes.Add , , "scrolling", "Enhanced Scrolling options..."
    
    'Tabs
    SSTree1.Nodes.Add , , "tabs", "Tab character support for creating columns..."
    SSTree1.Nodes.Add , , "linetype", "The LineType property allows you to choose Solid or Dotted treelines."
    SSTree1.Nodes.Add , , "sel", "Multiple selections:  Toggle and Range selections."
   
    'Picture, IndentStyle
    
    'countries - suitable to demonstrate sorting
    SSTree1.Nodes.Add , , "sorting", "Try sorting these...", "closedlock", "openlock"
    
    For i = 6 To 16    'these are the images that hold countries' flags
        Key = ImageList1(0).ListImages(i).Key
        Set n = SSTree1.Nodes.Add("sorting", ssatChild, , Key, Key)
        n.LoadStyleChildren = 3     'no children (so plus sign does not appear)
    Next
temp:
End Sub




Private Sub Form_DragDrop(source As Control, x As Single, y As Single)
MousePointer = vbDefault
End Sub

Private Sub Form_DragOver(source As Control, x As Single, y As Single, State As Integer)
MousePointer = vbNoDrop
DoEvents
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Headings As String
Dim RowIndex As Integer
Dim ItemString As String
Dim Sqlq As String


On Error GoTo BadLoad
InitialiseGrid
BuildTree
Exit Sub
BadLoad:
    MsgBox Err.Description & " Can't load properly...", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMain = Nothing
End Sub

Private Sub mCompose_NewMessage_Click()
frmMain.Show
End Sub

Private Sub mFile_Delete_Click()
DeleteMessage
InitialiseGrid
FillGrid
End Sub

Private Sub mFile_Exit_Click()
Unload Me
End Sub

Private Sub mFile_MessageScanInterval_Click()
'Load frmMessageScanTimer
frmMessageScanTimer.ScanInterval = giEmailScanInterval
frmMessageScanTimer.Show vbModal
giEmailScanInterval = frmMessageScanTimer.ScanInterval
giTimerCounter = giEmailScanInterval
If giEmailScanInterval > 0 Then Timer1(TIMER_MESSAGE_SCAN).Enabled = True
Set frmMessageScanTimer = Nothing
End Sub

Private Sub mFile_NewFolder_Click()
Dim rs As Recordset
frmFolderName.Caption = "New Folder"
frmFolderName.Show vbModal
If sFolderName = "" Then
    Exit Sub
Else
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.AddNew
    rs("Folder") = sFolderName
    rs.Update
    BuildTree
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub mFile_NewMessage_Click()
End Sub

Private Sub mFile_Open_Click()
DisplayMessage
End Sub

Private Sub MSFlexGrid1_DblClick()
DisplayMessage
End Sub

Private Sub MSFlexGrid1_DragDrop(source As Control, x As Single, y As Single)
MSFlexGrid1.Drag vbEndDrag
MousePointer = vbDefault
End Sub


Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then DeleteMessage
'InitialiseGrid
'FillGrid
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Timer1(TIMER_DRAG_DROP).Enabled = True
End If
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1(TIMER_DRAG_DROP).Enabled = False
MSFlexGrid1.DragIcon = Nothing
MSFlexGrid1.Drag vbEndDrag
End Sub

Private Sub mTools_EmailScan_Click()
Dim i As Integer
    On Error GoTo ScanError
    MousePointer = vbHourglass
    frmScanStatus.Show
    DoEvents
    If CheckConnection Then
       ' lblStatus = "Connecting to mail server for scan...."
        gServerState = POPSTATE
        If POP1.WinsockLoaded = False Then
            POP1.WinsockLoaded = True
        End If
        gPOPPass = ReadProfile("Options", "POPPassword")
        ConnectPOP
        If Not gPOPPass = "HSLAMXYDELKSHALP" Then
           'password okay
            DoEvents
            gNumMessages = POP1.MessageCount
            If Not gNumMessages = 0 Then
                    gFoundMessages = ScanTopHeaders("-----BEGIN PGP MESSAGE-----")
                    If gFoundMessages <> "" Then
                        ParseFoundMessages (gFoundMessages)
                        DoEvents
                        Beep
                    End If
            End If
            '---------------------------------------------
            'need this to enter the update state
            'we be done, so...
            '---------------------------------------------
            POP1.Action = 2
            POP1.WinsockLoaded = False
            gServerState = 0
        Else
           'password NOT okay
            gPOPPass = ""   'This will reset so PI can prompt again
            POP1.Action = 2
            POP1.WinsockLoaded = False
            
        End If
    End If
    'lblStatus = ""
    MousePointer = vbDefault
    Unload frmScanStatus
    Exit Sub
ScanError:
    MousePointer = vbDefault
    Unload frmScanStatus
    'lblStatus = ""
    Select Case Err.Number
        Case 20113
            '---------------------------------------------
            'Last action tied up network, so disconnect
            '---------------------------------------------
            POP1.Action = 2
            POP1.WinsockLoaded = False
            Err.Clear
            Exit Sub
        Case 25066
            '---------------------------------------------
            '[10065] No route to host.
            '---------------------------------------------
            
        Case Else
    End Select
    gServerState = 0
    Err.Clear
End Sub

Private Sub POP1_EndTransfer()
DoEvents
End Sub

Private Sub POP1_Error(ErrorCode As Integer, Description As String)
'lblStatus = "error occured" & ErrorCode & Description
Beep
End Sub

Private Sub POP1_Header(Field As String, Value As String)
 gMessageRecord.Header = gMessageRecord.Header & Field & ": " & Value & vbCrLf

End Sub

Private Sub POP1_PITrail(Direction As Integer, Message As String)
On Error GoTo POPerror1
    If gSMTPLog = 1 Then
        Print #gSMTPFile, "POP: " + Format$(Direction) + ": " + Message
    End If
    Exit Sub
POPerror1:
    MsgBox Err.Description & " Connection lost...", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

Private Sub POP1_Transfer(BytesTransferred As Long, Text As String)
Dim sText As String * 2048
       
        frmScanStatus.lblBytesTransferred = BytesTransferred
        If gBuffer.hMem = 0 Then
            gMessage = gMessage & Text & vbCrLf
        Else
            If Len(Text) > 2046 Then
                'lblStatus = "Buffer overflow error..."
                Beep
                Exit Sub
            End If
            If Len(gMessageRecord.Header) <> 0 Then
                gMessageRecord.Header = gMessageRecord.Header + Text
                Text = gMessageRecord.Header
                gMessageRecord.Header = ""
            End If
            Text = Text & vbCrLf
            sText = Text
            Call agCopyData(ByVal sText, ByVal gBuffer.Address, CLng(Len(Text)))
            gBuffer.Address = gBuffer.Address + Len(Text)
        End If
        DoEvents
End Sub

Private Sub SSPanel1_DragOver(source As Control, x As Single, y As Single, State As Integer)
MousePointer = vbNoDrop
End Sub

Private Sub SSPanel1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1(TIMER_DRAG_DROP).Enabled = False
MousePointer = vbDefault
End Sub

Private Sub SSRibbon1_Click(Index As Integer, Value As Integer)
' To simulate a pushbutton just reset the value to false
' and it will pop up
If Value = True Then
    Select Case Index
        Case 0
            frmPI.Show
        Case 1
            DeleteMessage
    End Select
End If
SSRibbon1(Index).Value = False
End Sub

Private Sub SSSplitter1_DragDrop(source As Control, x As Single, y As Single)
MSFlexGrid1.Drag vbEndDrag
MousePointer = vbDefault
End Sub

Private Sub SSSplitter1_DragOver(source As Control, x As Single, y As Single, State As Integer)
MousePointer = vbNoDrop
End Sub

Private Sub SSTree1_AfterLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)
Dim rs As Recordset
Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
If NewString = "" Then Exit Sub
rs.AddNew
rs("Folder") = NewString
rs.Update
rs.Close
Set rs = Nothing
BuildTree
End Sub

Private Sub SSTree1_AfterNodeDelete(Node As SSActiveTreeView.SSNode)
Dim rs As Recordset
Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
rs.FindFirst "Folder =" & "'" & Node & "'"
If rs.NoMatch Then
    Beep
    Exit Sub
Else
   'iresponse = MsgBox("You are about to delete these messages.  Are you sure?", vbYesNo + vbQuestion + vbApplicationModal, "Are you sure?")
   'If iresponse = vbNo Then Exit Sub
   rs.Delete
End If
rs.Close
Set rs = Nothing
BuildTree
End Sub

Private Sub SSTree1_BeforeNodeDelete(Node As SSActiveTreeView.SSNode, Cancel As SSActiveTreeView.SSReturnBoolean, DispPromptMsg As SSActiveTreeView.SSReturnBoolean)
Dim rs As Recordset
Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
rs.FindFirst "Folder =" & "'" & Node & "'"
If rs.NoMatch Then
    Beep
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
   If Not rs("Can Delete") Then
    Cancel = True
    MsgBox "Can't delete system folder", vbQuestion + vbApplicationModal, "System Folder"
   End If
End If
rs.Close
Set rs = Nothing
'BuildTree
End Sub

Private Sub SSTree1_DragDrop(source As Control, x As Single, y As Single)
Dim ssNodeTmp As SSNode
Dim ssNodeX As SSNode
Dim Sqlq As String
Dim i As Integer
Dim Heading As String
Dim NewFolderID As Long
Timer1(TIMER_DRAG_DROP).Enabled = False
MSFlexGrid1.Drag vbEndDrag
MousePointer = vbDefault 'vbNoDrop
Set ssNodeTmp = SSTree1.HitTest(x, y)
If ssNodeTmp Is Nothing Then Exit Sub
Set SSTree1.SelectedItem = ssNodeTmp
If source = MSFlexGrid1 Then
    'SSTree1.DragIcon = ImgList(5).Picture 'LoadPicture("d:\pi32\icons\drop1pg.ico") 'ImageList1(0) 'drop icon
    'gSelectedFolder = Node
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "Folder =" & "'" & ssNodeTmp & "'"
    If rs.NoMatch Then
        Beep
        Exit Sub
    Else
        NewFolderID = rs("Folder ID")
    End If
    
'Update
    rs.Close
    Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
    Sqlq = Sqlq & "(Messages.[Folder ID] = " & gFolderID & ")"
    Sqlq = Sqlq & "ORDER BY Messages.Date;"
    Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)

    If rs.EOF Then
        Beep
        Exit Sub
    Else
        rs.MoveFirst
    End If
    For i = 1 To MSFlexGrid1.Row - 1
        rs.MoveNext
    Next
    rs.Edit
    rs("Folder ID") = NewFolderID
    rs.Update
'Update FlexGrid
    DoEvents
    If MSFlexGrid1.Rows > 2 Then
        MSFlexGrid1.RemoveItem i
    Else
        MSFlexGrid1.Clear
        Heading = "    |^Read|^Att|<To/From                                 |<Subject                                  |<Date                          "
        MSFlexGrid1.FormatString = Heading
    End If
    rs.Close
    Set rs = Nothing
End If
  
End Sub

Private Sub SSTree1_DragOver(source As Control, x As Single, y As Single, State As Integer)
Dim ssNodeTmp As SSNode
'Dim ssNodeX As SSNode
   Set ssNodeTmp = SSTree1.HitTest(x, y)
    If Not ssNodeTmp Is Nothing Then
        Set SSTree1.SelectedItem = ssNodeTmp
        'ssNodeTmp.Font.bold = True
    End If

   ' Set ssNodeX = SSTree1.SelectedItem 'Set the item being dragged

End Sub
Private Sub SSTree1_NodeClick(Node As SSActiveTreeView.SSNode)
Dim rs As Recordset

'gSelectedFolder = Node
Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
rs.FindFirst "Folder =" & "'" & Node & "'"
If rs.NoMatch Then
    InitialiseGrid
    gFolderID = 0
    Beep
    rs.Close
    Set rs = Nothing
    Exit Sub
End If
gFolderID = rs("Folder ID")
rs.Close
Set rs = Nothing
InitialiseGrid
FillGrid
End Sub

Private Sub SSTree1_OLEDragOver(Data As SSActiveTreeView.SSDataObject, Effect As SSActiveTreeView.SSReturnLong, Button As Integer, Shift As Integer, x As Single, y As Single, State As SSActiveTreeView.SSReturnShort)
'Set SSTree1.DropHighlight = SSTree1.HitTest(x, y)

    'If Not SSTree1.DropHighlight Is Nothing Then
     '   If SSTree1.DropHighlight.Level > 1 Then
       '     Effect = ssOLEDropEffectNone
        '    Exit Sub
       ' End If

       ' Effect = ssOLEDropEffectCopy
    'End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
If Index = 0 Then
    Timer1(TIMER_DRAG_DROP).Enabled = False
    MSFlexGrid1.DragIcon = ImgList(4).Picture
    MSFlexGrid1.Drag vbBeginDrag
Else
    giTimerCounter = giTimerCounter - 1
    If giTimerCounter = 0 Then
        Timer1(TIMER_MESSAGE_SCAN).Enabled = False
        mTools_EmailScan_Click
        giTimerCounter = giEmailScanInterval
        Timer1(TIMER_MESSAGE_SCAN).Enabled = True
    End If
End If
End Sub

Public Function GetSelectedFolder() As Long

End Function

Public Sub InitialiseGrid()
Dim Headings As String
MSFlexGrid1.Clear
Headings = "    |^Read|^Att|<From                                     |<Subject                                  |<Date Sent/Received               "
MSFlexGrid1.Cols = 6
MSFlexGrid1.Rows = 1
MSFlexGrid1.FormatString = Headings
End Sub

Public Sub ConnectPOP()
    On Error GoTo ConnectError
    
   ' gPOPPass = ""
    If gPOPSeverName <> "" Then
        POP1.MailServer = gPOPSeverName
        POP1.User = gPOPName
        POP1.Password = gPOPPass
        POP1.Action = 1
    Else
        MsgBox ("A POP mail server hasn't been specified.")
    End If
    Exit Sub

ConnectError:
    Select Case Err
        Case 20172
            '---------------------------------------------
            'Invalid Password.
            '---------------------------------------------
            MsgBox "Password you specified is not valid."
            gPOPPass = "HSLAMXYDELKSHALP"
            'gPOPPass = ""
            Exit Sub
        Case 25066
            MsgBox "This is an ISP connection error.  Apparently, you have become disconnected from your Internet Service Provider.  Please connect to your ISP and retry."
        
        Case 26005
             End Select
    Err.Clear
End Sub


Sub GetPOPMessage(msgNumber As Integer)

    On Error GoTo GetPOPError
    gMessage = ""
    gMessageRecord.Header = ""
    'If AllocateMemory(frmMain.POP1.MessageSize) Then
        POP1.MaxLines = 0
        POP1.MessageNumber = msgNumber
        POP1.Action = 3
    'Else
      '  Err.Raise 99999, , "Can't allocate sufficient memory!"
    'End If
    Exit Sub

GetPOPError:
    MsgBox Err.Description & " (GetPOPMessage)"
    POP1.Action = 0
    Err.Clear
End Sub

Public Sub GetPOPTop(MsgNum As Integer)
   
    On Error GoTo POPTopError
    gMessage = ""
    gMessageRecord.EndTransfer = False
    POP1.MaxLines = 30
    POP1.MessageNumber = MsgNum
    POP1.Action = 3
Exit Sub

POPTopError:
    Select Case Err
    Case 20172
            '---------------------------------------------
            'Invalid Password.
            '---------------------------------------------
            MsgBox "Password you specified is not valid."
    Case 25058
            '---------------------------------------------
            'Lost the socket.
            '---------------------------------------------
            MsgBox "The connection was lost."
    Case Else
        MsgBox Err.Description + " (GetPOPTop)"
    End Select
    POP1.Action = 0
    Err.Clear
End Sub
Function ScanTopHeaders(search As String) As String
       
    'search - a string to search for within the header
    'ScanTopHeaders - returns a string containing found message numbers delimited
    'by spaces
    Dim NumMessages As Integer
    Dim tmpstr As String
    Dim j As Integer
    
    NumMessages = 0
    tmpstr = ""
    For j = 1 To gNumMessages
        DoEvents
        GetPOPTop (j)
        DoEvents
        gMessage = Mid$(gMessage, 1, Len(gMessage))
        If InStr(1, gMessage, search) Then
            NumMessages = NumMessages + 1
            tmpstr = tmpstr & Format(j) & " "
        End If
    Next
    ScanTopHeaders = RTrim$(tmpstr)
End Function

Sub ParseFoundMessages(TheMessages As String)
       
    Dim TempMessages As String
    Dim CurrentMessage As String
    Dim SectionName As String
    Dim res As Long
    Dim i As Integer
    Dim izap As Integer
    Dim tmpstr As String
        
    MousePointer = vbHourglass
    DoEvents
    If TheMessages <> "" Then
        i = 1
        TempMessages = TheMessages
        On Error GoTo POPParseError
        'see if messages should be deleted from the server after download
        SectionName = "Options"
        tmpstr = ReadProfile(SectionName, "ServerDelete")
        If tmpstr = "1" Or tmpstr = "" Then
            izap = 1
        Else
            izap = 0
        End If
        'parse through the string, extracting each message number
        Do Until i = 0
            i = InStr(1, TempMessages, " ")
            If i > 1 Then
                CurrentMessage = Mid$(TempMessages, 1, i - 1)
            Else
                CurrentMessage = TempMessages
            End If
            frmScanStatus.lblStatus = CurrentMessage
            DoEvents
            GetPOPMessage (CurrentMessage)
            
            'write it to database
            WriteMessageRecord (gMessage)
            If Not gBuffer.hMem = 0 Then
                If Not FreeMemory Then
                    MsgBox "Can't free memory..stopping", vbApplicationModal + vbError, App.Title
                    Beep
                    Exit Sub
                End If
            End If
    
  'now delete it from the server - if option is set
           
           izap = 1
            If izap = 1 Then
                POP1.MessageNumber = (val(CurrentMessage))
                POP1.Action = 4
            End If
            'bump the string
            TempMessages = Mid$(TempMessages, i + 1, Len(TempMessages) - i)
        Loop
        'there's some messages, or we would never had called
        'this routine, so display the list box
        'giMailbox = SHOW_IN_MAILBOX
        Beep   'don't show  frmMailBox.Show just beep to let them know
    End If
    MousePointer = vbDefault
    Exit Sub

POPParseError:
    MousePointer = vbDefault
    MsgBox Err.Description & " (ParseFoundMessages)"
    Err.Clear
End Sub

Sub POPLogin()
       
    If gPOPPass = "" Then
        frmPOPPassword.Show vbModal
    End If
End Sub

Private Sub DeleteMessage()
Dim i As Integer
Dim Heading As String
Dim iResponse As Integer
Static InHere As Boolean
Dim Sqlq As String
Dim rs As Recordset

If InHere Then Exit Sub

InHere = True
On Error GoTo BadKeyEntry:

Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
Sqlq = Sqlq & "(Messages.[Folder ID] = " & gFolderID & ")"
Sqlq = Sqlq & "ORDER BY Messages.Date;"
Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)

If rs.EOF Then
    Beep
    InHere = False
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.MoveFirst
End If
For i = 1 To MSFlexGrid1.Row - 1
    rs.MoveNext
Next

If gFolderID = gDeletedFolderID Then
    iResponse = MsgBox("Do you wish to permanently delete this message?", vbYesNo + vbQuestion + vbApplicationModal, "Delete message")
    If iResponse = vbYes Then
        Kill App.Path & "\mailbox\" & rs("Message File Name")
        rs.Delete
        MSFlexGrid1.RemoveItem MSFlexGrid1.Row
    End If
Else
    MSFlexGrid1.RemoveItem MSFlexGrid1.Row
    rs.Edit
    rs("Folder ID") = gDeletedFolderID
    rs.Update
End If

InHere = False
rs.Close
Set rs = Nothing
Exit Sub

BadKeyEntry:
    MsgBox Err.Description & " - Can't complete action you requested"
    Beep
    Err.Clear
    rs.Close
    Set rs = Nothing
    InHere = False
End Sub

Private Sub FillGrid()
Dim rs As Recordset
Dim Headings As String
Dim Sqlq As String
Dim RowIndex As Integer
Dim ItemString As String

Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
Sqlq = Sqlq & "(Messages.[Folder ID] = " & gFolderID & ")"
Sqlq = Sqlq & "ORDER BY Messages.Date;"
Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Exit Sub
End If
'Now fill the message area
RowIndex = 1

While Not rs.EOF
    MSFlexGrid1.Col = 1
    ItemString = "" & vbTab & "" & vbTab & "" & vbTab & rs("From/To") & vbTab & rs("Subject") & vbTab & rs("Date")
    MSFlexGrid1.AddItem ItemString, RowIndex
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = RowIndex
    If rs("Message Read") Then
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(OPEN_ENVELOPE).Picture
    Else
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(CLOSED_ENVELOPE).Picture
    End If
    MSFlexGrid1.Col = 2
    If rs("Attachment") Then
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(ATTACHMENT).Picture
    End If
    RowIndex = RowIndex + 1
    rs.MoveNext
Wend
MSFlexGrid1.FormatString = Headings
rs.Close
Set rs = Nothing
MSFlexGrid1.Row = 1
'MSFlexGrid1.Col = 1
MSFlexGrid1.RowSel = 1


End Sub

Private Sub DisplayMessage()
Dim FileName As String
Dim FileNum As Integer
Dim StartSection As Integer
Dim i As Integer
Dim Sqlq As String
Dim rs As Recordset

On Error GoTo BadFileFind
Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
Sqlq = Sqlq & "(Messages.[Folder ID] = " & gFolderID & ")"
Sqlq = Sqlq & "ORDER BY Messages.Date;"
Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)

If rs.EOF Then
    Beep
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.MoveFirst
End If

'Okay clear the message area
'StartSection = frmMain.MessageArea.SelStart
frmMain.MessageArea.Text = " "
For i = 1 To MSFlexGrid1.Row - 1
    rs.MoveNext
Next

'To get here it must be the right row
FileName = rs("Message File Name")
frmMain.Text2.Text = rs("Subject")
frmMain.Combo2.Text = rs("From/To")
frmMain.Text3.Text = rs("CC")
rs.Edit
rs("Message Read") = True
rs.Update

'Place into Message area
ReadMailMessage (FileName)

If Not rs("Attachment File Name") = "" Then
   
    FileName = rs("Attachment File Name")
    StartSection = frmMain.MessageArea.SelStart
    frmMain.MessageArea.SelStart = Len(frmMain.MessageArea)
    frmMain.MessageArea.SelText = "========attachment========" & vbCrLf & "Attachment has been saved as: " & App.Path & "\mailbox\" & FileName & vbCrLf
    frmMain.MessageArea.SelStart = Len(frmMain.MessageArea)
    frmMain.MessageArea.SelText = vbCrLf & "If the attachment has an ASC extension then use the 'PGP File Operations' to decrypt it"
    frmMain.MessageArea = frmMain.MessageArea & vbCrLf & "If the attachment has an ASC extension then use the 'PGP File Operations' to decrypt it"
End If
rs.Close
Set rs = Nothing
frmMain.Show
Exit Sub
BadFileFind:
    MsgBox Err.Description & " Can't find message", vbApplicationModal + vbCritical, App.Title
    Err.Clear
    rs.Close
    Set rs = Nothing
End Sub
