VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFileSafe 
   Caption         =   "PI File Safe"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8340
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   8100
      TabIndex        =   3
      Top             =   5790
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Align           =   1  'Align Top
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   8493
      _Version        =   131074
      PaneTree        =   "FileSafeForm.frx":0000
      Begin ComctlLib.ListView lvFileListView 
         Height          =   4755
         Left            =   2535
         TabIndex        =   1
         Top             =   30
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   8387
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SSActiveTreeView.SSTree SSTree1 
         CausesValidation=   0   'False
         Height          =   4755
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   8387
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
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "FileSafeForm.frx":0052
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":03A4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":06F6
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":0A48
            Key             =   "happy"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":0D62
            Key             =   "apathy"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":107C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":1396
            Key             =   "bulbon"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":16B0
            Key             =   "bulboff"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":19CA
            Key             =   "question"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":1CE4
            Key             =   "openlock"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":1FFE
            Key             =   "closedlock"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":2318
            Key             =   "exclamation"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":2632
            Key             =   "BrokenEnvelope"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":294C
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":2A9E
            Key             =   "Mask"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":2BF0
            Key             =   "EncryptedEnvelope"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":3142
            Key             =   "Closed Envelope"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":3694
            Key             =   "Open Envelope"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":3BE6
            Key             =   "FolderGroup"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":3DC0
            Key             =   "DragIcon"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":40DA
            Key             =   "DropIcon"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":43F4
            Key             =   "New Folder"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileSafeForm.frx":4836
            Key             =   "Open Eye"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mFileCreateFolderGroup 
         Caption         =   "Create Folder Group"
      End
      Begin VB.Menu mFileCreateSubFolder 
         Caption         =   "Create Sub Folder"
      End
      Begin VB.Menu mFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mPGP 
      Caption         =   "PGP Options"
   End
   Begin VB.Menu mnuTree1PopUpNewFolders 
      Caption         =   "TreeList1Context Menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu PopupNewFolderGroup 
         Caption         =   "NewFolderGroup"
         Enabled         =   0   'False
      End
      Begin VB.Menu PopupNewSubFolderGroup 
         Caption         =   "NewSubFolderGroup"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmFileSafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_GroupNodeMarkedForDelete As Boolean
Private Sub Form_Load()
Dim Win As New CWindow
Dim App As New CApplication

Dim i As Integer
'Dim Headings As String
Dim RowIndex As Integer
Dim ItemString As String
Dim Sqlq As String

'==============================
'Check for previous instance
'================================
If frmMain.ProgramIsAlreadyRunning Then Stop

Win.Center Me, Null
On Error Resume Next
'Make sure directories exisit

If Dir(App.Path & "FileSafe\") = "" Then
    MkDir App.Path & "FileSafe"
End If
On Error GoTo BadLoad

'If SecurityCheck = "INVALID" Then
'    gFullRelease = CheckLicenceExpired
'End If
'Me.Caption = "Private Email Mail (" & App.Version & ") for Win9x/Win2k/NT"
'Default to Mail Tree
'Call InitialiseGrid
'SSTree1(0).ImageList = ImageList1
'BuildTree (0)
'DisplayInBox

'Now load the file tree
SSTree1(0).ImageList = ImageList1
'BuildTree (1)
DisplayFileList
'RestoreMainSettings
'
'Lastly check to see if PGP is loaded
'

'
' Align all the controls
'
'InitialiseProgressBar
Exit Sub
BadLoad:
    MsgBox Err.Description & " Can't load properly...", vbApplicationModal + vbCritical, App.Title
    Err.Clear
    Resume Next
End Sub

Private Sub Form_Resize()
'Dim BottomMargin As Integer
'Dim LeftMargin As Integer
'Static StatusTop As Long
'DoEvents
On Error Resume Next
  ' BottomMargin = 800
   'LeftMargin = 200
   DoEvents
   If WindowState <> 1 Then
        DoEvents
       ' StatusBar.Panels(0).Left = MessageArea.Left
        StatusBar.Panels(1).Width = 3 * Me.Width / 4 '- StatusBar.Panels(1).Left
        StatusBar.Panels(2).Width = Me.Width / 4 - 20 '- StatusBar.Panels(2).Left
        'StatusBar.Panels(3).Width = MessageArea.Width / 3 '- StatusBar.Panels(3).Left
        DoEvents
        SSTree1(0).Height = StatusBar.Top
        SSSplitter1.Height = StatusBar.Top
        DoEvents
        ProgressBar1.Left = StatusBar.Panels(2).Left + 40
        ProgressBar1.Width = 8 * StatusBar.Panels(2).Width / 10
        ProgressBar1.Height = 8 * StatusBar.Height / 10
        ProgressBar1.Top = StatusBar.Top + 0.6 * (StatusBar.Height - ProgressBar1.Height)
        DoEvents
        'PIForm(gActivePIInstance).Enabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFileSafe = Nothing
End Sub

Private Sub lvFileListView_BeforeLabelEdit(Cancel As Integer)
'Dim i As Integer
'i = 1
End Sub

Private Sub lvFileListView_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Static SortOrder As Integer
' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
If SortOrder = 0 Then
    lvFileListView.SortOrder = lvwDescending
Else
    lvFileListView.SortOrder = lvwAscending
End If
SortOrder = Not SortOrder
lvFileListView.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lvFileListView.Sorted = True

End Sub

Private Sub lvFileListView_DblClick()
Dim res As Long
Dim obj As Object
Dim Attachment As String
Dim FileToLaunch As String
Dim msg As String
Dim iRes As Integer
Dim sTemporaryFile As String
'If Not gTemporaryFile = "" Then Kill gTemporaryFile

On Error GoTo BadAttachmentLaunch

Attachment = lvFileListView.SelectedItem.Text
'
'If it is asc then decode it within this app.
'
'tachment = lvFileListView.SelectedItem.Index
Attachment = lvFileListView.ListItems(lvFileListView.SelectedItem.Index).SubItems(4)
If GetExt(Attachment) = "asc" Then
    frmAttachmentOptions.lblFileName = Attachment
    frmAttachmentOptions.Caption = "Decrypt File"
    frmAttachmentOptions.Show vbModal
    iRes = frmAttachmentOptions.iRes
    Set frmAttachmentOptions = Nothing
    Select Case iRes
        Case vbYes
            
            vb2spgpContext.Initialise
            vb2spgpContext.FileIn = App.Path & "\FileSafe\" & Attachment '
            vb2spgpContext.FileOut = App.Path & "\temp\" & StripExt(Attachment) & "." & lvFileListView.ListItems(lvFileListView.SelectedItem.Index).SubItems(2)
            spgpDecryptFile
            FileToLaunch = vb2spgpContext.FileOut
        Case vbNo
            CommonDialog1.DialogTitle = "Decrypt " & Attachment & " and Save file as:"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "File Type (*." & GetExt(StripExt(Attachment)) & ")"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.FileName = IIf(InStr(1, StripExt(Attachment), ".") = 0, StripExt(Attachment) & ".htm", StripExt(Attachment)) 'TempPathLocation & StripExt(lvwAttachments.SelectedItem.Text)
            CommonDialog1.DefaultExt = GetExt(StripExt(Attachment))
            CommonDialog1.Action = 2
            ChDrive Mid(App.Path, 1, 3)
            ChDir App.Path
            vb2spgpContext.Initialise
            vb2spgpContext.FileIn = Attachment
            vb2spgpContext.FileOut = CommonDialog1.FileName '& CommonDialog1.DefaultExt
            spgpDecryptFile
            ShowStatus 1, "Decrypted file was saved successfully"
            Exit Sub
        Case vbCancel
            Exit Sub
        End Select
 
 Else
    FileToLaunch = Attachment
 
 End If
'
'Now try and launch it if we can
'
    DoEvents
    sTemporaryFile = FileToLaunch
    res = ShellExecute(Me.hWnd, "open", FileToLaunch, vbNullString, CurDir, SW_SHOW)
    DoEvents
    If res < 32 Then
             Kill FileToLaunch 'TempPathLocation & lvwAttachments.SelectedItem.Text
             Err.Raise "Error was encountered launching the application associated with this attachment.   Please check your " & TempPathLocation & " directory to makes sure there are no plain text (decrypted) files there."
    End If
    'Need these otherwise the temp file be deleted before the application is lauched
    ShowStatus 1, "Temporary file: " & vb2spgpContext.FileOut & " exists!"
    'Kill vb2spgpContext.FileOut
Exit Sub
BadAttachmentLaunch:
    MsgBox "An error has been encountered: " & Err.Description, vbApplicationModal + vbCritical, "Attachment Launch"
    Err.Clear
End Sub

Private Sub lvFileListView_DragDrop(Source As Control, x As Single, y As Single)
'MousePointer = vbDefault
'Exit Sub
lvFileListView.MousePointer = vbDefault
'lvFileListView.Drag vbCancel
'lvFileListView.MousePointer = ccDefault
'MousePointer = vbDefault
'MSFlexGrid1.DragIcon = Nothing
'MSFlexGrid1.MousePointer = flexDefault

DoEvents
Exit Sub
lvFileListView.MousePointer = vbDefault
Exit Sub
'MSFlexGrid1.Drag vbEndDrag
'MSFlexGrid1.MousePointer = vbDefault
lvFileListView.MousePointer = vbDefault

'If Not Source.name = "lvFileListView" Then
    'MousePointer = vbDefault
    'Source.DragIcon = ImgList(5).Picture
   ' MousePointer = vbNoDrop
  ' lvFileListView.DragIcon = Nothing
   ' lvFileListView.Drag vbEndDrag
   '''
   
   '''
   
   
   ' gDragCommenced = False
'End If
End Sub

Private Sub lvFileListView_DragOver(Source As Control, x As Single, y As Single, State As Integer)
'If Not Source.name = "lvFileListView" Then
   ' lvFileListView.DragIcon = Nothing
    'lvFileListView.Drag 0
    'Source.name.Drag 0
    'lvFileListView.DragIcon = Nothing
    'Source.DragIcon = ImgList(5).Picture
   ' MSFlexGrid1.MousePointer = vbNoDrop
   'MSFlexGrid1.Drag vbNoDrop 'EndDrag
   'MSFlexGrid1.MousePointer = vbNoDrop
  ' MSFlexGrid1.Drag 0
    'MSFlexGrid1.MousePointer = vbNoDrop
    'MSFlexGrid1.Drag 0
    On Error Resume Next
    
   If Not Source.name = "lvFileListView" Then
        Source.Value.Drag vbEndDrag
        lvFileListView.Drag vbEndDrag
        lvFileListView.MousePointer = ccNoDrop
   End If

End Sub

Private Sub lvFileListView_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'Exit Sub
If KeyCode = vbKeyDelete Then
    DeleteEncryptedFile
    FillListView
    'This selects a single row again to stop multiple deletes
   ' Grid.RowSelection = Grid.SelectedRow
End If
End Sub

Private Sub lvFileListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'gDragCommenced = False
MousePointer = vbDefault
End Sub

Private Sub lvFileListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static MousePressed As Boolean

If Not Button = vbLeftButton Then
    lvFileListView.MousePointer = ccArrow
    MousePressed = False
    Exit Sub
Else
    If MousePressed Then Exit Sub
End If
If Button = vbLeftButton Then
    MousePressed = True
    'Keep track of present folder
    SSTree1(0).Tag = SSTree1(0).SelectedItem.Index
    lvFileListView.DragIcon = frmMain.ImgList(4).Picture
    lvFileListView.Drag vbBeginDrag
    
End If


End Sub

Private Sub lvFileListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lvFileListView.MousePointer = ccArrow
'vFileListView.DragIcon = Nothing
'vFileListView.Drag vbEndDrag
'gDragCommenced = False
'vFileListView.MousePointer = vbNoDrop 'ccNoDrop
End Sub



Private Sub mFileCreateFolderGroup_Click()
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
     '   Set rs = DB.OpenRecordset("Nodes", dbOpenDynaset)
   ' Else
        Set rs = DB.OpenRecordset("File Nodes", dbOpenDynaset)
   ' End If
   
    rs.AddNew
    rs("Node Name") = sFolderName
    rs("Can Delete") = True
    rs.Update
    rs.Close
    Set rs = Nothing
    BuildTree (0)
 
'End If
End Sub

Private Sub mFileCreateSubFolder_Click()
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

'n.Text = "fdsfasdffsd24242dssfs"

If FolderName = "" Then
    gCancelAction = True
    Exit Sub
End If
'FolderName = Tree(Index).CreateFolderName

'lFolderIndex = SSTree1(Index).SelectedItem.Index
'Set n = SSTree1(Index).Nodes.Item(lFolderIndex)

'If Index = 0 Then
'    Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
'Else
    Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
'End If
rsNode.FindFirst "[Node Name] =" & "'" & n.Text & "'"
If Not rsNode.NoMatch Then
   ' If Index = 0 Then
   '     Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
  '  Else
        Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
   ' End If
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

Private Sub mFileExit_Click()
Unload Me
End Sub

Private Sub mPGP_Click()
frmPGPOptions.Show vbModal
End Sub

Private Sub PopupNewFolderGroup_Click()
Call mFileCreateFolderGroup_Click
End Sub

Private Sub PopupNewSubFolderGroup_Click()
Call mFileCreateSubFolder_Click
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
Dim rsTemp As Recordset
Dim rsFolder As Recordset
Dim rsFile As Recordset
Dim qd As QueryDef
'Exit Sub
'We have to first check to see if we are in a group folder
'If Not Left(Node.FullPath, InStrRev(Node.FullPath, "\")) = "" Then
   ' NodeName = Left(Node.FullPath, InStrRev(Node.FullPath, "\") - 1) 'strip off the "\"
'Else
   ' NodeName = Node
'End If
If m_GroupNodeMarkedForDelete Then
    'We are going to delete a group folder here
    If Index = 0 Then
        Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
    Else
        Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
    End If
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
    '    Set rsFolder = DB.OpenRecordset("Folders", dbOpenDynaset)
   ' Else
        Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
   ' End If
    rsFolder.FindFirst "[Folder] =" & "'" & StripItemCount(Node.Text) & "'"
    
    If Not rsFolder.NoMatch Then
       ' If Index = 0 Then
         '   Set qd = DB.QueryDefs("qdMessagesinFolder")
         '   qd.Parameters![FolderId] = rsFolder("Folder ID")
       ' Else
            Set qd = DB.QueryDefs("qdFilesinFolder")
            qd.Parameters![FolderId] = rsFolder("Folder ID")
       ' End If
        Set rsFile = qd.OpenRecordset()
        'Just move to deleted folder
        Set rsTemp = DB.OpenRecordset("File Folders", dbOpenDynaset)
        rsTemp.FindFirst "[Folder] =" & "'" & "Deleted Files" & "'"
        If rsTemp.EOF Then
            MsgBox "Can't find system: Deleted Folders.  Files will be deleted without being saved in the Deleted Folders.", vbApplicationModal + vbCritical, "Save Deleted files error"
            Do While Not rsFile.EOF
                rsFile.Delete
                rsFile.MoveNext
            Loop
        Else
            Do While Not rsFile.EOF
                rsFile.Edit
                rsFile("Files.Folder ID") = rsTemp("Folder ID")
                rsFile.Update
                rsFile.MoveNext
            Loop
        End If
        rsFolder.Delete
    Else
        Beep
        MsgBox "Folder is not in the database.  PI is continuing anyway."
    End If
End If

Set rsNode = Nothing
Set rsFolder = Nothing
'Set rsMessage = Nothing

    FillListView

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
    sPreviousNode = ""
End If
End Sub

Private Sub SSTree1_BeforeNodeDelete(Index As Integer, Node As SSActiveTreeView.SSNode, Cancel As SSActiveTreeView.SSReturnBoolean, DispPromptMsg As SSActiveTreeView.SSReturnBoolean)
'Dim rsFolder As Recordset
Dim rsNode As Recordset
Dim rsFolder As Recordset
'Look for root or node
Dim n As SSNode
Dim lFolderIndex As Long

'
'First Find the node name
'
'lFolderIndex = SSTree1(Index).SelectedItem.Index
'Set n = SSTree1(Index).Nodes.Item(lFolderIndex)



m_GroupNodeMarkedForDelete = False
'Tree(Index).Initialise
If Not Node.Children = 0 Then
    MsgBox "Delete sub-folders first", vbApplicationModal + vbCritical
    Cancel = True
    Exit Sub
End If
'lFolderIndex = SSTree1(Index).SelectedItem.Index
'Set n = SSTree1(Index).Nodes.Item(lFolderIndex)

'Okay now search the database for the sub folder or group name
If Node.Text = "" Then
    'We are going to delete a group folder here
    If Index = 0 Then
        Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
    Else
        Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
    End If
    rsNode.FindFirst "[Node Name] =" & "'" & StripItemCount(Node.Text) & "'"
    If Not rsNode.NoMatch Then
        If Not rsNode("Can Delete") Then
            Beep
            Cancel = True
            MsgBox "Can't delete Group Folder", vbQuestion + vbApplicationModal, "Group Folder"
        Else
            m_GroupNodeMarkedForDelete = True
        End If
    Else
            Beep
            MsgBox "Group Folder is not in the database.  PI is continuing anyway."
            m_GroupNodeMarkedForDelete = True
    End If
Else
    'Okay we are going to delete a folder here
    Set rsFolder = DB.OpenRecordset("File Folders", dbOpenDynaset)
    rsFolder.FindFirst "[Folder] =" & "'" & StripItemCount(Node.Text) & "'"
    If Not rsFolder.NoMatch Then
        If Not rsFolder("Can Delete") Then
            Beep
            Cancel = True
            MsgBox "Can't delete a system folder", vbQuestion + vbApplicationModal, "System Folder"
        Else
            m_GroupNodeMarkedForDelete = False
        End If
    Else
        Beep
        MsgBox "Folder is not in the database.  PI is continuing anyway."
    End If
End If

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
Dim lFolderIndex As Long
Dim rs As Recordset
Dim StoredFileName As String

On Error Resume Next
'lFolderIndex = SSTree1(Index).SelectedItem.Index
'Set ssNodeX = SSTree1(Index).Nodes.Item(lFolderIndex)
'ssNodeX.Previous
MousePointer = vbHourglass
DoEvents
'ssNodeTmp.Tag = SSTree1(Index).Tag
Set ssNodeTmp = SSTree1(Index).HitTest(x, y)
If ssNodeTmp Is Nothing Then Exit Sub
'Set SSTree1(Index).SelectedItem = ssNodeTmp
If Source.name = "MSFlexGrid1" Or Source.name = "lvFileListView" Then
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

    
        For i = 1 To lvFileListView.ListItems.Count
            If lvFileListView.ListItems(i).Selected = True Then
                StoredFileName = lvFileListView.ListItems(i).SubItems(4)
                'StoredFileName = lvFileListView.SelectedItem.SubItems(4)
                ShowStatus 1, "Moving file: " & StoredFileName
                Set rs = DB.OpenRecordset("Files", dbOpenDynaset)
                rs.FindFirst "[Stored File Name] =" & "'" & StoredFileName & "'"
                If rs.NoMatch Then
                    MsgBox "File not found.  Database error", vbApplicationModal + vbCritical, "Database Error"
                Else
                    rs.Edit
                    rs("Folder ID") = NewFolderID
                    rs.Update
                End If
            End If
        Next
            
End If
'ssNodeTmp..Previous.Selected = True
Set rs = Nothing
'SSTree1(Index).Tag = SSTree1(Index).SelectedItem
SSTree1(Index).SelectedItem = SSTree1(Index).Nodes(CInt(SSTree1(Index).Tag))
'ssNodeTmp.Text = SSTree1(Index).Tag
FillListView
lvFileListView.Drag vbEndDrag
'End If

AddFolderParameters (Index)
ssNodeTmp.Selected = False
MousePointer = vbDefault
ShowStatus 1, "Completed successfully.."
AddFolderParameters (Index)
End Sub

Private Sub SSTree1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
Dim ssNodeTmp As SSNode

If Index = 1 Then
    If Not Source.name = "lvFileListView" Then
        SSTree1(Index).Drag 0
        Exit Sub
    End If
End If
'If Index = 0 Then
'    If Not Source.name = "MSFlexGrid1" Then
 '       SSTree1(Index).Drag 0
 '       Exit Sub
 '   End If
'End If
'SSTree1(Index).Tag = SSTree1(Index).SelectedItem
SSTree1(Index).SetFocus
'Dim s As String

'SSTree1(Index).Tag = SSTree1(Index).Node
Set ssNodeTmp = SSTree1(Index).HitTest(x, y)
If Not ssNodeTmp Is Nothing Then
    Set SSTree1(Index).SelectedItem = ssNodeTmp
    Set SSTree1(Index).DropHighlight = Nothing
    'ssNodeTmp.Font.bold = True
End If

End Sub

Private Sub SSTree1_Expand(Index As Integer, Node As SSActiveTreeView.SSNode)
Dim rs As Recordset

'If it is at the folder level open it up and show the contents
'If Index = 0 Then
    'Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
'Else
  '  Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
'End If
'rs.FindFirst "Folder =" & "'" & Node & "'"
'I'f rs.NoMatch Then
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
'Else
'    Tree(Index).FolderId = rs("Folder ID")
'End If

'rs.Close
'Set rs = Nothing
'If Index = 0 Then
 '   InitialiseGrid ("AC")
  '  FillGrid
'End If
End Sub

Private Sub SSTree1_GotFocus(Index As Integer)
'giTreeIndex = Index
End Sub

Private Sub SSTree1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As String
 If Button = 2 Then   ' Check if right mouse button
                       ' was clicked.
      'PopupMenu frmMain.mnuTreeListContext   ' Display the File menu as a
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
'If Tree(Index).CancelClick Then
 '   Tree(Index).CancelClick = False
 '   Exit Sub
'End If

'Tree(Index).Initialise
' Index = 0 Then Grid.Initialise

lFolderIndex = SSTree1(Index).SelectedItem.Index
Set n = SSTree1(Index).Nodes.Item(lFolderIndex)


'Look for root or node
'If InStr(Node.FullPath, "\") > 0 Then
   ' Tree(Index).FolderName = n.Text
   ' Tree(Index).NodeName = n.Parent 'Left(n.FullPath, InStrRev(Node.FullPath, "\") - 1)
'Else
'    Tree(Index).NodeName = n.Text
'End If
FillListView
If n.Parent Is Nothing Then Exit Sub
'If Not n.Parent = "" Then
 '   If Index = 0 Then
  '      Set rs = DB.OpenRecordset("Nodes", dbOpenDynaset)
  '  Else
   '     Set rs = DB.OpenRecordset("File Nodes", dbOpenDynaset)
   ' End If
   ' rs.FindFirst "[Node Name] =" & "'" & StripItemCount(n.Parent) & "'"
  '  If rs.NoMatch Then
  '      Err.Raise 3004, "Can't find the node."
   ' Else
       ' Tree(Index).NodeName = rs("Node Name")
   '     'Tree(Index).NodeID = rs("Node ID")
  '  End If
'End If
'Check to see if it is a folder group
'MousePointer = vbHourglass

'If Not Left(Node.FullPath, InStrRev(Node.FullPath, "\")) = "" Then
 '   If Index = 0 Then
  '      Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
   ' Else
   '     Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
   ' End If
   ' rs.FindFirst "[Node Id] =" & Tree(Index).NodeID
   ' Do While Not rs.EOF 'Until
    '    If rs("Folder") = StripItemCount(n.Text) Then
     '       Tree(Index).FolderId = rs("Folder ID")
    '        'Tree(Index).FolderName = rs("Folder")
    '        Exit Do
    '    End If
    '    rs.MoveNext
   ' Loop

   FillListView

 MousePointer = vbDefault
 AddFolderParameters (Index)
Exit Sub
NodeError:
    ShowStatus 1, "Error: " & Err.Description
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


Private Sub SSTree1_OLEDragDrop(Index As Integer, Data As SSActiveTreeView.SSDataObject, Effect As SSActiveTreeView.SSReturnLong, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ssNodeTmp As SSNode
Dim ssNodeX As SSNode
Dim Sqlq As String
Dim i As Integer
Dim lWidth As Long
Dim Heading As String
Dim NewFolderID As Long
Dim rs As Recordset
Dim iResponse As Integer
Dim sFileName As String
Dim sStoredFileName As String
Dim oFso As New FileSystemObject
Dim ofile As File
Dim PGPKeyID As String
Dim DB_Index As Long
Dim qd As QueryDef
Dim sFolderName As String
'Dim rs As Recordset

Dim lFolderIndex As Long
Dim n As SSNode

On Error Resume Next
'
'First Find the node name
'
lFolderIndex = SSTree1(Index).SelectedItem.Index
Set n = SSTree1(Index).Nodes.Item(lFolderIndex)

On Error GoTo DropError
If Index = 0 Then
    Effect = ssOLEDropEffectNone
    Exit Sub
End If


'sFileName = Data.Files.Count
MousePointer = vbHourglass 'vbNoDrop
Set ssNodeTmp = SSTree1(Index).HitTest(x, y)
If ssNodeTmp Is Nothing Then Exit Sub
'SSTree1(0).Nodes.Item(2).Expanded = True
Set SSTree1(Index).SelectedItem = ssNodeTmp
sFolderName = StripItemCount(CStr(ssNodeTmp))
'Sqlq = SSTree1(Index).name
'Now store the files

GoTo x

Set SSTree1(Index).DropHighlight = SSTree1(Index).HitTest(x, y)

    If Not SSTree1(Index).DropHighlight Is Nothing Then
        If SSTree1(Index).DropHighlight.Level > 1 Then
            Effect = ssOLEDropEffectNone
            Exit Sub
        End If

        Effect = ssOLEDropEffectCopy
    End If
x:
'''''''''''
'If sFileName = MSFlexGrid1 Then
    'SSTree1.DragIcon = ImgList(5).Picture 'LoadPicture("d:\pi32\icons\drop1pg.ico") 'ImageList1(0) 'drop icon
    'gSelectedFolder = Node
    Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
   ' rs.FindFirst "Node ID =" & "'" & ssNodeTmp & "'"
    rs.FindFirst "[Folder] ='" & sFolderName & "'"
    If rs.NoMatch Then
        Beep
        Set rs = Nothing
        MsgBox "Folder does not exist in dababase.", vbApplicationModal + vbCritical, "Folder Error"
        Exit Sub
    Else
        'rs.MoveFirst
        NewFolderID = rs("Folder ID")
    End If
   'Tree (Index).FolderId = NewFolderID
'
' Okay initialise PGP stuff
'
    'If Not mPGP.Enabled Then
   '     MsgBox "Can't encrypt file unless PGP is installed.", vbApplicationModal + vbCritical, "Can't find PGP"
   '     Exit Sub
  '  End If
        
            SSTree1(Index).MousePointer = vbDefault
       ' End If
            
      '  If gPGPKeyID = "" Then
      '      Beep
      '      Exit Sub
      '  Else
      '      vb2spgpContext.SignKeyID = gPGPKeyID
      '  End If
 'ProgressBar1.Left = lblstatus.Left + lblstatus.Width + 10
 ProgressBar1.Visible = True
 Set rs = DB.OpenRecordset("Files", dbOpenDynaset)
 ProgressBar1.Max = Data.Files.Count
 'lblstatus.Visible = True
 For i = 1 To Data.Files.Count
 ShowStatus 1, "Checking file " & i
 DoEvents
Dim rsfiles As Recordset
Set qd = DB.QueryDefs("qdFilesInFolder")
qd.Parameters![FolderId] = NewFolderID
Set rsfiles = qd.OpenRecordset()
 
    rsfiles.FindFirst "[File Name] ='" & StripExt(StripFileName(Data.Files(i))) & "'"
    If Not rsfiles.NoMatch Then
        If rsfiles("File Type") = GetExt(Data.Files(i)) Then
            iResponse = MsgBox("The file '" & StripFileName(Data.Files(i)) & "' already exists in this folder.  To replace the current version click the 'YES' button and click 'NO' to skip this file.", vbYesNoCancel + vbQuestion + vbApplicationModal, "File Exists")
            Select Case iResponse
                Case vbCancel
                    Set rsfiles = Nothing
                    Set rs = Nothing
                    Exit Sub
                Case vbYes
                    rs.FindFirst "[File ID] =" & rsfiles("File ID")
                    rs.Edit
                    Set rsfiles = Nothing
                Case vbNo
                    GoTo NextExit
            End Select
        Else
            rs.AddNew
        End If
    Else
        rs.AddNew
    End If
    
'Place the file reference into the database

    sFileName = Data.Files(i)
    Set ofile = oFso.GetFile(sFileName)
    DB_Index = rs("File ID")
    rs("Folder ID") = NewFolderID
    rs("File Name") = StripExt(StripFileName(sFileName))
    rs("File Size") = ofile.Size
    rs("File Type") = GetExt(Data.Files(i))
    rs("Date Modified") = ofile.DateLastModified
    Randomize (DB_Index)
    sStoredFileName = App.Path & "\FileSafe\" & Format(10000000 * (Rnd(DB_Index) + Rnd(DB_Index + 1)), "#######") & ".asc"
    rs("Stored File Name") = StripFileName(sStoredFileName)
    
    
    '
    'Now encrypt the file
    '
    ShowStatus 1, "Encrypting file: " & sFileName

    DoEvents
   ' vb2spgpContext.FileIn = sFileName
  '  vb2spgpContext.FileOut = sStoredFileName
   ' spgpEncryptfile
  '  Set ofile = oFso.GetFile(sStoredFileName)
   ' ofile.Attributes = ReadOnly
   ' Set ofile = Nothing
    
    'Don't update until this is all done.
    rs.Update

'Update after all is okay
    
    FillListView
    ProgressBar1.Value = i
    AddFolderParameters (Index)
    DoEvents
NextExit:
Next
rs.Close
ssNodeTmp.Selected = False
ProgressBar1.Visible = False
MousePointer = vbDefault
ShowStatus 1, " " 'lblstatus.Visible = False
'SSTree1(Index).SelectedItem.e
Set SSTree1(Index).DropHighlight = Nothing
'SSTree1(Index).SelectedItem
Exit Sub
DropError:
    MsgBox "An error has occured: " & Err.Description, vbApplicationModal + vbCritical, "Error in Drag and Drop"
    Err.Clear
    ProgressBar1.Visible = False
    rs.Close
    MousePointer = vbDefault
    Set SSTree1(Index).DropHighlight = Nothing
    'lblstatus.Visible = False
End Sub
Private Sub DeleteEncryptedFile()
Dim i As Integer
Dim qd As QueryDef
Dim StoredFileName As String
Dim StartRow As Integer
Dim EndRow As Integer
Dim DeletedFolderID As Long
Dim rsAttachment As Recordset
Dim iResponse As Integer
Static InHere As Boolean
Dim Sqlq As String
Dim rs As Recordset
Dim oFso As New FileSystemObject
Dim ofile As File

'If InHere Then Exit Sub

'InHere = True
On Error GoTo BadKeyEntry
Set rs = DB.OpenRecordset("File Folders", dbOpenDynaset)
rs.FindFirst "Folder =" & "'" & "Deleted Files" & "'"
If rs.NoMatch Then
    Err.Raise 1002, "System Folder has been deleted"
Else
    DeletedFolderID = rs("Folder ID")
End If
rs.Close
MousePointer = vbHourglass
For i = 1 To lvFileListView.ListItems.Count
    If lvFileListView.ListItems(i).Selected = True Then
        StoredFileName = lvFileListView.ListItems(i).SubItems(4)
        lvFileListView.ListItems(i).Ghosted = True
        Set rs = DB.OpenRecordset("Files", dbOpenDynaset)
        rs.FindFirst "[Stored File Name] =" & "'" & StoredFileName & "'"
        If rs.NoMatch Then
            Exit Sub
        Else
            If rs("Folder ID") = DeletedFolderID Then
            'If it is in the deleted then just delete
                rs.Delete
                ShowStatus 1, "Deleting file: " & StoredFileName
                Set ofile = oFso.GetFile(App.Path & "\FileSafe\" & StoredFileName)
                ofile.Attributes = Normal
                WipeFile (App.Path & "\FileSafe\" & StoredFileName)
                On Error GoTo BadKeyEntry
            Else
                ShowStatus 1, "Moving " & StoredFileName & " to 'Deleted Files Folder'."
                rs.Edit
                rs("Folder ID") = DeletedFolderID
                rs.Update
            End If
        End If
    End If
    DoEvents
Next
Set ofile = Nothing

'InHere = False
rs.Close
Set rs = Nothing
MousePointer = vbDefault
ShowStatus 1, "Completed successfully"
Exit Sub

BadKeyEntry:
    MsgBox Err.Description & " - There was an error deleting the file."
    Beep
    Err.Clear
    InHere = False
   ' FillGrid
    MousePointer = vbDefault
End Sub
Public Sub ShowStatus(Panel As Integer, Status As String)
StatusBar.Style = sbrNormal
If TextWidth(Status & "  ") > StatusBar.Panels(Panel).Width Then StatusBar.Panels(Panel).Width = TextWidth(Status & " ")
StatusBar.Panels.Item(Panel) = Status
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
Private Sub FillListView()
Dim rs As Recordset
Dim pos1 As Integer
Dim pos2 As Integer
Dim Sqlq As String
Dim RowIndex As Integer
Dim ItemString As String
Dim FromString As String
'Dim obj As SSNode
Dim ToString As String
Dim i As Integer
Dim sName As String
Dim itmX As ListItem

On Error GoTo FillFileListError
'Dim Node As SSActiveTreeView.SSNode
'Set Node = SSActiveTreeView

'ItemString = Node.FullPath
'set Node. .Text = "DDDD"
Me.MousePointer = vbHourglass
lvFileListView.Visible = False
lvFileListView.ListItems.Clear
'obj.no
        Sqlq = "Select DISTINCTROW * FROM Files WHERE "
        Sqlq = Sqlq & "(Files.[Folder ID] = " & GetSelectedFolderID(1) & ") "
        Sqlq = Sqlq & "ORDER BY Files.[File Name] DESC;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
   

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
End If


lvFileListView.View = lvwReport
While Not rs.EOF
      Set itmX = lvFileListView.ListItems.Add()
     ' itmX.Icon = ""
      itmX.Text = rs("File Name")
      itmX.SubItems(1) = rs("File Size")
      itmX.SubItems(2) = rs("File Type")
      itmX.SubItems(3) = rs("Date Modified")
      itmX.SubItems(4) = rs("Stored File Name")
             
    rs.MoveNext
Wend
'MSFlexGrid1.FormatString = Headings

rs.Close
Set rs = Nothing
   
    lvFileListView.Visible = True

Me.MousePointer = vbDefault

DoEvents
Exit Sub

FillFileListError:
    MsgBox "Error returned filling the file list: " & Err.Description, vbCritical + vbApplicationModal, "Internal Error"
    Err.Clear
    Me.MousePointer = vbDefault
End Sub
Public Sub DisplayFileList()

   Dim clmX As ColumnHeader
   Dim itmX As ListItem
   Dim objNode As SSNode
   Dim i As Integer
   Dim Headings(5) As String
   Headings(1) = "File Name"
   Headings(2) = "Size"
   Headings(3) = "Type"
   Headings(4) = "Modified"
   Headings(5) = "Stored File Name"

   ' Add 3 ColumnHeader objects to the control.
   For i = 1 To 5
      Set clmX = lvFileListView.ColumnHeaders.Add()
      clmX.Text = Headings(i)
   Next i
   '
   'Now show the tree
   BuildTree (0)
    SSTree1(0).Nodes.Item(1).Expanded = True
  ' i = SSTree1(0).Nodes.Count
'For i = 1 To SSTree1(1).Nodes.Count
 '   If UCase(SSTree1(1).Nodes(i)) = UCase("Deleted Files") Then Exit For
'Next i
'This is the Inbox Node Object
'Set objNode = SSTree1(1).Nodes(i)
'objNode.Selected = True
   
   
   ' Set View to Report.
 '  lvFileListView.View = lvwReport
   
   ' Add 10 ListItems to the control.
 '  For i = 1 To 10
  '    Set itmX = lvFileListView.ListItems.Add()
  '    itmX.Text = "ListItem " & i
  ''    itmX.SubItems(1) = "Subitem 1"
  '    itmX.SubItems(2) = "Subitem 2"
 '  Next i


End Sub
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
    'If Index = 0 Then
     '   Set rsNode = DB.OpenRecordset("Nodes", dbOpenDynaset)
     '   Set n = SSTree1(Index).Nodes.Add(, , , "Private Idaho E-Mail", "EncryptedEnvelope", "EncryptedEnvelope")
   ' Else
        Set rsNode = DB.OpenRecordset("File Nodes", dbOpenDynaset)
        Set n = SSTree1(Index).Nodes.Add(, , , "Private Idaho File Safe", "EncryptedEnvelope", "EncryptedEnvelope")
    
    'End If
    n.Font.Bold = True
    Set n = Nothing
  ' If Not rsNode.EOF Then rsNode.MoveFirst
    Do While Not rsNode.EOF
        sNodeName = rsNode("Node Name")
        Set n = SSTree1(Index).Nodes.Add(, , sNodeName, sNodeName, "FolderGroup", "FolderGroup")
        n.Font.Bold = True
        n.Expanded = True
       ' Set n = Nothing
       ' If Index = 0 Then
         '   Set qd = DB.QueryDefs("qdSubFolders")
       ' Else
            Set qd = DB.QueryDefs("qdFileSubFolders")
       ' End If
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
 'n.Expanded = True
Exit Sub
BadTreeBuild:
    MsgBox "Private Idaho Error:  " & Err.Description, vbCritical + vbApplicationModal, "Build Tree"
    Err.Clear
    Exit Sub

End Sub
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
