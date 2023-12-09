VERSION 5.00
Object = "{33337253-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "nntp40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNewsReader 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "News Reader and Poster"
   ClientHeight    =   7425
   ClientLeft      =   1185
   ClientTop       =   1770
   ClientWidth     =   8715
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7425
   ScaleWidth      =   8715
   Begin Threed.SSFrame SSFrame2 
      Height          =   1755
      Left            =   180
      TabIndex        =   20
      Top             =   2280
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   3096
      _Version        =   131074
      BackStyle       =   1
      Caption         =   "Personal Info"
      Begin VB.TextBox tFrom 
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
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox tUserID 
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
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox tPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave User ID and Password blank if not required!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   1500
         TabIndex        =   24
         Top             =   1320
         Width           =   4005
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1755
      Index           =   0
      Left            =   5880
      TabIndex        =   16
      Top             =   1860
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3096
      _Version        =   131074
      BackStyle       =   1
      Caption         =   "Attached File - Options"
      Begin VB.TextBox tSegmentSize 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   1500
         TabIndex        =   17
         Text            =   "0"
         ToolTipText     =   "Enter File Segment Size"
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Size"
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
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   35
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label tFileSize 
         BackStyle       =   0  'Transparent
         Caption         =   "FileSize"
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
         Left            =   1020
         TabIndex        =   34
         Top             =   660
         Width           =   1515
      End
      Begin Threed.SSCheck CheckSplitFile 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   131074
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Split File into segments"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File/Segment Size"
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
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No. File Segments: "
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
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   1500
         TabIndex        =   18
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Browse"
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
      Left            =   5100
      TabIndex        =   15
      Top             =   1890
      Width           =   675
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
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
      Left            =   6300
      TabIndex        =   14
      Top             =   4200
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8250
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tSubject 
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
      Left            =   1140
      TabIndex        =   7
      Top             =   1320
      Width           =   4605
   End
   Begin VB.TextBox tAttachedFile 
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
      Left            =   1140
      TabIndex        =   6
      Tag             =   "File to be encoded"
      ToolTipText     =   "Select attachment file from File menu"
      Top             =   1890
      Width           =   3885
   End
   Begin RichTextLib.RichTextBox rtMessageArea 
      Height          =   2355
      Left            =   120
      TabIndex        =   5
      Top             =   4860
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4154
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"News Agent.frx":0000
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
   Begin VB.TextBox tNewsServer 
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
      Left            =   1140
      TabIndex        =   1
      Text            =   "news"
      Top             =   600
      Width           =   4635
   End
   Begin VB.TextBox tNewsgroup 
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
      Left            =   1140
      TabIndex        =   0
      Text            =   "alt.test"
      Top             =   960
      Width           =   4605
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1335
      Index           =   1
      Left            =   5880
      TabIndex        =   30
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2355
      _Version        =   131074
      BackStyle       =   1
      Caption         =   "Header Options"
      Begin Threed.SSCheck chkXNoArchive 
         Height          =   315
         Left            =   300
         TabIndex        =   31
         Top             =   300
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   131074
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "X-No-Archive:"
         Value           =   1
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3690
      TabIndex        =   32
      Top             =   4230
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblsplitstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "File split status:"
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
      Left            =   2580
      TabIndex        =   33
      Top             =   4260
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Threed.SSCommand bPost 
      Height          =   375
      Left            =   6300
      TabIndex        =   13
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   131074
      ForeColor       =   0
      Caption         =   "Post Article"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Area"
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
      Index           =   11
      Left            =   120
      TabIndex        =   12
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "status"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4260
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment:"
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
      Height          =   285
      Index           =   8
      Left            =   150
      TabIndex        =   10
      Top             =   1890
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attached File Name (Select from File menu)"
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
      Height          =   255
      Index           =   2
      Left            =   1380
      TabIndex        =   8
      Top             =   1680
      Width           =   3585
   End
   Begin NNTPLibCtl.NNTP NNTP1 
      Left            =   7830
      Top             =   0
      CurrentArticle  =   ""
      CurrentGroup    =   ""
      NewsServer      =   ""
      Password        =   ""
      User            =   ""
      WinsockLoaded   =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"News Agent.frx":0082
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Newsgroups:"
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
      Left            =   90
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "News Server:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mAttachFile 
         Caption         =   "Attach a file"
      End
      Begin VB.Menu mDecodeArticle 
         Caption         =   "Decode Article"
      End
      Begin VB.Menu spli 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mPostArticle 
      Caption         =   "Post Article"
   End
End
Attribute VB_Name = "frmNewsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rNewsgroups(1 To 100) As String
Dim rLastArticle(1 To 100) As Integer
Dim rNrNewsgroups As Integer

Dim rSubjectKeywords(1 To 100) As String
Dim rNrSubjectKeywords As Integer

Dim rFromKeywords(1 To 100) As String
Dim rNrFromKeywords As Integer

Dim rBodyKeywords(1 To 100) As String
Dim rNrBodyKeywords As Integer

Dim rSearchWeight As Integer
Dim rArticleFrom As String
Dim rArticleDate As String
Dim rArticleSubject As String
Dim rArticleBody As String

Private Sub bPost_Click()
Dim SectionName As String
Dim i As Integer
Dim Segments As Integer
Dim SplitStatus As Integer
'Dim FileLocation(1 To 1000)   As String
Dim SplitCount As Long

bPost.Enabled = False

If tNewsServer = "" Then
    If MailConnector.NNTPServerName = "" Then
        Beep
        Exit Sub
    End If
    tNewsServer = MailConnector.NNTPServerName
Else
    MailConnector.NNTPServerName = tNewsServer
    SectionName = "Options"
    WriteProfile SectionName, "NNTPServerName", MailConnector.NNTPServerName
End If

On Error GoTo BadPost
If Not tUserID = "" Then
    NNTP1.Password = tPassword
    NNTP1.User = tUserID
End If
'
' Okay split the file if necessary
'

If CLng(tSegmentSize) > 0 Then
    SplitCount = FileLen(tAttachedFile) / CLng(tSegmentSize)
    If SplitCount > 1000 Then
        MsgBox "Number of file segments that will result from the file size you have used will be too large.  The maximum number of files this program can generate is 1000.", vbCritical + vbApplicationModal, "File Split Error."
        lblstatus = "Post unsuccessful."
        lblstatus = ""
        bPost.Enabled = True
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    lblsplitstatus.Visible = True
    ProgressBar1.Visible = True
    If SplitCount > 1 Then SplitStatus = SplitFile(tAttachedFile, Val(tFileSize), Val(tSegmentSize), ProgressBar1, Segments)
    lblsplitstatus.Visible = False
    ProgressBar1.Visible = False
    If Not SplitStatus = 0 Then
        MsgBox "Error splitting the file.", vbCritical + vbApplicationModal, "File Split Error."
        lblstatus = "Post unsuccessful."
        lblstatus = ""
        bPost.Enabled = True
        Me.MousePointer = vbDefault
        Exit Sub
    End If
Else
    SplitCount = 1
End If
NNTP1.WinsockLoaded = True
NNTP1.NewsServer = MailConnector.NNTPServerName
NNTP1.Action = 10 'reset headers
lblstatus = "Connecting to server..."

NNTP1.Action = a_Connect 'connect to server
lblstatus = "Connected..."
NNTP1.OtherHeaders = NNTP1.OtherHeaders & IIf(Abs(chkXNoArchive.Value) = ssCBChecked, "X-No-Archive: Yes" & vbCrLf, "")

For i = 1 To SplitCount
    On Error GoTo BadPost
    NNTP1.ArticleText = rtMessageArea.Text & vbCrLf
    If Not tAttachedFile.Text = "" Then
        If SplitCount = 1 Then
            NNTP1.AttachedFile = EncodeAttachement(tAttachedFile.Text) 'EncodeAttachement(StripExt(tAttachedFile.Text))
        Else
            NNTP1.AttachedFile = EncodeAttachement(tAttachedFile.Text & "." & Format(i, "000"))
        End If
    End If
    NNTP1.Newsgroups = tNewsgroup.Text
    NNTP1.Subject = tSubject.Text & IIf(SplitCount = 1, "", " (" & StripFileName(NNTP1.AttachedFile) & ")" & " (1/" & CStr(i) & ")")
    NNTP1.From = tFrom.Text
    lblstatus = "Posting file " & i
    NNTP1.Action = a_PostArticle
    On Error Resume Next
    If iFileExists(NNTP1.AttachedFile) Then Kill NNTP1.AttachedFile
    DoEvents
Next
NNTP1.Action = a_Disconnect
Me.MousePointer = vbDefault
lblstatus = SplitCount & " files were posted successfully."
bPost.Enabled = True

Exit Sub
BadPost:
    lblstatus = ""
    lblsplitstatus.Visible = False
    ProgressBar1.Visible = False
    bPost.Enabled = True
    MsgBox "Posting failed with error: " & Err.Description, vbCritical + vbApplicationModal, "Post Error"
    On Error Resume Next
    NNTP1.Action = a_Disconnect
    Err.Clear
    Me.MousePointer = vbDefault
    If iFileExists(NNTP1.AttachedFile) Then Kill NNTP1.AttachedFile
End Sub

Private Sub btnBrowse_Click()
mAttachFile_Click
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub CheckSplitFile_Click(Value As Integer)
If Value = -1 Then
    tSegmentSize.Enabled = True
Else
    tSegmentSize.Enabled = False
End If
End Sub

Private Sub Form_Load()
Dim SectionName As String

tNewsServer = MailConnector.NNTPServerName
tFrom = MailConnector.EmailAddress
lblstatus = ""

SectionName = "NewsAccount"
tUserID = ReadProfile(SectionName, "NewsUserID")
tPassword = ReadProfile(SectionName, "NewsPassword")
tFrom = ReadProfile(SectionName, "NewsFrom")

End Sub

Private Sub Form_Resize()

On Error Resume Next

rtMessageArea.Height = ScaleHeight - rtMessageArea.Top - rtMessageArea.Left

End Sub

Private Function GetToken(Text$, Separators$) As String

If Text$ = "" Then
    GetToken = ""
    Exit Function
End If

Do While 0 <> InStr(Separators$, Left$(Text$, 1))
    Text$ = Mid$(Text$, 2)
Loop

If Text$ = "" Then
    GetToken = ""
    Exit Function
End If

Dim x%
For x% = 1 To Len(Separators$)
    Dim Separator$: Separator$ = Mid$(Separators$, x%, 1)
    Dim i%: i% = InStr(Text$, Separator$)
    If i% > 0 Then
        GetToken = Left$(Text$, i% - 1)
        Text$ = Mid$(Text$, i% + 1)
        Exit Function
    End If
Next x%

GetToken = Text$
Text$ = ""

End Function

Private Sub Form_Unload(Cancel As Integer)
Dim SectionName As String
'Dim KeyName As String

SectionName = "NewsAccount"
WriteProfile SectionName, "NewsUserID", tUserID
WriteProfile SectionName, "NewsPassword", tPassword
WriteProfile SectionName, "NewsFrom", tFrom


Set frmNewsReader = Nothing
End Sub

Private Sub mAttachFile_Click()
'Set frmEncodeFile.MessageArea = Me.rtMessageArea
'frmEncodeFile.FileName = tAttachedFile
'frmEncodeFile.Show vbModal
On Error Resume Next
CommonDialog1.DialogTitle = "Open file"
CommonDialog1.Flags = &H2& + &H4&
CommonDialog1.Filter = "Text Files *.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.CancelError = True
CommonDialog1.InitDir = App.Path
CommonDialog1.Action = 1
ChDir App.Path
tAttachedFile = CommonDialog1.FileName
tSegmentSize = FileLen(tAttachedFile)
tFileSize = tSegmentSize
If CLng(tSegmentSize) > 0 Then
    lblSegments = CLng(FileLen(tAttachedFile) / CLng(tSegmentSize))
Else
     lblSegments = FileLen(tAttachedFile)
End If
End Sub

Private Sub mDecodeArticle_Click()
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
If Len(rtMessageArea.Text) = 0 Then
    MsgBox "The is nothing in the message area.", vbExclamation + vbApplicationModal, "Empty Message Area"
    DoEvents
    Exit Sub
End If
'FileName = "temp"
'lblstatus = "Looking for beginning of file..."
'foo = InStr(1, rtMessageArea.Text, "name=")
'i = 0
'If Not foo = 0 Then
  ' foo = foo + Len("name=""")
   ' Do While i < 128
        'TextLine = Mid(rtMessageArea.Text, foo + i, 1)
        'If TextLine = vbCr Or TextLine = vbLf Or TextLine = """" Then Exit Do
        'FileName = FileName & TextLine
        'i = i + 1
   ' Loop
    'PIForm(gActivePIInstance).ShowStatus ("Found file: " & FileName)
    'DoEvents
 'End If
'If Not InStr(1, rtMessageArea, "base64") = 0 Then
 '   PIForm(gActivePIInstance).NetCode1.Format = f_BASE64
'Else
'    PIForm(gActivePIInstance).NetCode1.Format = f_UUEncode
'End If
    
PIForm(gActivePIInstance).NetCode1.MaxFileSize = 0
PIForm(gActivePIInstance).NetCode1.Overwrite = True
lblstatus = "Decoding file: " & FileName
DoEvents
PIForm(gActivePIInstance).NetCode1.FileName = App.Path & "\" & FileName
On Error GoTo ImportError
PIForm(gActivePIInstance).NetCode1.EncodedData = rtMessageArea.Text
PIForm(gActivePIInstance).NetCode1.Action = 3 'Decode to file
PIForm(gActivePIInstance).NetCode1.Action = 0

DoEvents
MousePointer = vbDefault

CommonDialog1.DialogTitle = "Save file"
CommonDialog1.FilterIndex = 1
If Not FileName = "" Then
    CommonDialog1.Filter = GetExt(FileName)
    CommonDialog1.FileName = App.Path & "\" & FileName
Else
    CommonDialog1.Filter = GetExt(PIForm(gActivePIInstance).NetCode1.FileName)
    CommonDialog1.FileName = PIForm(gActivePIInstance).NetCode1.FileName
End If
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave

FileNum = FreeFile
Open CommonDialog1.FileName & "." & CommonDialog1.Filter For Output As FileNum
Print #FileNum, PIForm(gActivePIInstance).NetCode1.DecodedData
Close FileNum
lblstatus = "Saved file at: " & CommonDialog1.FileName & "." & CommonDialog1.Filter
ChDir App.Path

Exit Sub
ImportError:
    Reset
    Beep
    MsgBox Err.Description, vbApplicationModal, App.Title
    MousePointer = vbDefault
    ChDir App.Path
    Err.Clear
End Sub



Private Sub mExit_Click()
Unload Me
End Sub


Private Sub mPostArticle_Click()
mPostArticle.Enabled = False
bPost_Click
mPostArticle.Enabled = True
End Sub

Private Sub NNTP1_EndTransfer()
lblstatus = "Transfer completed..."
End Sub

Private Sub NNTP1_Header(Field As String, Value As String)

Dim i%, c%

'Select Case UCase$(Field)
    
   ' Case "SUBJECT":
        
      '  rArticleSubject = Value
        
        'If cbSubjectCase = 1 Then c% = 0 Else c% = 1
        
        'For i% = 1 To rNrSubjectKeywords
          '  If 0 <> InStr(1, Value, rSubjectKeywords(i%), c%) Then
           '     rSearchWeight = rSearchWeight + 1
           ' End If
        'Next i%
    
    'Case "FROM":
        
      '  rArticleFrom = Value

      '  If cbFromCase = 1 Then c% = 0 Else c% = 1
        
       ' For i% = 1 To rNrFromKeywords
            'If 0 <> InStr(1, Value, rFromKeywords(i%), c%) Then
            '    rSearchWeight = rSearchWeight + 1
            'End If
        'Next i%

    'Case "DATE":
        
      '  rArticleDate = Value

'End Select

End Sub

Private Sub NNTP1_StartTransfer()

rArticleBody = ""
lblstatus = "Transfer started..."
End Sub

Private Sub NNTP1_Transfer(BytesTransferred As Long, Text As String)

Dim i%, c%

'If cbBodyCase = 1 Then c% = 0 Else c% = 1

'For i% = 1 To rNrBodyKeywords
    'If 0 <> InStr(1, Text, rBodyKeywords(i%), c%) Then
    '    rSearchWeight = rSearchWeight + 1
    'End If
'Next i%

'save up to 30K of article body for sending it out
'(if more of the article needs to be saved, then the contents
'maybe saved into a temporary file, and then emailed using
'the AttachedFile property of the SMTP control)
'If Len(rArticleBody) < 30000 Then
  '  rArticleBody = rArticleBody & Chr$(13) & Chr$(10) & Text
'End If
lblstatus = "Bytes transferred " & BytesTransferred
End Sub

Private Sub ParseKeywords()


End Sub

Private Sub ParseNewsgroupList()



End Sub


Public Function EncodeAttachement(sFileToEncode As String) As String
Dim EncodedData As String
Dim FileNum As Integer
Dim TextLine As String
Dim NumBytes As Long
Dim msg As String
Dim foo As Long
Dim FileSize As Long
Dim LineCount As Long

'House keeping
'On Error Resume Next
If tAttachedFile = "" Then
    EncodeAttachement = ""
    Exit Function
End If
'MkDir App.Path & "\temp"
'Err.Clear
'TempPathLocation & StripFileName(PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)) & ".asc"
           

On Error GoTo ImportError
lblstatus = "Encoding " & sFileToEncode & "..." 'tAttachedFile & "..."
'FileCopy tAttachedFile, App.Path & "\temp\tAttachedFile"

'On Error GoTo ImportError

'Select Case cmbEncodingType
    'Case "Base-64"
       ' PIForm(gActivePIInstance).NetCode1.Format = 1
        'PIForm(gActivePIInstance).nntp1.
    'Case "UU Encoded"
        PIForm(gActivePIInstance).NetCode1.Format = f_UUEncode '= '0
    'Case Else
       ' Beep
        'Exit Function
'End Select

PIForm(gActivePIInstance).NetCode1.MaxFileSize = 0
PIForm(gActivePIInstance).NetCode1.FileName = StripFileName(sFileToEncode)
PIForm(gActivePIInstance).NetCode1.Overwrite = True
PIForm(gActivePIInstance).NetCode1.DecodedData = sFileToEncode
'On Error Resume Next

PIForm(gActivePIInstance).NetCode1.EncodedData = TempPathLocation & StripFileName(sFileToEncode)
PIForm(gActivePIInstance).NetCode1.Action = 2
PIForm(gActivePIInstance).NetCode1.Action = 0
DoEvents
EncodeAttachement = PIForm(gActivePIInstance).NetCode1.EncodedData

Exit Function
ImportError:
    Err.Raise 2003, , Err.Description
End Function
    

Private Sub SSCheck1_Click(Value As Integer)

End Sub

Private Sub tAttachedFile_KeyPress(KeyAscii As Integer)
Beep
KeyAscii = 0
End Sub

Private Sub tSegmentSize_Change()
On Error Resume Next
If CheckSplitFile.Value = ssCBChecked Then
    lblSegments = 0
    Exit Sub
End If
If CLng(tSegmentSize) > 0 Then
    lblSegments = CLng(FileLen(tAttachedFile) / CLng(tSegmentSize))
Else
     lblSegments = FileLen(tAttachedFile)
End If
End Sub

