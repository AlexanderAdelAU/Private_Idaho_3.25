VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "Threed20.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFileSplitter 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "File Splitter Utility"
   ClientHeight    =   3930
   ClientLeft      =   1185
   ClientTop       =   1770
   ClientWidth     =   8145
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
   ScaleHeight     =   3930
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Files to Split"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   7785
      Begin VB.CommandButton bBrowse 
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
         Left            =   6510
         TabIndex        =   13
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox tAttachedFile 
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
         Left            =   1260
         TabIndex        =   7
         Tag             =   "File to be encoded"
         Top             =   390
         Width           =   5115
      End
      Begin VB.TextBox tSegmentSize 
         Alignment       =   1  'Right Justify
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
         Left            =   3270
         TabIndex        =   6
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "bytes"
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
         Index           =   3
         Left            =   3240
         TabIndex        =   17
         Top             =   780
         Width           =   765
      End
      Begin VB.Label tFileSize 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1260
         TabIndex        =   16
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File size:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   780
         Width           =   1065
      End
      Begin Threed.SSCommand bSplit 
         Height          =   375
         Left            =   6510
         TabIndex        =   14
         Top             =   1290
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   131074
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Split File"
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3270
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File to split:"
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
         Left            =   330
         TabIndex        =   11
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the file size for each segment here: "
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
         Height          =   345
         Index           =   9
         Left            =   300
         TabIndex        =   10
         Top             =   1200
         Width           =   3045
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Number of segments to be created:"
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
         Left            =   60
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "bytes"
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
         Index           =   0
         Left            =   4740
         TabIndex        =   8
         Top             =   1200
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   2
      Top             =   3000
      Width           =   5235
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.CommandButton bexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   1
      Top             =   3120
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "lblstatus"
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
      Left            =   210
      TabIndex        =   4
      Top             =   2670
      Width           =   5235
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Private Idaho File Splitter Utility."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   1800
      TabIndex        =   0
      Top             =   150
      Width           =   4905
   End
   Begin VB.Menu mFIle 
      Caption         =   "File"
      Begin VB.Menu mFileToSplit 
         Caption         =   "Open File to Split"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmFileSplitter"
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



Private Sub bBrowse_Click()
lblstatus = ""
tSegmentSize = ""
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

Dim objFiles As New Collection
Dim objFile As File
Dim FileFunctions As New cFileFunctions

    Set objFiles = FileFunctions.FindFile(StripFileName(tAttachedFile), StripPath(tAttachedFile))
Set objFile = objFiles(1)
tSegmentSize = Format(objFile.Size, "###,###")
tFileSize = Format(objFile.Size, "###,###")

If CLng(tSegmentSize) > 0 Then
    lblSegments = CLng(tFileSize) / CLng(tSegmentSize)
Else
     lblSegments = tFileSize
End If
End Sub

Private Sub bexit_Click()
Unload Me
End Sub

Private Sub bSplit_Click()
Dim SectionName As String
Dim i As Integer
Dim Segments As Integer
Dim SplitStatus As Integer
'Dim FileLocation(1 To 1000)   As String
Dim SplitCount As Long


'
' Okay split the file if necessary
'
On Error GoTo BadPost
lblstatus = ""
DoEvents
If CDbl(tSegmentSize) > 0 Then
    SplitCount = CDbl(tFileSize) / CDbl(tSegmentSize)
    If SplitCount > 1000 Then
        MsgBox "Number of file segments that will result from the file size you have used will be too large.  The maximum number of files this program can generate is 1000.", vbCritical + vbApplicationModal, "File Split Error."
        lblstatus = "Post unsuccessful."
        lblstatus = ""
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    ProgressBar1.Visible = True
    SplitStatus = SplitFile(tAttachedFile, CDbl(tFileSize), Val(tSegmentSize), ProgressBar1, Segments)
    ProgressBar1.Visible = False
    If Not SplitStatus = 0 Then
        MsgBox "Error splitting the file.", vbCritical + vbApplicationModal, "File Split Error."
        lblstatus = "Post unsuccessful."
        lblstatus = ""
        Me.MousePointer = vbDefault
        Exit Sub
    End If
Else
    SplitCount = 1
End If


Me.MousePointer = vbDefault
lblstatus = SplitCount & " files were saved successfully."
DoEvents

Exit Sub
BadPost:
    lblstatus = ""
    ProgressBar1.Visible = False
    MsgBox "Posting failed with error: " & Err.Description, vbCritical + vbApplicationModal, "Post Error"
    On Error Resume Next
     Err.Clear
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
lblstatus = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmNewsReader = Nothing
End Sub




Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mFileToSplit_Click()
Call bBrowse_Click

End Sub

Private Sub tSegmentSize_Change()

On Error Resume Next
If CDbl(tSegmentSize) > 0 Then
    lblSegments = Format(CDbl(tFileSize) / CDbl(tSegmentSize), "###,###")
Else
     lblSegments = tFileSize
End If
'tSegmentSize = Format(tSegmentSize, "###,###")

End Sub
