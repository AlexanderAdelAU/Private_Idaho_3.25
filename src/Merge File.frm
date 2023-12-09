VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMergeFile 
   Caption         =   "Merge a Split File"
   ClientHeight    =   4395
   ClientLeft      =   2400
   ClientTop       =   2100
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   7125
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   315
      Left            =   5940
      TabIndex        =   4
      Top             =   3900
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files to merge"
      Height          =   3105
      Left            =   150
      TabIndex        =   1
      Top             =   660
      Width           =   6795
      Begin VB.CommandButton cmdViewFile 
         Caption         =   "View File"
         Height          =   375
         Left            =   5760
         TabIndex        =   14
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Merge Files"
         Height          =   375
         Left            =   5760
         TabIndex        =   13
         Top             =   2220
         Width           =   975
      End
      Begin VB.Frame frmPercentComplete 
         Caption         =   "Percent Complete"
         Height          =   585
         Left            =   120
         TabIndex        =   7
         Top             =   2370
         Width           =   3165
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   225
            Left            =   120
            TabIndex        =   8
            Top             =   210
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   397
            _Version        =   327682
            Appearance      =   0
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   570
         Width           =   5235
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Status: "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "lblStatus"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "File size: "
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   2010
         Width           =   765
      End
      Begin VB.Label lblFileSize 
         BackStyle       =   0  'Transparent
         Caption         =   "lblFileSize"
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   2010
         Width           =   1695
      End
      Begin VB.Label lblMergeFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the FIRST file in the set that you wish to create a merge from, for example FileName.x.001"
         Height          =   555
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Segment file name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1425
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Private Idaho File Merge Utility."
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
      Index           =   1
      Left            =   1050
      TabIndex        =   5
      Top             =   90
      Width           =   4905
   End
End
Attribute VB_Name = "frmMergeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gszFileToLaunch As String
Private Sub cmdViewFile_Click()
Dim FileToLaunch As String
Dim res As Long
'
'Now try and launch it if we can
'
FileToLaunch = gszFileToLaunch
    DoEvents
    'gTemporaryFile = FileToLaunch
    res = ShellExecute(Me.hWnd, "open", FileToLaunch, vbNullString, CurDir, SW_SHOW)
    DoEvents
    If res < 32 Then
             Kill FileToLaunch 'TempPathLocation & lvwAttachments.SelectedItem.Text
             Err.Raise "Error was encountered launching the application associated with this attachment.   Please check your " & TempPathLocation & " directory to makes sure there are no plain text (decrypted) files there."
    End If
End Sub

Private Sub Command1_Click()
    Dim FNameNoExt As String
    Dim x As Integer
    Dim Segments As Integer
    'Dim s As String
    Dim J As Integer
    
    'Call the function
    lblStatus = ""
    frmPercentComplete.Visible = True
    lblStatus = MergeFiles(Text1, ProgressBar1, Segments)
   ' If Not lblstatus = "" Then Exit Sub
    J = InStrRev(Text1, ".", , vbTextCompare)
    If J = 0 Then   'File name does not contain a '.' character
        FNameNoExt = Text1
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(Text1, J - 1)
    End If
    lblMergeFileName(2) = "File saved as: " & FNameNoExt
    gszFileToLaunch = FNameNoExt
    frmPercentComplete.Visible = False
End Sub

Private Sub Command2_Click()
   ' Label3 = ""
    'Initialize the common dialog control and show it
    CommonDialog1.DialogTitle = "Select file segment to merge"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    Text1 = CommonDialog1.FileName
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Initialize
    lblStatus.Caption = ""
    lblFileSize.Caption = "0"
    Text1 = ""
    frmPercentComplete.Visible = False
End Sub
