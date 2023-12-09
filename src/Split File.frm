VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplitFile 
   Caption         =   "Split a file"
   ClientHeight    =   4005
   ClientLeft      =   2400
   ClientTop       =   2100
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6975
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4170
      TabIndex        =   9
      Top             =   2580
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   765
      Left            =   150
      TabIndex        =   5
      Top             =   2940
      Width           =   5055
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File to split"
      Height          =   2235
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6165
      Begin VB.TextBox txtNumSegments 
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Text            =   "2"
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   5100
         TabIndex        =   11
         Top             =   330
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Split File"
         Height          =   405
         Left            =   1230
         TabIndex        =   10
         Top             =   1680
         Width           =   2625
      End
      Begin VB.TextBox txtSplitAtByte 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   "0"
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1350
         TabIndex        =   0
         Top             =   330
         Width           =   3705
      End
      Begin VB.Label Label7 
         Caption         =   "File Size (Bytes)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblFileLength 
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Number of segments"
         Height          =   495
         Left            =   3600
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Split at byte:"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Split at byte:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "File name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label3 
      Caption         =   "Error Code:"
      Height          =   255
      Left            =   210
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "frmSplitFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim x As Integer
    Dim Segments As Integer
    
    'Call the function
    ProgressBar1 = 0
    x = SplitFile(txtFileName, dgFileSize, CDbl(txtSplitAtByte), ProgressBar1, txtNumSegments)
 
    'Inform the user about the call success or failure
    Label4 = x
    If x = 0 Then
        MsgBox "The process completed successfully." & Chr(10) & "The file was split to " & Segments & " segments.", vbInformation
    Else
        MsgBox "An error occured!", vbInformation
    End If

End Sub

Private Sub Command2_Click()

Dim lFileSize As Double

Dim Temp As Double
Dim lpFileSize As Long
Const OPEN_EXISTING = 3
Const FILE_SHARE_READ = &H1
Const GENERIC_READ = &H80000000
Dim fHandle As Long
Dim mLong As Double

    'Initialize the common dialog control and show it
    CommonDialog1.DialogTitle = "Select file to split"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
  '  fHandle = CreateFile(CommonDialog1.FileName, GENERIC_READ, FILE_SHARE_READ, _
                    ByVal 0&, OPEN_EXISTING, 0, 0)
    txtFileName = CommonDialog1.FileName
'Call CloseHandle(fHandle)

   ' lFileSize = GetFileSize(fHandle, lpFileSize)
   If FileLen(CommonDialog1.FileName) < 0 Then
        mLong = 2147483648#
        Temp = mLong + Val(FileLen(CommonDialog1.FileName))
        dgFileSize = mLong + Temp
   Else
     dgFileSize = Val(FileLen(CommonDialog1.FileName))
    End If
    lblFileLength = dgFileSize
    
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Initialize
   ' Label4.Caption = ""
   'Text1 = ""
    'Text2 = ""
    Dim x As Integer
    Dim Segments As Double
   ProgressBar1.Value = 0
    'Call the function
   ' x = SplitFile(txtFileName, Val(txtSplitAtByte), txtNumSegments, ProgressBar1)
 
    'Inform the user about the call success or failure
   ' Label4 = x
   ' ProgressBar1.Value = x
    'If x = 0 Then
        'MsgBox "The process completed successfully." & Chr(10) & "The file was split to " & Segments & " segments.", vbInformation
    'Else
     '   MsgBox "An error occured!", vbInformation
   ' End If

End Sub

