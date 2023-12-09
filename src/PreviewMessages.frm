VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmPreviewMessages 
   Caption         =   "Preview Messages"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMailBox 
      Height          =   2535
      Left            =   420
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   6555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"PreviewMessages.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   420
      TabIndex        =   9
      Top             =   60
      Width           =   6585
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   4
      Left            =   6450
      TabIndex        =   8
      Top             =   750
      Width           =   885
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   7
      Top             =   750
      Width           =   885
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   2
      Left            =   3780
      TabIndex        =   6
      Top             =   750
      Width           =   885
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   1
      Left            =   2430
      TabIndex        =   5
      Top             =   750
      Width           =   885
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   1050
      TabIndex        =   4
      Top             =   720
      Width           =   885
   End
   Begin Threed.SSCommand btnCancel 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   3780
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   873
      _Version        =   131074
      Caption         =   "Cancel"
   End
   Begin Threed.SSCommand btnDownload 
      Height          =   495
      Left            =   2820
      TabIndex        =   2
      Top             =   3780
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   873
      _Version        =   131074
      ForeColor       =   32768
      Caption         =   "Download Messages"
   End
   Begin Threed.SSCommand btnDelete 
      Height          =   495
      Left            =   390
      TabIndex        =   1
      Top             =   3780
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   131074
      ForeColor       =   255
      Caption         =   "Delete Selected Messages"
   End
End
Attribute VB_Name = "frmPreviewMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gszNumMessages As Integer

Private Function ScanHeaders()
'search - a string to search for within the header
    'ScanTopHeaders - returns a string containing found message numbers delimited
    'by spaces
    'Dim NumMessages As Integer
    'Dim tmpstr As String
    Dim msgRecord As String
    Dim J As Integer
    
    'NumMessages = 0
    'tmpstr = ""
    For J = 1 To gszNumMessages
        'DoEvents
        msgRecord = ""
       ' If search = "" Then
            'This is a simple fix to get all messages
         '   NumMessages = NumMessages + 1
         '   tmpstr = tmpstr & Format(J) & " "
       ' Else
           frmMain.GetPOPTop (J)
            'From
             msgRecord = gMessageRecord.From
            'To
            msgRecord = msgRecord & Chr(9) & gMessageRecord.To
            'Subject
            msgRecord = msgRecord & Chr(9) & gMessageRecord.Subject
            'Message size
            msgRecord = msgRecord & Chr(9) & gMessageRecord.MessageSize
            'Attachment
            msgRecord = msgRecord & Chr(9) & IIf(InStr(1, gMessageRecord.Header, "boundary=") = 0, "Yes", "No")
            lstMailBox.AddItem msgRecord
    Next
End Function

Private Sub btnCancel_Click()
gFoundMessages = ""
Unload Me
End Sub

Private Sub btnDelete_Click()
Dim i As Integer
Dim s As String
'Dim gFoundMessages2 As Integer
'Dim gMessagesToBeDeleted As String
's = gFoundMessages
For i = 0 To gszNumMessages - 1
    If lstMailBox.Selected(i) Then
        If gMessagesToBeDeleted = "" Then
            gMessagesToBeDeleted = CStr(i + 1)
        Else
            gMessagesToBeDeleted = gMessagesToBeDeleted & " " & CStr(i + 1)
        End If
    End If
Next
Unload Me
End Sub

Private Sub btnDownload_Click()
gMessagesToBeDeleted = ""
Unload Me
End Sub

Private Sub Form_Load()
lblHeading(0) = "From"
lblHeading(1) = "To"
lblHeading(2) = "Subject"
lblHeading(3) = "Size"
lblHeading(4) = "Attachment"
ScanHeaders
End Sub

Private Sub Form_Resize()
'ButtonClearance = Me.height - Button.top + button height = 720
'listbox bottom clearance = button.top - listox.height = 3330-2535=895
 On Error Resume Next
 If WindowState <> 1 Then
    btnDelete.Top = Me.Height - btnDelete.Height - 720
    btnDownload.Top = Me.Height - btnDownload.Height - 720
    btnCancel.Top = Me.Height - btnCancel.Height - 720
    lstMailBox.Height = btnDelete.Top - 895
    lstMailBox.Width = Me.Width - 895
    'If Me.Width < btnCancel.Left + btnCancel.Width + 995 Then
       ' Me.Width = btnCancel.Left + btnCancel.Width + 995
   ' End If
    
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPreviewMessages = Nothing
End Sub

