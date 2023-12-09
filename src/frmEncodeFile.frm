VERSION 5.00
Begin VB.Form EncodeFile 
   Caption         =   "Encode a File"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4980
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3810
      TabIndex        =   5
      Top             =   2280
      Width           =   1035
   End
   Begin VB.ComboBox cmbEncodingType 
      Height          =   315
      ItemData        =   "frmEncodeFile.frx":0000
      Left            =   90
      List            =   "frmEncodeFile.frx":000A
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1470
      Width           =   2445
   End
   Begin VB.TextBox txtFileName 
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   690
      Width           =   4965
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Browse"
      Height          =   405
      Left            =   5160
      TabIndex        =   1
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Endcoding Type"
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   1230
      Width           =   4905
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the name and path of the file you wish to encode."
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   4905
   End
End
Attribute VB_Name = "EncodeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBrowse_Click()
    
CommonDialog1.DialogTitle = "Open file"
CommonDialog1.Flags = &H2& + &H4&
CommonDialog1.Filter = "Text Files *.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.CancelError = True
CommonDialog1.InitDir = gPathName
CommonDialog1.Action = 1
ChDir App.Path


End Sub

Private Sub btnOk_Click()
Dim FileNum As Integer
Dim TextLine As String
Dim NumBytes As Long
Dim msg As String
Dim foo As Long
Dim FileSize As Long
Dim LineCount As Long
On Error GoTo ImportError
FileNum = FreeFile
Open CommonDialog1.FileName For Input As FileNum
If Len(MessageArea.Text) > 0 Then
    foo = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Send Feedback")
    If foo = vbNo Then
        Exit Sub
    End If
End If
frmMain.MessageArea.SelStart = 0
frmMain.MessageArea = ""
msg = ""
MousePointer = vbHourglass
frmMain.NNTP1.EncodedData = CommonDialog1.FileName
frmMain.lblstatus = "Encloding file " & CommonDialog1.FileName & "..."
DoEvents
'(CommonDialog1.FileName)
        
frmMain.MessageArea = msg
MousePointer = vbDefault
ImportError:
    Close FileNum
    MsgBox Err.Description & " in File Import", vbApplicationModal, App.Title
    Err.Clear
End Sub

