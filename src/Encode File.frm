VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEncodeFile 
   Caption         =   "Encode a File"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Left            =   3840
      TabIndex        =   5
      Top             =   2280
      Width           =   1035
   End
   Begin VB.ComboBox cmbEncodingType 
      Height          =   315
      ItemData        =   "Encode File.frx":0000
      Left            =   90
      List            =   "Encode File.frx":000A
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
   Begin VB.Label lblstatus 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   4935
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
Attribute VB_Name = "frmEncodeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MessageArea As RichTextBox
Private Sub btnBrowse_Click()
    
On Error Resume Next
CommonDialog1.DialogTitle = "Open file"
CommonDialog1.Flags = &H2& + &H4&
CommonDialog1.Filter = "Text Files *.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.CancelError = True
CommonDialog1.InitDir = App.Path
CommonDialog1.Action = 1
ChDir App.Path
txtFileName = CommonDialog1.FileName

End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnok_Click()
Dim EncodedData As String
Dim FileNum As Integer
Dim TextLine As String
Dim NumBytes As Long
Dim msg As String
Dim foo As Long
Dim FileSize As Long
Dim LineCount As Long
On Error GoTo ImportError

MousePointer = vbHourglass
lblstatus = "Encoding " & txtFileName & "..."

On Error Resume Next
Select Case cmbEncodingType
    Case "Base-64"
        PIForm(gActivePIInstance).NetCode1.Format = 1
        'PIForm(gActivePIInstance).nntp1.
    Case "UU Encoded"
        PIForm(gActivePIInstance).NetCode1.Format = 0
    Case Else
        Beep
        Exit Sub
End Select

PIForm(gActivePIInstance).NetCode1.MaxFileSize = 0
PIForm(gActivePIInstance).NetCode1.FileName = CommonDialog1.FileTitle
PIForm(gActivePIInstance).NetCode1.Overwrite = True
PIForm(gActivePIInstance).NetCode1.DecodedData = txtFileName
'On Error Resume Next
Kill App.Path & "\pi00.b64"

PIForm(gActivePIInstance).NetCode1.EncodedData = App.Path & "\pi00.b64"
PIForm(gActivePIInstance).NetCode1.Action = 2
PIForm(gActivePIInstance).NetCode1.Action = 0
EncodedData = GetFileText(PIForm(gActivePIInstance).NetCode1.EncodedData)
Kill App.Path & "\pi00.b64"
MessageArea.SelStart = Len(MessageArea)
MessageArea.SelText = vbCrLf
MessageArea.SelText = EncodedData
DoEvents
MousePointer = vbDefault
Unload Me
Exit Sub
ImportError:
    Reset
    MsgBox Err.Description & " in File Import", vbApplicationModal, App.Title
    Err.Clear
    MousePointer = vbDefault
    Kill App.Path & "\pi00.b64"
End Sub

Private Sub Form_Load()
txtFileName = ""
cmbEncodingType.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmEncodeFile = Nothing
End Sub
