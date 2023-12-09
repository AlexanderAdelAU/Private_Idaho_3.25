VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFindPGPKeys 
   Caption         =   "Find PGPKeys Utility"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6630
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fNymKeys 
      Caption         =   "Find PGPTools Directory"
      Height          =   2955
      Left            =   330
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton BtnOK 
         Caption         =   "OK"
         Height          =   345
         Left            =   4350
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   5460
         TabIndex        =   5
         Top             =   2400
         Width           =   945
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "Browse"
         Height          =   315
         Left            =   5460
         TabIndex        =   4
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton bKeys 
         Caption         =   "Search For PGPKeys.exe"
         Height          =   375
         Left            =   210
         TabIndex        =   2
         Top             =   2370
         Width           =   2175
      End
      Begin VB.TextBox txtFilePath 
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Text            =   "C:\Program Files\Network Associates"
         Top             =   1560
         Width           =   5085
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   $"Find PGPKeys.frx":0000
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   1
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   5685
      End
   End
End
Attribute VB_Name = "frmFindPGPKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bKeys_Click()
Dim FileFunctions As New cFileFunctions
Dim objFiles As New Collection
Dim objFile As File
Dim iResponse As Long
Dim SectionName As String
Dim sPath As String

SectionName = "PGP Options"
On Error GoTo bKeysError
bKeys.Enabled = False
'sPath = ReadProfile(SectionName, "PGPKeys.exe Location")
'If sPath = "" Then
    'Search for the file
    Me.MousePointer = vbHourglass
    DoEvents
    Set objFiles = FileFunctions.FindFile("PGPKeys.exe", "C:\Program Files")
    Me.MousePointer = vbDefault
    DoEvents
    'Add the results to the listbox
    For Each objFile In objFiles
        If UCase(objFile.name) = UCase("PGPKeys.exe") Then
            sPath = objFile.Path
            'iResponse = ShellExecute(Me.hWnd, "open", sPath, vbNullString, CurDir, SW_SHOW)
           ' If iResponse < 30 Then
            '    Err.Raise 6006, "Error executing PGPKeys.exe", iResponse
            'Else
                WriteProfile SectionName, "PGPKeys.exe Location", objFile.Path
            'End If
            Unload Me
        End If
    Next
'Else
     'Me.MousePointer = vbHourglass
     'iResponse = ShellExecute(Me.hWnd, "open", sPath, vbNullString, CurDir, SW_SHOW)
     'If iResponse < 30 Then
     ''           Err.Raise 6006, "Error executing PGPKeys.exe", iResponse
    ' End If
     'Me.MousePointer = vbDefault
Set objFile = Nothing
Set objFiles = Nothing
Set FileFunctions = Nothing
bKeys.Enabled = True
Exit Sub
bKeysError:
     Me.MousePointer = vbDefault
    'bKeys.Enabled = True
    WriteProfile SectionName, "PGPKeys.exe Location", ""
    MsgBox "Error detected while trying to find or execute the PGPKeys.exe file.  Error reported as: " & Err.Description, vbApplicationModal + vbCritical, "File Location Error"
    Err.Clear
End Sub

Private Sub btnBrowse_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Find PGPKeys.exe"
CommonDialog1.Flags = &H2& + &H4&
CommonDialog1.Filter = "PGP Files (PGP*.exe)|PGP*.exe"
'"Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.CancelError = True
CommonDialog1.InitDir = txtFilePath
CommonDialog1.Action = 1
txtFilePath = CommonDialog1.FileName
ChDir App.Path
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnok_Click()

If Not txtFilePath = "" Then
    WriteProfile "PGP Options", "PGPKeys.exe Location", txtFilePath
End If
Unload Me
End Sub

Private Sub Form_Load()

    txtFilePath = ReadProfile("PGP Options", "PGPKeys.exe Location")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFindPGPKeys = Nothing
End Sub
