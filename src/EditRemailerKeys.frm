VERSION 5.00
Begin VB.Form frmAddRemailer 
   Caption         =   "Add Remailer"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "You have an Option of:"
      Height          =   1815
      Left            =   360
      TabIndex        =   4
      Top             =   180
      Width           =   5475
      Begin VB.OptionButton Option1 
         Caption         =   $"EditRemailerKeys.frx":0000
         Height          =   1035
         Index           =   1
         Left            =   420
         TabIndex        =   6
         Top             =   660
         Width           =   4275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Creating a set of New Keys for this remailer, or"
         Height          =   435
         Index           =   0
         Left            =   420
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   4275
      End
   End
   Begin VB.TextBox txtRemailerName 
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   3390
      Width           =   5685
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Top             =   4590
      Width           =   1155
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   4230
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   $"EditRemailerKeys.frx":011A
      Height          =   1155
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   5835
   End
End
Attribute VB_Name = "frmAddRemailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOkay_Click()
Dim cmd As String
Dim ClipText As String
Dim FileNum As Integer
Dim AppendString As String
Dim OutFileName As String

On Error GoTo BadKeys
    If Len(txtRemailerName.Text) = 0 Then
            MsgBox "I can't proceed without remailer name.  Please enter a remailer name if you wish to proceed..", vbCritical + vbApplicationModal, App.Title
            Exit Sub
    End If
    If Option1(1).Value = True Then
      '  CreatePGPKeyPair
    Else
        If Len(txtRemailerName) = 0 Then
            MsgBox "I can't proceed without a public Key.  Please ensure the message area in PI has a key in it.", vbCritical + vbApplicationModal, App.Title
            Exit Sub
        End If
        
        'Okay lets add the public key
          
        ClipText = txtRemailerName
        FileNum = FreeFile
        If InStr(1, gPGPFile, ":") = 0 Then
            OutFileName = gPGPPath & "\" & gPGPFile & ".out"
            Open OutFileName For Output As FileNum
        Else
             OutFileName = gPGPFile & ".out"
            Open OutFileName For Output As FileNum
        End If
        Print #FileNum, ClipText
        Close #FileNum
        cmd = App.Path + "\" & gPIPIF & " -ka " & OutFileName
        
        CheckLen (cmd)
        ExecCmd (cmd)
        On Error Resume Next
        Kill OutFileName
    End If
    
    On Error GoTo BadKeys
   ' 'UpdatePublicKeysFile
    
    'Now load into Private.txt file
    FileNum = FreeFile
    If iFileExists(App.Path & "\PRIVATE.TXT") Then
        Open App.Path & "\PRIVATE.TXT" For Append As FileNum
    Else
        Open App.Path & "\PRIVATE.TXT" For Output As FileNum
    End If
   ' AppendString = ""
   Dim StartSpace As Integer
   Dim StartDelim As Integer
   Dim EndDelim As Integer
   Dim FullName As String
   
   StartSpace = InStr(1, txtRemailerName, " ", vbTextCompare)
   If Not StartSpace = 0 Then
        FullName = Trim(Mid(txtRemailerName, 1, StartSpace - 1))
        txtRemailerName = Trim(Mid(txtRemailerName, StartSpace + 1))
   Else
        MsgBox "Are you sure you have entered the emial address correctly?", vbCritical, App.Title
        Exit Sub
   End If
    StartDelim = InStr(1, txtRemailerName, "<", vbTextCompare)
    If StartDelim <> 0 Then
        txtRemailerName = Mid(txtRemailerName, StartDelim + 1)
    End If
    EndDelim = InStr(1, txtRemailerName, ">", vbTextCompare)
    If EndDelim <> 0 Then
        RemailerContext.name = Mid(txtRemailerName, 1, EndDelim - 1)
    End If
    AppendString = "$remailer{" & """" & FullName & """" & "} = "
    AppendString = AppendString & """<" & txtRemailerName & "> cpunk pgp pgponly""" & ";" & vbCrLf
     AppendString = AppendString & FullName & " " & txtRemailerName & "       ########   51:00   30.50%"
    Print #FileNum, AppendString
    Close #FileNum
    
    'Now reinitialise remailers
    DoEvents
    'using this to up the number of remailers is not good - fix later
    'gTotalMatchedRemailers = gTotalMatchedRemailers + 1
    gTotalRemailers = 0
    frmRemailerList.InitializeRemailers (App.Path + "\remailer.htm")
    frmRemailerList.InitializeRemailers (App.Path + "\private.txt")
    frmRemailerList.SortRemailers
    frmRemailerList.FillRemailerList
    
    Unload Me
    Exit Sub
BadKeys:
    Reset
    MsgBox Err.Description & " in Newkeys)", vbCritical + vbApplicationModal, App.Title
    Err.Clear
End Sub

Private Sub Form_Load()
Dim Win As New CWindow
    
Win.Center Me, Null
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAddRemailer = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
'If Option1(1).Value = True Then
   ' txtRemailerName.Enabled = True
'Else
  '  txtRemailerName.Enabled = False
'End If
    
End Sub

