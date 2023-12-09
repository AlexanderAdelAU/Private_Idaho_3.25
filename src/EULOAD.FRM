VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4185
   ClientLeft      =   3285
   ClientTop       =   1800
   ClientWidth     =   6075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkTopic       =   "frmSplash"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4635
      TabIndex        =   3
      Top             =   3600
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2100
      TabIndex        =   2
      Top             =   3600
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2355
      Width           =   3645
   End
   Begin VB.Timer Timer1 
      Left            =   210
      Top             =   1530
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Supports PGP Version 5 and 6.  Does not support PGP Version 7 yet."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   810
      TabIndex        =   9
      Top             =   1800
      Width           =   2625
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4410
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   870
      X2              =   5250
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version for NT/W2k/W9x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   810
      TabIndex        =   7
      Top             =   1320
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   480
      Picture         =   "EULOAD.frx":0000
      Top             =   60
      Width           =   5250
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Verify Password"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   1
      Left            =   660
      TabIndex        =   6
      Top             =   2820
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   690
      TabIndex        =   5
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't forget your password!"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2115
      TabIndex        =   4
      Top             =   3210
      Width           =   3300
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public LoginSucceeded As Boolean
Private FirstTime As Boolean
Private Sub cmdCancel_Click()
  'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    End
End Sub

Private Sub cmdOK_Click()
Dim rs As Recordset
    'check for correct password
    On Error Resume Next
    Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
    If Not rs.EOF Then rs.MoveFirst
    If FirstTime Then
        If txtPassword(0) = "" Then Exit Sub
        If txtPassword(0) = txtPassword(1) Then
             rs.AddNew
             rs("password") = txtPassword(0)
             rs("Expired") = False 'Initialise this
             rs.Update
             rs.Close
        Else
            lblPrompt = "Passwords did not verify..!"
        End If
    End If
   ' rs.MoveFirst
   
     If txtPassword(0) = rs("password") Then
        rs.Close
        Set rs = Nothing
        txtPassword(0).PasswordChar = ""
        txtPassword(0) = "checking password..."
        DoEvents
       ' Load frmMain
        Me.Hide
        Unload Me
        frmMain.Show
       ' DoEvents
        'Unload Me
        'MsgBox "Debug 3", vbApplicationModal

    Else
        lblPrompt = "Invalid password"
        txtPassword(0) = ""
        txtPassword(0).SetFocus
        'SendKeys "{Home}+{End}"
    End If
    
End Sub



Private Sub Form_Load()
Dim Win As New CWindow
Dim App As New CApplication
    Win.Center Me, Null
    Win.OnTop(Me) = True
    lblVersion = App.Version
    
    
   '     On Error GoTo Er_Hndl
   ' Gvbox1.FileName = App.Path & "\pi version 5 splash screen.gif"
  '  If Gvbox1.FileName <> "" Then
   '     Gvbox1.Visible = True
   '     Caption = "Private Idaho Version" & App.Version
   ' End If
    DoEvents
    Set Win = Nothing
    Set App = Nothing
    login
    txtPassword(0) = ""
    Load frmMain
    Exit Sub

Er_Hndl:
    Beep
    MsgBox Error$, 48, "Graphics Viewer"
    'Resume
    Exit Sub
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSplash = Nothing
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Public Sub login()
Dim msg As String
Dim rs As Recordset

Set rs = DB.OpenRecordset("Users", dbOpenDynaset)

If rs.EOF Then
   ' Unload frmSplash
    DoEvents
    FirstTime = True
    lblLabels(1).Visible = True
    txtPassword(1).Visible = True
    
    msg = "You need to enter a password to access your mail - as this is the first time you have been prompted for your login password, type in your password into the text box, and verify it by typing it in again." & vbCrLf & vbCrLf
    msg = msg & "Don't forget this password otherwise you will not be able to get into Private Idaho in future.  Encrypting the database secures your Nym IDs and other personal information." & vbCrLf & vbCrLf
    msg = msg & "Please be aware that any messages in plain text will still be accessible because the actual message content is not stored in the database - only references to where the files are located are stored in the database.  So if you want your information to be protected, don't leave plain text lying around, ie view the message and don't save it to disk."
   
    MsgBox msg, vbExclamation + vbApplicationModal, "Password Required"
    'Exit Sub
End If
rs.Close
End Sub
