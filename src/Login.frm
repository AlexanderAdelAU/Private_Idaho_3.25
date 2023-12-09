VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Private Idaho Login"
   ClientHeight    =   2820
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1666.148
   ScaleMode       =   0  'User
   ScaleWidth      =   4309.762
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   735
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1335
      TabIndex        =   4
      Top             =   1980
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2580
      TabIndex        =   5
      Top             =   1980
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1125
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   210
      Picture         =   "Login.frx":0000
      Top             =   60
      Width           =   570
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't forget your password!"
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "You need to login to Private Idaho. "
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   2
      Left            =   1350
      TabIndex        =   6
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Verify Password"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
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
            MsgBox "Passwords did not verify..!", , "Login"
        End If
    End If
    rs.MoveFirst
    If txtPassword(0) = rs("password") Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        Me.Hide
        LoginSucceeded = True
        frmMain.Show
        rs.Close
        Set rs = Nothing

    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword(0).SetFocus
        SendKeys "{Home}+{End}"
    End If
    
End Sub

Private Sub Form_Load()
Dim msg As String
Dim rs As Recordset

Set rs = DB.OpenRecordset("Users", dbOpenDynaset)

If rs.EOF Then
    Unload frmSplash
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
