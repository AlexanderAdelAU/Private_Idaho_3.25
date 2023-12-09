VERSION 5.00
Begin VB.Form frmFingerCommand 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send Finger Commands"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHostAddress 
      Height          =   345
      Left            =   330
      TabIndex        =   3
      Text            =   "anon.lcs.mit.edu"
      Top             =   390
      Width           =   5505
   End
   Begin VB.CommandButton btncmd 
      Caption         =   "Exit"
      Height          =   375
      Index           =   1
      Left            =   4740
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton btncmd 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3450
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtCommand 
      Height          =   645
      Left            =   330
      TabIndex        =   0
      Text            =   "groups+alt.*(security|privacy)@anon.lcs.mit.edu"
      Top             =   1080
      Width           =   5475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Finger Command"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   870
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Host Address"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   150
      Width           =   2235
   End
End
Attribute VB_Name = "frmFingerCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncmd_Click(Index As Integer)
Dim rUser$, rHost$

On Error GoTo Fingererror
If Index = 0 Then
       
Me.MousePointer = vbHourglass
PIForm(gActivePIInstance).MousePointer = vbHourglass
PIForm(gActivePIInstance).IPPort1.WinsockLoaded = True

rHost$ = txtHostAddress
'rUser$ = "mail2news"
rUser$ = txtCommand
'rUser$ = "groups+.*scientology"
'rUser$ = "groups+alt.*(security|privacy)"
'rUser$ = "complaints"
'rUser$ = "finger"
'rUser$ = "mix"
'rUser$ = "mix-info"
'rUser$ = "mix-admin"
'rUser$ = "mail2news"
'rUser$ = "groups"
'rUser$ = "Nym.ID"
'rUser$ = "nymhelp"
'rUser$ = "nymhelp-html"
'rUser$ = "premail-info"

'If tHost = "" Then Beep: Exit Sub

'Raph's List
'rHost$ = "kiwi.cs.berkeley.edu"
'rUser$ = "remailer-list"

'PIForm(gActivePIInstance).MessageArea.SelFontName = "arial"
'PIForm(gActivePIInstance).MessageArea.SelBold = False
'PIForm(gActivePIInstance).MessageArea.SelItalic = False
'PIForm(gActivePIInstance).MessageArea.SelStrikeThru = False
'PIForm(gActivePIInstance).MessageArea.SelFontSize = 8.25

'Matt Ghio's List
'rUser$ = "remailer.help.all"
'rHost$ = "chaos.taylored.com"

'Matt's pinging service
'rUser$ = "remailer-list"
'rHost$ = "chaos.taylored.com"

'rHost$ = "204.95.228.28"


PIForm(gActivePIInstance).ShowStatus 1, ""

PIForm(gActivePIInstance).IPPort1.EOL = vbCrLf

'close old connections (if any)
If PIForm(gActivePIInstance).IPPort1.Connected Then PIForm(gActivePIInstance).IPPort1.Connected = False

'x% = InStr(tHost, "@")

'If x% <> 0 Then
'    rUser$ = Left$(tHost, x% - 1)
'    rHost$ = Mid$(tHost, x% + 1)
'Else
'    rUser$ = ""
'    rHost$ = tHost
'End If

PIForm(gActivePIInstance).IPPort1.RemoteHost = rHost
PIForm(gActivePIInstance).IPPort1.RemotePort = 79 'finger service

DoEvents
'attempt connection
PIForm(gActivePIInstance).ShowStatus 1, "Connecting to " & txtHostAddress & "...   "
PIForm(gActivePIInstance).IPPort1.WinsockLoaded = True
PIForm(gActivePIInstance).IPPort1.Connected = True

'wait until the connection is achieved
'(timeout in 10 seconds)
After10Seconds = Now + 10# / (3600# * 24#)
Do Until Now > After10Seconds
    If PIForm(gActivePIInstance).IPPort1.Connected Then
        PIForm(gActivePIInstance).ShowStatus 1, "Connected to " & PIForm(gActivePIInstance).IPPort1.RemoteHost & "...   "
        Exit Do
    End If
    DoEvents
Loop
If Not PIForm(gActivePIInstance).IPPort1.Connected Then
    MsgBox "Connection ended."
    PIForm(gActivePIInstance).ShowStatus 1, ""
    Exit Sub
End If

'send the data
PIForm(gActivePIInstance).IPPort1.DataToSend = rUser$ + vbCrLf

Else
    Unload Me
End If
Me.MousePointer = vbDefault
PIForm(gActivePIInstance).MousePointer = vbDefault
'PIForm(gActivePIInstance).IPPort1.Linger = True
Exit Sub
Fingererror:
    Beep
    Me.MousePointer = vbDefault
    PIForm(gActivePIInstance).MousePointer = vbDefault
    PIForm(gActivePIInstance).ShowStatus 1, "Finger error..."
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
PIForm(gActivePIInstance).IPPort1.WinsockLoaded = False
Set frmFingerCommand = Nothing
Me.MousePointer = vbDefault
PIForm(gActivePIInstance).MousePointer = vbDefault
End Sub

