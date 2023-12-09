VERSION 5.00
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-mail Transfer Options"
   ClientHeight    =   5745
   ClientLeft      =   780
   ClientTop       =   360
   ClientWidth     =   7605
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
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   7605
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Multiple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   12
      Top             =   1755
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   6240
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   6240
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   3840
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   15
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "You can specify multiple applications to transfer Private Idaho information to."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   7215
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"EUTRNOPT.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   7215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"EUTRNOPT.frx":00A8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   7335
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"EUTRNOPT.frx":01A3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "E-mail application window to transfer to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   3495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Workfile:      EUTRNOPT.FRM (Form5)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/17/97
'Description:   Added comments.







    Dim SectionName As String
    Dim KeyName As String
    Dim KeyValue As String
    SectionName = "Mailer Info"
    If Combo1.Text <> gEmailer Then
        KeyName = "Emailer"
        KeyValue = Combo1.Text
        WriteProfile SectionName, KeyName, KeyValue
        gEmailer = Combo1.Text
    End If
    If Text7.Text <> gtranScript Then
        KeyName = "gtranScript"
        KeyValue = Text7.Text
        WriteProfile SectionName, KeyName, KeyValue
        gtranScript = Text7.Text
    End If
    Unload Form5
End Sub

Private Sub Command2_Click()
'Workfile:      EUTRNOPT.FRM (Form5)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/17/97
'Description:   Added comments.







    Unload Form5
End Sub

Private Sub Command3_Click()
'Workfile:      EUTRNOPT.FRM (Form5)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/17/97
'Description:   Added comments.







'    frmPI.CommonDialog1.HelpFile = App.Path + "\PIDAHO.HLP"
'    frmPI.CommonDialog1.HelpCommand = &H1
'    frmPI.CommonDialog1.HelpContext = 10
'    frmPI.CommonDialog1.Action = 6
End Sub

Private Sub Command4_Click()
'Workfile:      EUTRNOPT.FRM (Form5)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/17/97
'Description:   Added comments.







    frmMultipleTransfers.Show 1
End Sub

Private Sub Command5_Click(Index As Integer)
'Workfile:      EUTRNOPT.FRM (Form5)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/17/97
'Description:   Added comments.







    Combo1.Text = RTrim$(gMailerInfo(Index).MailName)
    Text7.Text = RTrim$(gMailerInfo(Index).Script)
End Sub

Private Sub Form_Load()
'Workfile:      EUTRNOPT.FRM (Form5)
'Copyright:     Copyright © 1995, 1996, 1997 Joel McNamara
'Copyright:     Copyright © 1997 BC
'Author:        Joel McNamara
'Description:   @@?
'---------------------------------------------
'Change History:
'---------------------------------------------
'Developer:     BC
'Change Date:   05/17/97
'Description:   Added comments.







    SetEmailers
    Combo1.Text = gEmailer
    Text7.Text = gtranScript
    Call LoadTaskList
End Sub

Private Sub LoadTaskList()
Dim CurrentWindow As Long
Dim ListItem As String
Dim Length As Integer

    'Get the hWnd of the first item in the master list
    'so we can process the task list entries (top-level only).
    CurrentWindow = GetWindow(Form5.hwnd, GW_HWNDFIRST)
    ' Loop while the hWnd returned by GetWindow is valid.
    While CurrentWindow <> 0
        'Get the length of the task name identified by
        'CurrentWindow in the list.
        Length = GetWindowTextLength(CurrentWindow)
        'Get the task name of the task in the master list.
        ListItem = Space(Length + 1)
        Length = GetWindowText(CurrentWindow, ListItem, Length + 1)
        'If there is an actual task name in the list, add the
        'item to the list.
        If Length > 0 Then
            Combo1.AddItem ListItem
        End If
        'Get the next task list item in the master list.
        CurrentWindow = GetWindow(CurrentWindow, GW_HWNDNEXT)
        'Process Windows events.
        DoEvents
    Wend
End Sub

