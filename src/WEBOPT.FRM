VERSION 5.00
Begin VB.Form Form15 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web Browser Options"
   ClientHeight    =   6960
   ClientLeft      =   540
   ClientTop       =   360
   ClientWidth     =   8445
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6960
   ScaleWidth      =   8445
   Begin VB.TextBox Text5 
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
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "http://www.anonymizer.com:8080"
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
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
      Left            =   6885
      TabIndex        =   6
      Top             =   5325
      Width           =   975
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5790
      TabIndex        =   5
      Top             =   5325
      Width           =   975
   End
   Begin VB.TextBox Text4 
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
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Text            =   "~"
      Top             =   3540
      Width           =   3495
   End
   Begin VB.TextBox Text3 
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
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Text            =   "^l"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "Netscape - ["
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   "c:/netscape/netscape.exe"
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "C2 Web anonymizer URL:"
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
      Left            =   360
      TabIndex        =   23
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   225
      X2              =   8145
      Y1              =   5850
      Y2              =   5850
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WEBOPT.frx":0000
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
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WEBOPT.frx":0101
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
      Height          =   615
      Left            =   360
      TabIndex        =   21
      Top             =   3900
      Width           =   7815
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WEBOPT.frx":01CD
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
      Height          =   420
      Left            =   360
      TabIndex        =   20
      Top             =   3000
      Width           =   7695
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"WEBOPT.frx":026A
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
      Left            =   360
      TabIndex        =   19
      Top             =   2040
      Width           =   7815
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter the full path of the Web browser you use."
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
      Left            =   360
      TabIndex        =   18
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Special keys for key sequences:"
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
      Left            =   585
      TabIndex        =   17
      Top             =   5970
      Width           =   2775
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For keys to be held down at the same time, enclose them with ( )"
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
      Left            =   4890
      TabIndex        =   16
      Top             =   6465
      Width           =   3375
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For the Shift key use a +"
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
      Left            =   4890
      TabIndex        =   15
      Top             =   6225
      Width           =   2175
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For a return/enter use ~"
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
      Left            =   2490
      TabIndex        =   14
      Top             =   6465
      Width           =   2175
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For a tab use {TAB}"
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
      Left            =   2490
      TabIndex        =   13
      Top             =   6225
      Width           =   2175
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For the Alt key use a %"
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
      Left            =   210
      TabIndex        =   12
      Top             =   6465
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For the Ctrl key use a ^"
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
      Left            =   210
      TabIndex        =   11
      Top             =   6225
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Key sequence to get URL:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3540
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Key sequence to prompt for URL:"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Browser window search string:"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Web browser path:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

'    frmPI.CommonDialog1.HelpFile = App.Path + "\PIDAHO.HLP"
'    frmPI.CommonDialog1.HelpCommand = &H1
'    frmPI.CommonDialog1.HelpContext = 5
'    frmPI.CommonDialog1.Action = 6
End Sub

Private Sub Command2_Click()
    Dim SectionName As String
    
    SectionName = "Web Info"
    If Text1.Text <> gBrowserPath Then
        gBrowserPath = Text1.Text
        WriteProfile SectionName, "BrowserPath", gBrowserPath
    End If
    If Text2.Text <> gBrowserString Then
        gBrowserString = Text2.Text
        WriteProfile SectionName, "BrowserString", gBrowserString
    End If
    If Text3.Text <> gURLStart Then
        gURLStart = Text3.Text
        WriteProfile SectionName, "URLStart", gURLStart
    End If
    If Text4.Text <> gURLEnd Then
        gURLEnd = Text4.Text
        WriteProfile SectionName, "URLEnd", gURLEnd
    End If
    If Text5.Text <> gc2WWWAnon Then
        gc2WWWAnon = Text5.Text
        WriteProfile SectionName, "gc2WWWAnon", gc2WWWAnon
    End If
    Unload Form15
End Sub

Private Sub Command3_Click()

    Unload Form15
End Sub

Private Sub Form_Load()
    Text1.Text = gBrowserPath
    Text2.Text = gBrowserString
    Text3.Text = gURLStart
    Text4.Text = gURLEnd
    Text5.Text = gc2WWWAnon
End Sub

Private Function GetAppWindow(searchStr As String) As String
Dim CurrentWindow As Long
Dim Length As Integer
Dim ListItem As String

    GetAppWindow = "null"
    'Get the hWnd of the first item in the master list
    'so we can process the task list entries (top-level only).
    CurrentWindow = GetWindow(frmPI.hwnd, GW_HWNDFIRST)
    ' Loop while the hWnd returned by GetWindow is valid.
    While CurrentWindow <> 0
        'Get the length of the task name identified by
        'CurrentWindow in the list.
        Length = GetWindowTextLength(CurrentWindow)
        'Get the task name of the task in the master list.
        ListItem$ = Space$(Length + 1)
        Length = GetWindowText(CurrentWindow, ListItem$, Length + 1)
        'If there is an actual task name in the list, add the
        'item to the list.
        If Length > 0 Then
            If InStr(1, ListItem$, searchStr) Then
                GetAppWindow = ListItem$
                Exit Function
            End If
        End If
        'Get the next task list item in the master list.
        CurrentWindow = GetWindow(CurrentWindow, GW_HWNDNEXT)
        'Process Windows events.
        DoEvents
    Wend
End Function

Private Sub GetURL(URLStr As String)
Dim tmpstr As String

    tmpstr = Text2.Text
    tmpstr = GetAppWindow(tmpstr)
    If tmpstr <> "null" Then
        AppActivate tmpstr
        SendKeys "% x", True
        SendKeys Text3.Text, True
        SendKeys URLStr, True
        SendKeys Text4.Text, True
    End If
End Sub
