VERSION 5.00
Begin VB.Form frmCheckMail 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PGP messages"
   ClientHeight    =   2850
   ClientLeft      =   720
   ClientTop       =   1110
   ClientWidth     =   8505
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   8505
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O&ptions"
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
      Left            =   4800
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Compress"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Undelete"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command4 
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
      Height          =   300
      Left            =   7320
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Decr&ypt"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Save"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Delete"
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
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "subject"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "sent"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "from"
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
      Left            =   720
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmCheckMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
    If List1.ListIndex <> -1 Then
        i = ReadIndex(List1.ListIndex + 1)
        ReadMessage (i)
        frmMain.MessageArea.Text = gMessageRecord.Contents
        Unload Me
    Else
        MsgBox "Select a message."
    End If
End Sub

Private Sub Command2_Click()
    Dim Y As Long
    Dim tmpstr As String
    
    If List1.ListIndex <> -1 Then
        Y = ReadIndex(List1.ListIndex + 1)
        WriteDeleteFlag (Y)
        tmpstr = List1.List(List1.ListIndex)
        List1.List(List1.ListIndex) = "d" + Mid$(tmpstr, 2, Len(tmpstr) - 1)
    Else
        MsgBox "Select a message."
    End If
End Sub

Private Sub Command3_Click()
Dim Y As Long

    If List1.ListIndex <> -1 Then
        Y = ReadIndex(List1.ListIndex + 1)
        ReadMessage (Y)
        frmMain.MessageArea.Text = gMessageRecord.Contents
        gPOPState = POPDECRYPT
        WriteReadFlag (Y)
        Unload Me
        'sleazy hack to call the decrypt routine
        DecryptMessageText
    Else
        MsgBox "Select a message."
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
Dim Y As Long
Dim tmpstr As String

    If List1.ListIndex <> -1 Then
        Y = ReadIndex(List1.ListIndex + 1)
        WriteUndeleteFlag (Y)
        tmpstr = List1.List(List1.ListIndex)
        If Mid$(tmpstr, 1, 1) = "d" Then
            List1.List(List1.ListIndex) = " " + Mid$(tmpstr, 2, Len(tmpstr) - 1)
        End If
    Else
        MsgBox "Select a message."
    End If
End Sub

Private Sub Command6_Click()

    CompressMessages
    Unload Me
End Sub

Private Sub Command7_Click()
    Form28.Show vbModal
End Sub

Private Sub Form_Load()

    Dim IdxFile As Integer
    Dim i As Integer
    Dim IdxLen As Long
    Dim Msgindex As Long
    
    IdxFile = FreeFile
    IdxLen = Len(MasterIndex)
    Open App.Path + "\msgidx.pid" For Random As IdxFile Len = IdxLen
    Msgindex = LOF(IdxFile) / IdxLen
    Close IdxFile
    For i = 1 To CInt(Msgindex)
        ReadMessage (ReadIndex(i))
        CreateListItem
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCheckMail = Nothing
End Sub

Private Sub List1_DblClick()
    Dim Y As Long
    
    Y = ReadIndex(List1.ListIndex + 1)
    ReadMessage (Y)
    gPOPState = POPDECRYPT
    frmMain.MessageArea.Text = gMessageRecord.Contents
    WriteReadFlag (Y)
    Unload Me
    'sleazy hack to call the decrypt routine
    DecryptMessageText
End Sub
