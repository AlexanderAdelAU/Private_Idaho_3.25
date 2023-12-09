VERSION 5.00
Begin VB.Form Form17 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit options"
   ClientHeight    =   5295
   ClientLeft      =   1275
   ClientTop       =   1200
   ClientWidth     =   6975
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
   LinkTopic       =   "Form17"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   6975
   Begin VB.OptionButton Option4 
      Caption         =   "Custom fonts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   " 75 characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   " 70 characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   4305
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   " 65 characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5760
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text2 
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
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2280
      Width           =   6495
   End
   Begin VB.TextBox Text1 
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
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Word wrap line length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Signature"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Header"
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim FileNum As Integer
    Dim i As Integer
    
    '---------------------------------------------
    'save the user's preferred message header
    '---------------------------------------------
    If Text1.Text <> gHeader Then
        gHeader = Text1.Text
        '---------------------------------------------
        'see if file exists, if so then delete
        '---------------------------------------------
        If iFileExists(App.Path + "\HEADER.TXT") Then
            Kill App.Path + "\HEADER.TXT"
        End If
        '---------------------------------------------
        'write string to new file
        '---------------------------------------------
        FileNum = FreeFile
        Open App.Path + "\HEADER.TXT" For Output As FileNum
        Print #FileNum, gHeader
        Close FileNum
    End If
    '---------------------------------------------
    'save the user's preferred signature
    '---------------------------------------------
    If Text2.Text <> gSig Then
        gSig = Text2.Text
        '---------------------------------------------
        'see if file exists, if so then delete
        '---------------------------------------------
        If iFileExists(App.Path + "\SIG.TXT") Then
            Kill App.Path + "\SIG.TXT"
        End If
        '---------------------------------------------
        'write string to new file
        '---------------------------------------------
        FileNum = FreeFile
        Open App.Path + "\SIG.TXT" For Output As FileNum
        Print #FileNum, gSig
        Close FileNum
    End If
   ' PIForm(gActivePIInstance).MessageArea.Width = PIForm(gActivePIInstance).Width - 600
   
    Select Case gwrapNum
    Case -1
        'PIForm(gActivePIInstance).MessageArea.SelRightIndent = 20 ' = 10 '.Width = PIForm(gActivePIInstance).Width - 705
        
        'i = 65 * 0.55 * ScaleX(PIForm(gActivePIInstance).MessageArea.SelFontSize, vbPoints, vbTwips)
        PIForm(gActivePIInstance).MessageArea.RightMargin = 0
    Case 65
         i = 65 * 0.55 * ScaleX(PIForm(gActivePIInstance).MessageArea.SelFontSize, vbPoints, vbTwips)
        PIForm(gActivePIInstance).MessageArea.RightMargin = i
    Case 70
         i = 70 * 0.55 * ScaleX(PIForm(gActivePIInstance).MessageArea.SelFontSize, vbPoints, vbTwips)
        PIForm(gActivePIInstance).MessageArea.RightMargin = i
    Case 75
        i = 75 * 0.55 * ScaleX(PIForm(gActivePIInstance).MessageArea.SelFontSize, vbPoints, vbTwips)
        PIForm(gActivePIInstance).MessageArea.RightMargin = i
    End Select
    WriteProfile "Settings", "Wrap", Format$(gwrapNum)
    Unload Form17
End Sub

Private Sub Command2_Click()
'Workfile:      EUSIG.FRM (form17)
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







    Unload Me
End Sub

Private Sub Form_Load()
Dim FileNum As Integer
Dim tmpstr As String

    FileNum = FreeFile
    Text1.Text = gHeader
    If gHeader = "" Then
        'if file exists then open
        If iFileExists(App.Path + "\HEADER.TXT") Then
            Open App.Path + "\HEADER.TXT" For Input As FileNum
            tmpstr = Input$(LOF(FileNum), FileNum)
            Text1.Text = tmpstr
            gHeader = tmpstr
            Close FileNum
        End If
    End If
    FileNum = FreeFile
    Text2.Text = gSig
    If gSig = "" Then
        'if file exists then open
        If iFileExists(App.Path + "\SIG.TXT") Then
            Open App.Path + "\SIG.TXT" For Input As FileNum
            tmpstr = Input$(LOF(FileNum), FileNum)
            Text2.Text = tmpstr
            gSig = tmpstr
            Close FileNum
        End If
    End If
    Select Case gwrapNum
    Case 0
        Option4.Value = True
        Option1.Value = False
        Option2.Value = False
        Option3.Value = False
    Case 65
        Option4.Value = False
        Option1.Value = True
        Option2.Value = False
        Option3.Value = False
    Case 70
        Option4.Value = False
        Option1.Value = False
        Option2.Value = True
        Option3.Value = False
    Case 75
        Option4.Value = False
        Option1.Value = False
        Option2.Value = False
        Option3.Value = True
    End Select
End Sub

Private Sub Option1_Click()
    gwrapNum = 65
End Sub

Private Sub Option2_Click()

    gwrapNum = 70
End Sub

Private Sub Option3_Click()
    gwrapNum = 75
End Sub

Private Sub Option4_Click()
    gwrapNum = -1
End Sub
