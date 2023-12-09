VERSION 5.00
Begin VB.Form frmMailHeader 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit and View Mail Headers"
   ClientHeight    =   3615
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   6600
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
   LinkTopic       =   "Form34"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHeader 
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
      Index           =   7
      Left            =   1980
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   6
      Left            =   90
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2760
      Width           =   1785
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   5
      Left            =   1980
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2340
      Width           =   4335
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   4
      Left            =   90
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2340
      Width           =   1785
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   3
      Left            =   1980
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1950
      Width           =   4335
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   2
      Left            =   90
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1950
      Width           =   1785
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   1
      Left            =   1980
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Apply"
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
      Left            =   4110
      TabIndex        =   1
      Top             =   3210
      Width           =   975
   End
   Begin VB.TextBox txtHeader 
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
      Index           =   0
      Left            =   90
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   1785
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
      Left            =   5280
      TabIndex        =   2
      Top             =   3210
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Header Value:"
      Height          =   255
      Index           =   2
      Left            =   2010
      TabIndex        =   12
      Top             =   1320
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Header ID: "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   $"Mailhdr.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   6225
   End
End
Attribute VB_Name = "frmMailHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim i As Integer
    For i = 6 To UBound(MailHeader)
        MailHeader(i).ID = txtHeader((i - 6) * 2)
        MailHeader(i).Value = txtHeader((i - 6) * 2 + 1)
    Next
   ' MailHeader(0).Value = MailHeader(1).ID & MailHeader(1).Value & vbCrLf
   ' MailHeader(0).Value = MailHeader(0).Value & MailHeader(2).ID & MailHeader(2).Value & vbCrLf
   ' MailHeader(0).Value = MailHeader(0).Value & MailHeader(3).ID & MailHeader(3).Value & vbCrLf
   ' MailHeader(0).Value = MailHeader(0).Value & MailHeader(4).ID & MailHeader(4).Value & vbCrLf
    'MailHeader(0) = MailHeader(0) & MailHeader(5).ID & MailHeader(5).Value & vbCrLf
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim ItemString As String
Dim Header As String
        For i = 6 To UBound(MailHeader)
       ' MSFlexGrid1.Row = i

        txtHeader((i - 6) * 2) = MailHeader(i).ID
        txtHeader((i - 6) * 2 + 1) = MailHeader(i).Value
    Next
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

