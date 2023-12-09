VERSION 5.00
Begin VB.Form Form25 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple IDs"
   ClientHeight    =   3675
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   6480
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
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   6480
   Begin VB.CommandButton Command4 
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
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command5 
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
      Left            =   4200
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text8 
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
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
      Width           =   4095
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
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox Text6 
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
      Left            =   2160
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   4095
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
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   4095
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
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   4095
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
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4095
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
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
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
      TabIndex        =   17
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
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
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
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
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
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
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
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
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
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
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
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
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Dim SectionName As String
    Dim KeyName As String
    Dim KeyValue As String
    
    If gMultiType = 1 Then
        SectionName = "PGP Options"
        KeyName = "User ID 1"
        KeyValue = Text1.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 2"
        KeyValue = Text2.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 3"
        KeyValue = Text3.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 4"
        KeyValue = Text4.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 5"
        KeyValue = Text5.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 6"
        KeyValue = Text6.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 7"
        KeyValue = Text7.Text
        WriteProfile SectionName, KeyName, KeyValue
        KeyName = "User ID 8"
        KeyValue = Text8.Text
        WriteProfile SectionName, KeyName, KeyValue
    End If
    Unload Me
End Sub

Private Sub Form_Load()

    Dim SectionName As String
    Dim tmpstr As String
    
    If gMultiType = 1 Then
        tmpstr = "PGP User ID "
        Label1.Caption = tmpstr + "1"
        Label2.Caption = tmpstr + "2"
        Label3.Caption = tmpstr + "3"
        Label4.Caption = tmpstr + "4"
        Label5.Caption = tmpstr + "5"
        Label6.Caption = tmpstr + "6"
        Label7.Caption = tmpstr + "7"
        Label8.Caption = tmpstr + "8"
        SectionName = "PGP Options"
        Text1.Text = ReadProfile(SectionName, "User ID 1")
        Text2.Text = ReadProfile(SectionName, "User ID 2")
        Text3.Text = ReadProfile(SectionName, "User ID 3")
        Text4.Text = ReadProfile(SectionName, "User ID 4")
        Text5.Text = ReadProfile(SectionName, "User ID 5")
        Text6.Text = ReadProfile(SectionName, "User ID 6")
        Text7.Text = ReadProfile(SectionName, "User ID 7")
        Text8.Text = ReadProfile(SectionName, "User ID 8")
    End If
    If gMultiType = 0 Then
        tmpstr = "nym ID "
        Label1.Caption = tmpstr + "1"
        Label2.Caption = tmpstr + "2"
        Label3.Caption = tmpstr + "3"
        Label4.Caption = tmpstr + "4"
        Label5.Caption = tmpstr + "5"
        Label6.Caption = tmpstr + "6"
        Label7.Caption = tmpstr + "7"
        Label8.Caption = tmpstr + "8"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form25 = Nothing
End Sub
