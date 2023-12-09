VERSION 5.00
Begin VB.Form frmMultipleTransfers 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple transfer applications"
   ClientHeight    =   6090
   ClientLeft      =   915
   ClientTop       =   630
   ClientWidth     =   7590
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
   LinkTopic       =   "Form30"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   7590
   Begin VB.TextBox Text12 
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4800
      Width           =   4455
   End
   Begin VB.TextBox Text11 
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
      Left            =   2640
      TabIndex        =   11
      Top             =   5160
      Width           =   4455
   End
   Begin VB.TextBox Text10 
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
      Left            =   2640
      TabIndex        =   9
      Top             =   4440
      Width           =   4455
   End
   Begin VB.TextBox Text9 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   3600
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   8
      Top             =   3960
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   6
      Top             =   3240
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
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
      Height          =   300
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   4455
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
      Left            =   5160
      TabIndex        =   13
      Top             =   5640
      Width           =   975
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
      Left            =   6240
      TabIndex        =   14
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Help"
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
      Left            =   4065
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application script"
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
      TabIndex        =   17
      Top             =   5160
      Width           =   2400
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application window name"
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
      TabIndex        =   18
      Top             =   4800
      Width           =   2400
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application 4 name"
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
      TabIndex        =   27
      Top             =   4440
      Width           =   2400
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application script"
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
      TabIndex        =   26
      Top             =   3960
      Width           =   2400
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application window name"
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
      TabIndex        =   25
      Top             =   3600
      Width           =   2400
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application 3 name"
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
      TabIndex        =   24
      Top             =   3240
      Width           =   2400
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application script"
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
      TabIndex        =   23
      Top             =   2760
      Width           =   2400
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application window name"
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
      TabIndex        =   22
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application 2 name"
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
      TabIndex        =   21
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application script"
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
      TabIndex        =   20
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application window name"
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
      TabIndex        =   19
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application 1 name"
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
      TabIndex        =   16
      Top             =   840
      Width           =   2400
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   $"MULTRANS.frx":0000
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
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmMultipleTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Workfile:      MULTRNS.FRM (frmMultipleTransfers)
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
    SectionName = "Options"
    KeyName = "App1Name"
    KeyValue = Text4.Text
    If KeyValue <> "" Then
        frmPI.TransferApp1.Caption = KeyValue
        frmPI.TransferApp1.Visible = True
        frmPI.TransferApps.Visible = True
    Else
        frmPI.TransferApps.Visible = False
    End If
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App1Wind"
    KeyValue = Text1.Text
    WriteProfile SectionName, KeyName, KeyValue
    KeyValue = Text7.Text
    KeyName = "App1Script"
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App2Name"
    KeyValue = Text2.Text
    If KeyValue <> "" Then
        frmPI.TransferApp2.Caption = KeyValue
        frmPI.TransferApp2.Visible = True
    Else
        frmPI.TransferApp2.Visible = False
    End If
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App2Wind"
    KeyValue = Text5.Text
    WriteProfile SectionName, KeyName, KeyValue
    KeyValue = Text3.Text
    KeyName = "App2Script"
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App3Name"
    KeyValue = Text6.Text
    If KeyValue <> "" Then
        frmPI.TransferApp3.Caption = KeyValue
        frmPI.TransferApp3.Visible = True
    Else
        frmPI.TransferApp3.Visible = False
    End If
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App3Wind"
    KeyValue = Text9.Text
    WriteProfile SectionName, KeyName, KeyValue
    KeyValue = Text8.Text
    KeyName = "App3Script"
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App4Name"
    KeyValue = Text10.Text
    If KeyValue <> "" Then
        frmPI.TransferApp4.Caption = KeyValue
        frmPI.TransferApp4.Visible = True
    Else
        frmPI.TransferApp4.Visible = False
    End If
    WriteProfile SectionName, KeyName, KeyValue
    KeyName = "App4Wind"
    KeyValue = Text12.Text
    WriteProfile SectionName, KeyName, KeyValue
    KeyValue = Text11.Text
    KeyName = "App4Script"
    WriteProfile SectionName, KeyName, KeyValue
    Unload frmMultipleTransfers
End Sub

Private Sub Command2_Click()
'Workfile:      MULTRNS.FRM (frmMultipleTransfers)
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







    Unload frmMultipleTransfers
End Sub

Private Sub Command3_Click()
'Workfile:      MULTRNS.FRM (frmMultipleTransfers)
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

Private Sub Form_Load()
    Dim SectionName As String
    Dim tmpstr As String
    
    SectionName = "Options"
    tmpstr = ReadProfile(SectionName, "App1Name")
    Text4.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App1Wind")
    Text1.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App1Script")
    Text7.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App2Name")
    Text2.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App2Wind")
    Text5.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App2Script")
    Text3.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App3Name")
    Text6.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App3Wind")
    Text9.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App3Script")
    Text8.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App4Name")
    Text10.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App4Wind")
    Text12.Text = tmpstr
    tmpstr = ReadProfile(SectionName, "App4Script")
    Text11.Text = tmpstr
End Sub

