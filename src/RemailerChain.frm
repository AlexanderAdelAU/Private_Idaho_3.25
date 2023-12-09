VERSION 5.00
Begin VB.Form frmRemailerChain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remailer Chain"
   ClientHeight    =   5610
   ClientLeft      =   975
   ClientTop       =   810
   ClientWidth     =   8295
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
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   8295
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Auto Select"
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
      Left            =   6990
      TabIndex        =   14
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Load"
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
      Left            =   6990
      TabIndex        =   13
      Top             =   4590
      Width           =   975
   End
   Begin VB.CommandButton Command3 
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
      Height          =   315
      Left            =   6990
      TabIndex        =   12
      Top             =   4950
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Clear &All"
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
      Left            =   6990
      TabIndex        =   10
      Top             =   3990
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Remove"
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
      Left            =   6990
      TabIndex        =   9
      Top             =   3630
      Width           =   975
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
      Width           =   6435
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
      TabIndex        =   3
      Top             =   5160
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
      Left            =   4200
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   6345
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Chained order"
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
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "up time"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "latency"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "name"
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
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "last updated"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"RemailerChain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   360
      TabIndex        =   1
      Top             =   210
      Width           =   7305
   End
End
Attribute VB_Name = "frmRemailerChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Integer
Dim J As Integer

    gnumRemailers = List2.ListCount
    For i = 0 To gnumRemailers - 1
        J = InStr(1, List2.List(i), " ")
        Remailers(i + 1) = frmRemailerList.GetRemailer((Mid(List2.List(i), 1, J - 1)))
    Next
    Unload Me
End Sub

Private Sub Command2_Click()
    gCancelAction = True
    gnumRemailers = 0
    Unload Me
End Sub

Private Sub Command3_Click()
    NotYet
End Sub

Private Sub Command4_Click()
    If List2.ListIndex <> -1 Then
        List2.RemoveItem List2.ListIndex
    End If
End Sub

Private Sub Command5_Click()
    List2.Clear
End Sub

Private Sub Command6_Click()
    NotYet
End Sub

Private Sub Command7_Click()
    NotYet
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim maxName As Integer
Dim maxUp As Integer
Dim maxLatent As Integer
Dim tempString As String
Dim ListHandle As Long
ReDim lTabPos(2) As Long
   
    lTabPos(0) = 10
    lTabPos(1) = 32
    lTabPos(2) = 45
    maxName = 0
    maxUp = 0
    maxLatent = 0
    If Not gRemailerSelectCaption = "" Then lblCaption = gRemailerSelectCaption
'
'Need to do this to update the remailers in case the nym has upset the list
'
    frmRemailerList.SortRemailers
    frmRemailerList.FillRemailerList
    
    ListHandle = frmRemailerChain.List1.hWnd
    Call SetTabStops(CLng(ListHandle), 3, lTabPos())
    ListHandle = frmRemailerChain.List2.hWnd
    Call SetTabStops(CLng(ListHandle), 3, lTabPos())
    For i = 1 To frmRemailerList.List3.ListCount - 1
        List1.AddItem frmRemailerList.List3.List(i)
    Next
    Label2.Caption = Label6.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRemailerChain = Nothing
End Sub

Private Sub List1_Click()
    List2.AddItem List1.Text
End Sub

