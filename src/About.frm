VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Private I Email"
   ClientHeight    =   4380
   ClientLeft      =   1110
   ClientTop       =   720
   ClientWidth     =   7305
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
   ScaleHeight     =   4380
   ScaleWidth      =   7305
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
      Left            =   5580
      TabIndex        =   1
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "The spgp.dll library used by this programme is copyright to S.R.Heller."
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
      Height          =   555
      Index           =   1
      Left            =   4560
      TabIndex        =   6
      Top             =   2190
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.itech.net.au/pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblNag 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   180
      TabIndex        =   4
      Top             =   2910
      Visible         =   0   'False
      Width           =   5745
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":00DE
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1140
      Width           =   6735
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Private i email is a PGP, anonymous remailer, and nym server utility.  It helps keep your e-mail communications secure."
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()


    Dim strFile As String
    Dim udtFileInfo As FILEINFO
    On Error Resume Next


Exit Sub

  '  With CommonDialog1
   '     .Filter = "All Files (*.*)|*.*"
   '     .ShowOpen
    '    strFile = .FileName
   '     If Err.Number = cdlCancel Or strFile = "" Then Exit Sub
 '   End With

    strFile = "c:\windows\system32\PGPSDK.dll"

    If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
        MsgBox "No version available For this file", vbInformation
        Exit Sub
    End If

   ' Label3 = "Company Name: " & udtFileInfo.CompanyName & vbCrLf
    'Label3 = Label3 & "File Description:" & udtFileInfo.FileDescription & vbCrLf
    'Label3 = Label3 & "File Version:" & udtFileInfo.FileVersion & vbCrLf
   ' Label3 = Label3 & "Internal Name: " & udtFileInfo.InternalName & vbCrLf
   ' Label3 = Label3 & "Legal Copyright: " & udtFileInfo.LegalCopyright & vbCrLf
   ' Label3 = Label3 & "Original FileName:" & udtFileInfo.OriginalFileName & vbCrLf
   ' Label3 = Label3 & "Product Name:" & udtFileInfo.ProductName & vbCrLf
   ' Label3 = Label3 & "Product Version: " & udtFileInfo.ProductVersion & vbCrLf

End Sub

Private Sub Form_Load()
Dim Win As New CWindow
Dim App As New CApplication
Win.Center Me, Null
Const NameLength = 28
Dim name As String * NameLength
Dim SerNumLen As Integer
Dim i As Long
lblVersion = App.Version
   ' Set application title
   Me.Caption = "About " & App.Title

   ' Center the form
   Win.Center Me, Null
   Win.OnTop(Me) = True
   
  
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAbout = Nothing
End Sub


