VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmPGPUseOptions 
   Caption         =   "PGP Use - Options"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   5415
      Begin Threed.SSOption optPGP 
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         _Version        =   131074
         ForeColor       =   16711680
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "PGP Use Options.frx":0000
         Caption         =   "Don't use PGP"
         Value           =   -1
      End
      Begin Threed.SSOption optPGP 
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         _Version        =   131074
         ForeColor       =   16711680
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "PGP Use Options.frx":05C4
         Caption         =   "I'll try again later"
      End
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin Threed.SSCheck chkShowAgain 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   131074
      ForeColor       =   255
      Caption         =   "Don't show this again on startup."
      MaskColor       =   16777215
   End
   Begin VB.Label Label1 
      Caption         =   $"PGP Use Options.frx":0C26
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmPGPUseOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnOk_Click()
Dim SectionName As String
If optPGP(0).Value = True Then
    gPGPVersion = PGPNotFound
Else
    gPGPVersion = NoPGP
End If
SectionName = "Options"
WriteProfile SectionName, "PGPStatus", gPGPVersion
'SectionName = "Options"
Unload Me
End Sub

Private Sub Form_Load()
If gPGPVersion = PGPNotFound Then
    optPGP(0).Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPGPUseOptions = Nothing
End Sub

