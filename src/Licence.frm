VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Check"
   ClientHeight    =   3600
   ClientLeft      =   3630
   ClientTop       =   2445
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3600
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   3525
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   6218
      _Version        =   131074
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   3690
         TabIndex        =   3
         Top             =   2610
         Width           =   1065
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   405
         Left            =   1800
         TabIndex        =   1
         Top             =   2610
         Width           =   1065
      End
      Begin MSMask.MaskEdBox txtLicenceNumber 
         Height          =   375
         Left            =   2490
         TabIndex        =   0
         Top             =   1380
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "#####-#######"
         PromptChar      =   "_"
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "Licence.frx":0000
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Private Idaho - Licence Validation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   3
         Left            =   1020
         TabIndex        =   5
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   1290
         TabIndex        =   4
         Top             =   1410
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim LenWritten As Integer
    Dim SerialNumber As String
    Dim EncryptedSerialNumber As String
    Const NumHexChars = 26
    Dim HexSerialNumber As String * NumHexChars
    Dim SectionName As String

    SectionName = "PGP Options"
    Const Password As String = "~!@#$%^&*()"
    'These are written to the ini file and then read by main
    'LenWritten = WritePrivateProfileString("Licence Details", "Name", txtName, giniFile)
    'LenWritten = WritePrivateProfileString("Licence Details", "Company Name", txtCompanyName, giniFile)
    SerialNumber = txtLicenceNumber
    SimpleCrypt SerialNumber, EncryptedSerialNumber, Password 'Plaintext, CipherText, KeyValue
    HexSerialNumber = BinHex(EncryptedSerialNumber)
    
    'gPGPVersion = WriteProfile(SectionName, "PGP Version")
    'WriteProfile SectionName, "Registration Number", HexSerialNumber
    Dim rs As Recordset
    Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
    rs.Edit
    rs("Serial Number") = HexSerialNumber
    rs.Update
    rs.Close
   
   ' LenWritten = WritePrivateProfileString("Licence Details", "Serial Number", HexSerialNumber, giniFile)
    lblLicence = SecurityCheck
    Unload Me
   
End Sub


Private Sub Form_Load()
Dim Win As New CWindow
Dim App As New CApplication
 
 Me.Caption = App.Title & " - Licence Details"
 Win.Center Me, Null
 'Win.OnTop(Me) = True
 txtLicenceNumber.SelStart = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
'Set Win = Nothing
'Set App = Nothing
Set frmLicence = Nothing
End Sub


Private Sub txtLicenceNumber_GotFocus()
txtLicenceNumber.SelStart = 0
End Sub

