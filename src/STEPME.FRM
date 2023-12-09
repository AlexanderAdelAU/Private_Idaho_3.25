VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form32 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Step me through..."
   ClientHeight    =   4680
   ClientLeft      =   2175
   ClientTop       =   615
   ClientWidth     =   7320
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
   LinkTopic       =   "Form32"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4680
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3735
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"STEPME.frx":0000
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Done"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Workfile:      STEPME.FRM (Form32)
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







    SetWindowPos Form32.hWnd, -2, 0, 0, 0, 0, &H50 'This will put "Always Visible" off!
    Unload Form32
End Sub

Private Sub Form_Load()
'Workfile:      STEPME.FRM (Form32)
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







    'Form32.Top = PIForm(gActivePIInstance).Top + 500
   ' Form32.Left = (PIForm(gActivePIInstance).Width - Form32.Width) + 500
    'SetWindowPos Form32.hWnd, -1, 0, 0, 250, 370, &H50
    'Form32.Text1 = "Encrypting a message" + gCRLF + gCRLF + "Step 1.  Enter the user ID on the To: line.  This is usually a person's e-mail address.  TIP: Click on To: to get a list of PGP public keys."
End Sub

Private Sub Form_Resize()
Dim lWidth As Long
Dim ButtonTop As Long
 If WindowState <> 1 Then
    On Error Resume Next
    Command1.Top = Height - 1000
    RichTextBox1.Width = Width - RichTextBox1.Left - Width * 0.05
    RichTextBox1.Height = Command1.Top - Command1.Height - 10
    Command1.Left = RichTextBox1.Left + RichTextBox1.Width - Command1.Width
    
 End If
End Sub
