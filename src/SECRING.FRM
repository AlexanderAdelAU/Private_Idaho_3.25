VERSION 5.00
Begin VB.Form Form24 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Select secret key"
   ClientHeight    =   3390
   ClientLeft      =   1350
   ClientTop       =   1830
   ClientWidth     =   6675
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
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   6675
   Begin VB.CommandButton Command5 
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
      Left            =   4320
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5400
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox List1 
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
      Height          =   2565
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Workfile:      SECRING.FRM (Form24)
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
    SectionName = "PGP Options"
    List1.AddItem ReadProfile(SectionName, "User ID 1")
    List1.AddItem ReadProfile(SectionName, "User ID 2")
    List1.AddItem ReadProfile(SectionName, "User ID 3")
    List1.AddItem ReadProfile(SectionName, "User ID 4")
    List1.AddItem ReadProfile(SectionName, "User ID 5")
    List1.AddItem ReadProfile(SectionName, "User ID 6")
    List1.AddItem ReadProfile(SectionName, "User ID 7")
    List1.AddItem ReadProfile(SectionName, "User ID 8")
End Sub
