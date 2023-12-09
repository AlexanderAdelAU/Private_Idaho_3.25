VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "Threed20.ocx"
Begin VB.Form frmMenuOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   1950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   4948
      _Version        =   131074
      AutoSize        =   2
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Create a sub folder under the selected group folder."
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   873
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "MenuOptions.frx":0000
         Caption         =   "New Folder"
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   1
      End
   End
End
Attribute VB_Name = "frmMenuOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Win As New CWindow

Win.Center Me, Null

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMenuOptions = Nothing
End Sub



Private Sub SSRibbon1_Click(Index As Integer, Value As Integer)
If Not Value Then Exit Sub
frmMain.CreateSubFolder
SSRibbon1(Index).Value = False
frmMain.DisplayInBox
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

End Sub
