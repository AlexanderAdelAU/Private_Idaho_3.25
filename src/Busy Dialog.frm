VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmBusy 
   Caption         =   "Busy Dialog"
   ClientHeight    =   2070
   ClientLeft      =   2010
   ClientTop       =   5535
   ClientWidth     =   4005
   ControlBox      =   0   'False
   Icon            =   "Busy Dialog.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2070
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   1200
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1410
      TabIndex        =   2
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Timer tmrBusy 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Caption         =   "Progress: 1 of 100"
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "We are busy now. Please wait ..."
      Height          =   705
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   3585
   End
End
Attribute VB_Name = "frmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Local property storage variables
Dim mnBarPercent        As Integer
Dim msMessage           As String
Dim msBarCaption        As String
Dim mnStyle             As Integer
Dim mbAllowCancel       As Boolean
Dim frmCalling          As Object

Dim mnFrame             As Integer

Private Sub cmdAction_Click()

frmCalling.BusyCancel = True

End Sub

Private Sub Form_Load()

' Setup the defaults
mnFrame = 0         ' First frame
Me.Style = 0        ' Animation
Me.AllowCancel = True

End Sub

Private Sub tmrBusy_Timer()

' Advance to the next animation frame
mnFrame = mnFrame + 1

' Wrap around to the first frame if at the last one
If mnFrame > 7 Then mnFrame = 0

' Load the icon for the current frame
'imgBusy.Picture = imgFrame(mnFrame).Picture

End Sub

Public Property Get BarPercent() As Integer

BarPercent = mnBarPercent

End Property

Public Property Let BarPercent(nBarPercent As Integer)

mnBarPercent = nBarPercent
ProgressBar1.Value = mnBarPercent

End Property

Public Property Get AllowCancel() As Boolean

AllowCancel = mbAllowCancel

End Property

Public Property Let AllowCancel(bAllowCancel As Boolean)

mbAllowCancel = bAllowCancel
cmdAction.Visible = mbAllowCancel

Select Case mbAllowCancel
    Case True
        frmBusy.Height = 2475
        
    Case False
        frmBusy.Height = 1995
    
End Select
    
End Property

Public Property Get CallingForm() As Object

Set CallingForm = frmCalling

End Property

Public Property Let CallingForm(objCallingForm As Object)

Set frmCalling = objCallingForm

End Property

Public Property Get Style() As Integer

Style = mnStyle

End Property

Public Property Let Style(nStyle As Integer)

mnStyle = nStyle

' Check for legal property values
Select Case mnStyle
    Case 0          ' Animation
       ' imgBusy.Visible = True
        tmrBusy.Enabled = True
        lblProgress.Visible = False
        ProgressBar1.Visible = False
            
    Case 1          ' Progress bar
       ' imgBusy.Visible = False
        tmrBusy.Enabled = False
        lblProgress.Visible = True
        ProgressBar1.Visible = True
        ProgressBar1.Value = 1
    
End Select

End Property

Public Property Get BarCaption() As String

BarCaption = msBarCaption

End Property

Public Property Let BarCaption(sBarCaption As String)

msBarCaption = sBarCaption
lblProgress.Caption = msBarCaption

End Property

Public Property Get Message() As String

Message = msMessage

End Property

Public Property Let Message(ByVal sMessage As String)

msMessage = sMessage
lblMessage.Caption = msMessage

End Property
