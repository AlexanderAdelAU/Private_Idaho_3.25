VERSION 5.00
Begin VB.Form frmMessageScanTimer 
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   4500
      TabIndex        =   3
      Top             =   870
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "The server will be scanned for either PGP or normal messages depending on the time you enter.  Zero (0) means don't scan."
      Height          =   555
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   60
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "minutes."
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Scan every:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   780
      Width           =   1155
   End
End
Attribute VB_Name = "frmMessageScanTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_EmailScanInterval As Integer
Private m_TimerInterval As Integer
Private Sub Command1_Click()
On Error Resume Next
'If Text1.Text = "0" Or Text1.Text = "" Then
   ' Timer1(1).Enabled = False
'Else
    'Timer1(1).Interval = CInt(1000 * 30) 'set it to go off every minute
    m_TimerInterval = CInt(1000 * 30)
    'giEmailScanInterval = val(Text1.Text) * 2
    m_EmailScanInterval = val(Text1.Text) * 2
    'giTimerCounter = giEmailScanInterval
    'Timer1(1).Enabled = True
'End If
Unload Me
'Err.Clear
End Sub
Private Sub Form_Load()
Text1.Text = CInt(m_EmailScanInterval / 2)
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Public Property Get ScanInterval() As Integer
ScanInterval = m_EmailScanInterval
End Property
Public Property Get TimerSetting() As Integer
TimerSetting = m_TimerInterval
End Property
Public Property Let ScanInterval(ByVal ScanInterval As Integer)
 m_EmailScanInterval = ScanInterval
End Property
Public Property Let TimerSetting(ByVal TimerSetting As Integer)
 m_TimerInterval = TimerSetting
End Property
