VERSION 5.00
Begin VB.Form frmDisplayMessageHeader 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Header"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHeader 
      Height          =   4185
      Left            =   510
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   5625
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   405
      Left            =   4890
      TabIndex        =   0
      Top             =   4560
      Width           =   1245
   End
End
Attribute VB_Name = "frmDisplayMessageHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnok_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'DisplayHeader
End Sub

Private Sub Form_Load()

DisplayHeader
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDisplayMessageHeader = Nothing
End Sub

Private Sub DisplayHeader()
Dim rs As Recordset
On Error GoTo BadMessageDisplay
Set rs = DB.OpenRecordset(Grid.SelectedQuery, dbOpenDynaset)
If rs.EOF Then
    Beep
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.MoveFirst
End If
For i = 1 To frmMain.MSFlexGrid1.Row - 1
    rs.MoveNext
Next
txtHeader = rs("MIME Message Header")

rs.Close

Exit Sub
BadMessageDisplay:
    MsgBox Err.Description & " Can't find message", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

