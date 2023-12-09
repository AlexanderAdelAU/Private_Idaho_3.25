VERSION 5.00
Begin VB.Form frmDisplayMessageHeader 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4185
      Left            =   510
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmDisplayMessageHeader.frx":0000
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
Private Sub btnOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FileName As String
Dim FileNum As Integer
Dim StartSection As Integer
Dim i As Integer
Dim J As Integer
Dim Sqlq As String
Dim rs As Recordset
Dim lListItem As ListItem
Dim AttachmentFileName As String

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

'Okay clear the message area
frmPI.MessageArea.Text = " "
For i = 1 To frmMain.MSFlexGrid1.Row - 1
    rs.MoveNext
Next

'To get here it must be the right row
'FileName = rs("Message File Name")
FileName = rs("MIME Message")
frmPI.lvwAttachments.ListItems.Clear
frmPI.txtsubject.Text = rs("Subject")
If rs("Message Sent") Then
    frmPI.btnTo.Item(0).Caption = "To"
    frmPI.txtTo.Text = rs("To")
Else
    frmPI.txtTo.Text = rs("From")
    frmPI.btnTo.Item(0).Caption = "From"
End If
frmPI.txtCC.Text = rs("CC")
FileName = App.Path & "\mailbox\" & FileName
'Clipboard.SetText (rs("MIME Message Header"))
gMessageRecord.Header = IIf(IsNull(rs("MIME Message Header")), "", rs("MIME Message Header"))
gMessage = GetFileText(FileName) 'rs("MIME Message")
'Clipboard.SetText (GetFileText(FileName))
rs.Edit
rs("Message Read") = True
rs.Update
rs.Close
If Not InStr(1, gMessageRecord.Header, "boundary=") = 0 Then
    frmPI.MessageArea.Text = gMessageRecord.Header & vbCrLf & gMessage
   ' frmPI.Show
    frmPI.DecodeMessage
Else
    frmPI.MessageArea.Text = gMessage
End If
frmPI.Show
Exit Sub
    frmPI.MIME1.Action = a_ResetData
    'frmPI.MIME1.MessageHeaders = gMessageRecord.Header
    frmPI.MIME1.Message = FileName
    frmPI.MIME1.Action = a_DecodeFromFile

    frmPI.MessageArea.Text = StripNulls(frmPI.MIME1.PartDecodedString(0))
    For i = 1 To frmPI.MIME1.PartCount - 1
        If frmPI.MIME1.PartContentDisposition(i) = "attachment" Then
            AttachmentFileName = frmPI.MIME1.PartName(i)
            J = frmPI.FileIconsImageList1.GetFileIconNum(StripFileName(AttachmentFileName))
            Set lListItem = frmPI.lvwAttachments.ListItems.Add(, , AttachmentFileName, J, J)
        End If
    Next
Set rs = Nothing
'frmpi.Show
frmPI.Show
Exit Sub
BadMessageDisplay:
    MsgBox Err.Description & " Can't find message", vbApplicationModal + vbCritical, App.Title
    Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDisplayMessageHeader = Nothing
End Sub
