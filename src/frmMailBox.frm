VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMailBox 
   Caption         =   "Mailbox"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "frmMailBox.frx":0000
      Height          =   3075
      Left            =   210
      TabIndex        =   1
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5424
      _Version        =   65541
      Rows            =   7
      Cols            =   6
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   7080
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.Image ImgList 
      Height          =   240
      Index           =   2
      Left            =   7800
      Picture         =   "frmMailBox.frx":0015
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgList 
      Height          =   240
      Index           =   1
      Left            =   7380
      Picture         =   "frmMailBox.frx":0145
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgList 
      Height          =   240
      Index           =   0
      Left            =   7020
      Picture         =   "frmMailBox.frx":0699
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click to retrieve message."
      Height          =   315
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   4440
      Width           =   2595
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press D to delete a message...."
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   4140
      Width           =   2595
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   3810
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMailBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As Recordset
Private Const CLOSED_ENVELOPE As Integer = 0
Private Const OPEN_ENVELOPE As Integer = 1
Private Const ATTACHMENT As Integer = 2



Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Activate()
Dim Itemstring As String
Dim RowIndex As Integer
Dim Headings As String
Exit Sub
RowIndex = 1
DoEvents
'in future just check for extra row and then add it..
If rs.RecordCount > MSFlexGrid1.Rows Then
    rs.MoveFirst
    MSFlexGrid1.Clear
    While MSFlexGrid1.Col = 1
        Itemstring = "" & vbTab & "" & vbTab & "" & vbTab & rs("From/To") & vbTab & rs("Subject") & vbTab & rs("Date")
        MSFlexGrid1.AddItem Itemstring, RowIndex
    
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = RowIndex
        If rs("Message Read") Then
            MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
            Set MSFlexGrid1.CellPicture = ImgList(OPEN_ENVELOPE).Picture
        Else
           MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
           Set MSFlexGrid1.CellPicture = ImgList(CLOSED_ENVELOPE).Picture
        End If
        MSFlexGrid1.Col = 2
        If rs("Attachment") Then
            MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
            Set MSFlexGrid1.CellPicture = ImgList(ATTACHMENT).Picture
        End If
        RowIndex = RowIndex + 1
        rs.MoveNext
    Wend
End If

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Headings As String
Dim RowIndex As Integer
Dim Itemstring As String
Dim Sqlq As String


On Error GoTo BadLoad
Select Case giMailbox
    Case SHOW_IN_MAILBOX
        lblTitle = "Incoming Message Archive - PGP ONLY"
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Incoming Message] = True) AND (Messages.[Message Deleted] = False) "
        Sqlq = Sqlq & "ORDER BY Messages.Date;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
        Headings = "    |^Read|^Att|<From                                     |<Subject                                  |<Date Received                "
    Case SHOW_OUT_MAILBOX
        lblTitle = "Outgoing Message Archive - All Messages"
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Incoming Message] = False) AND (Messages.[Message Deleted] = False) "
        Sqlq = Sqlq & "ORDER BY Messages.Date;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
        Headings = "    |^Read|^Att|<To                                      |<Subject                                  |<Date Sent                    "
    Case SHOW_DELETED_MAILBOX
        lblTitle = "Deleted Message Archive - All Messages"
        Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
        Sqlq = Sqlq & "(Messages.[Message Deleted] = True) "
        Sqlq = Sqlq & "ORDER BY Messages.Date;"
        Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
        Headings = "    |^Read|^Att|<To/From                                 |<Subject                                  |<Date                          "
    End Select
MSFlexGrid1.Cols = 6
MSFlexGrid1.Rows = 1
MSFlexGrid1.FormatString = Headings

RowIndex = 1
MSFlexGrid1.Clear
'MSFlexGrid1.Row = 1

While Not rs.EOF
    
    MSFlexGrid1.Col = 1
    Itemstring = "" & vbTab & "" & vbTab & "" & vbTab & rs("From/To") & vbTab & rs("Subject") & vbTab & rs("Date")
    MSFlexGrid1.AddItem Itemstring, RowIndex
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = RowIndex
    If rs("Message Read") Then
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(OPEN_ENVELOPE).Picture
    Else
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(CLOSED_ENVELOPE).Picture
    End If
    MSFlexGrid1.Col = 2
    If rs("Attachment") Then
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        Set MSFlexGrid1.CellPicture = ImgList(ATTACHMENT).Picture
    End If
    RowIndex = RowIndex + 1
    rs.MoveNext
Wend
MSFlexGrid1.FormatString = Headings
Exit Sub
BadLoad:
    MsgBox Err.Description & " Can't load properly...", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
DoEvents
Set frmMailBox = Nothing
End Sub

Private Sub MSFlexGrid1_DblClick()
Dim FileName As String
Dim FileNum As Integer
Dim StartSection As Integer
Dim i As Integer

On Error GoTo BadFileFind

rs.MoveFirst
If rs.EOF Then Exit Sub
'Okay clear the message area
'StartSection = frmMain.MessageArea.SelStart
frmMain.MessageArea.Text = " "
For i = 1 To MSFlexGrid1.Row - 1
    rs.MoveNext
   ' If rs.EOF Then Exit Sub
Next

'To get here it must be the right row
FileName = rs("Message File Name")
frmMain.Text2.Text = rs("Subject")
frmMain.Combo2.Text = rs("From/To")
rs.Edit
rs("Message Read") = True
rs.Update

'Place into Message area
ReadMailMessage (FileName)

If Not rs("Attachment File Name") = "" Then
   
    FileName = rs("Attachment File Name")
    StartSection = frmMain.MessageArea.SelStart
    frmMain.MessageArea.SelStart = Len(frmMain.MessageArea)
    frmMain.MessageArea.SelText = "========attachment========" & vbCrLf & "Attachment has been saved as: " & App.Path & "\mailbox\" & FileName & vbCrLf
    frmMain.MessageArea.SelStart = Len(frmMain.MessageArea)
    frmMain.MessageArea.SelText = vbCrLf & "If the attachment has an ASC extension then use the 'PGP File Operations' to decrypt it"
    frmMain.MessageArea = frmMain.MessageArea & vbCrLf & "If the attachment has an ASC extension then use the 'PGP File Operations' to decrypt it"
End If
frmMain.Show
Exit Sub
BadFileFind:
    MsgBox Err.Description & " Can't find message", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Dim Key As String
Dim i As Integer
Dim Heading As String
Dim Response As Integer
Static InHere As Boolean
Dim Sqlq As String

If InHere Then
    KeyAscii = 0
    Exit Sub
End If
On Error GoTo BadKeyEntry
InHere = True
Key = Chr$(KeyAscii)
Select Case Key
    Case "0" To "9"
      
    Case vbCr
        SendKeys "{TAB}"
        KeyAscii = 0
     
    Case Chr(&H7F)  'Delete Key
    Case "D", "d"
        'Dim FileName As String
       ' Dim i As Integer

        rs.MoveFirst
        If rs.EOF Then Exit Sub
        For i = 1 To MSFlexGrid1.Row - 1
            rs.MoveNext
        Next
        Select Case giMailbox
            Case SHOW_IN_MAILBOX
                rs.Edit
                rs("Message Deleted") = True
                rs.Update
                DoEvents
                Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
                Sqlq = Sqlq & "(Messages.[Incoming Message] = True) AND (Messages.[Message Deleted] = False) "
                Sqlq = Sqlq & "ORDER BY Messages.Date;"
                Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
            Case SHOW_OUT_MAILBOX
                rs.Edit
                rs("Message Deleted") = True
                rs.Update
                DoEvents
                Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
                Sqlq = Sqlq & "(Messages.[Incoming Message] = False) AND (Messages.[Message Deleted] = False) "
                Sqlq = Sqlq & "ORDER BY Messages.Date;"
                Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
            Case SHOW_DELETED_MAILBOX
                'To get here it must be the right row
                   
                Response = MsgBox("You are about to permenately delete the selected message from your system.  Do you wish to proceed?", vbYesNo + vbExclamation + vbApplicationModal, App.Title)
                If Response = vbNo Then Exit Sub
                On Error Resume Next
                If Not rs("Message File Name") = "" Then
                    Kill App.Path & "\mailbox\" & rs("Message File Name")
                End If
                If Not rs("Attachment File Name") = "" Then
                    Kill App.Path & "\mailbox\" & rs("Attachment File Name")
                End If
                rs.Delete
                Sqlq = "Select DISTINCTROW * FROM Messages WHERE "
                Sqlq = Sqlq & "(Messages.[Message Deleted] = True) "
                Sqlq = Sqlq & "ORDER BY Messages.Date;"
                Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
        
        End Select
        DoEvents
        'MSFlexGrid1.Col = 0
        'MSFlexGrid1.Row = i
        'i = 0
         If MSFlexGrid1.Rows > 2 Then
           MSFlexGrid1.RemoveItem i
        Else
            MSFlexGrid1.Clear
            Heading = "    |^Read|^Att|<To/From                                 |<Subject                                  |<Date                          "
            MSFlexGrid1.FormatString = Heading
            'MSFlexGrid1.Cols = 6
'MSFlexGrid1.Rows = 1
'MSFlexGrid1.FormatString = Headings
        End If
        
    Case Chr(8)    'Backspace
    Case "-"
    Case "."
    Case ","
    Case "/"
    Case "$"
    'Case "a" To "z"
    'Case "A" To "Z"
    Case Else
        KeyAscii = 0
End Select
InHere = False
Exit Sub

BadKeyEntry:
    frmMain.lblStatus = Err.Description & " - Can't complete action you requested"
    Beep
    Err.Clear
    InHere = False
End Sub

