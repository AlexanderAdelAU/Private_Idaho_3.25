VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNymsList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nyms Management"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Nym"
      Height          =   405
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   5160
      Width           =   2145
   End
   Begin VB.CheckBox chkEncrypt 
      Caption         =   "Encrypt using my default key!"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   405
      Index           =   0
      Left            =   6360
      TabIndex        =   2
      Top             =   5160
      Width           =   2145
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Nyms List.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   8265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Nyms List.frx":00E5
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8265
   End
End
Attribute VB_Name = "frmNymsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkEncrypt_Click()
MsgBox "Sorry not yet implemented", vbApplicationModal + vbExclamation, "Encrypt Nyms"
Exit Sub
If chkEncrypt.Value = vbChecked Then
    Command1(0).Caption = "Encrypt and Save!"
Else
    Command1(0).Caption = "Done"
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
    Case 1
        DeleteNym
End Select
End Sub

Private Sub Form_Load()
InitialiseGrid
FillGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmNymlist = Nothing
End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then DeleteNym
End Sub
Private Sub DeleteNym()
Dim i As Integer
Dim Heading As String
Dim iResponse As Integer
Static InHere As Boolean
Dim Sqlq As String
Dim rs As Recordset

If InHere Then Exit Sub

InHere = True
On Error GoTo BadKeyEntry:

Sqlq = "Select DISTINCTROW * FROM Nyms ORDER BY Nyms.[Nym Email];"
Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)

If rs.EOF Then
    Beep
    InHere = False
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.MoveFirst
End If
For i = 1 To MSFlexGrid1.Row - 1
    rs.MoveNext
Next

iResponse = MsgBox("Do you wish to permanently delete this Nym?", vbYesNo + vbQuestion + vbApplicationModal, "Delete message")
If iResponse = vbYes Then
    rs.Delete
    MSFlexGrid1.RemoveItem MSFlexGrid1.Row
End If


InHere = False
rs.Close
Set rs = Nothing
Exit Sub

BadKeyEntry:
    MsgBox "An error has occured.  Error description is: " & Err.Description, vbApplicationModal + vbCritical
    Err.Clear
    rs.Close
    InHere = False
End Sub

Private Sub InitialiseGrid()
Dim Headings As String
MSFlexGrid1.Clear
Headings = "    |<Nym Full Name           |<Nym Email                     |<Nym Server                 |<Server and Remailer Keys              "
MSFlexGrid1.Cols = 5
MSFlexGrid1.Rows = 1
MSFlexGrid1.FormatString = Headings
End Sub

Public Sub FillGrid()
Dim rs As Recordset
Dim Headings As String
Dim Sqlq As String
Dim RowIndex As Integer
Dim ItemString As String


On Error GoTo BadKeyEntry:

Sqlq = "Select DISTINCTROW * FROM Nyms "
Sqlq = Sqlq & "ORDER BY Nyms.[Nym Email];"
Set rs = DB.OpenRecordset(Sqlq, dbOpenDynaset)
If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Exit Sub
End If
'Now fill the message area
RowIndex = 1

While Not rs.EOF
    MSFlexGrid1.Col = 1
    ItemString = "" & vbTab & rs("Nym Full Name") & vbTab & rs("Nym Email") & vbTab & rs("Nym Server") & vbTab & rs("Nym Passphrases")
    MSFlexGrid1.AddItem ItemString, RowIndex
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = RowIndex
   
    MSFlexGrid1.Col = 2
    
    RowIndex = RowIndex + 1
    rs.MoveNext
Wend
MSFlexGrid1.FormatString = Headings
rs.Close
Set rs = Nothing
MSFlexGrid1.Row = 1
MSFlexGrid1.RowSel = 1

Exit Sub

BadKeyEntry:
    MsgBox "An error has occured.  Error description is: " & Err.Description, vbApplicationModal + vbCritical
    Err.Clear
    rs.Close
    InHere = False
End Sub
