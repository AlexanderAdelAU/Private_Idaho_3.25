VERSION 5.00
Begin VB.Form frmUpgrade 
   Caption         =   "PI Database Upgrade"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox lbldrive 
      Height          =   345
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upgrade Private Idaho to 3.2.20"
      Height          =   585
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   2925
   End
   Begin VB.Label Label2 
      Caption         =   "Private Idaho directory: "
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   1710
      Width           =   1755
   End
   Begin VB.Label lblstatus 
      Caption         =   "Label2"
      Height          =   585
      Left            =   450
      TabIndex        =   2
      Top             =   3090
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "This upgrade will create copy the upgraded file to the directory specified."
      Height          =   1245
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim tNewTable As TableDef
Dim tField As Field
Dim iNewIndex As Index
Dim rs As Recordset
Dim gNym As String
Dim gfullNym As String
Dim gNymServer As String
    
On Error GoTo CantCreate
''''''''''
GoTo Upgrade_exe
'''''''''''
lblstatus = ""
If Command1.Caption = "Exit" Then Unload Me
Set db = DBEngine.Workspaces(0).OpenDatabase((lbldrive _
         & IIf(Right$(lbldrive, 1) <> "\", "\", "") _
         & "PI32PostOffice.MDB"))
Set tNewTable = db.CreateTableDef("Nyms")
With tNewTable
   .Fields.Append .CreateField("Nym ID", dbLong)
   .Fields.Append .CreateField("Nym Email", dbText, 32)
   .Fields.Append .CreateField("Nym Full Name", dbText, 32)
   .Fields.Append .CreateField("Nym Server", dbText, 32)
   .Fields.Append .CreateField("Nym Passphrases", dbText, 255)
   
End With
tNewTable.Fields("Nym ID").Attributes = dbAutoIncrField
tNewTable.Fields("Nym Email").DefaultValue = ""
tNewTable.Fields("Nym Full Name").DefaultValue = ""
tNewTable.Fields("Nym Server").DefaultValue = ""
tNewTable.Fields("Nym Passphrases").DefaultValue = ""
db.TableDefs.Append tNewTable

'Create Index
Set iNewIndex = tNewTable.CreateIndex("PrimaryKey")

With iNewIndex
    .Fields.Append .CreateField("Nym ID")
    .Primary = True
    .Unique = True
End With
tNewTable.Indexes.Append iNewIndex
'Set tField = tNewTable.CreateField("Nym ID")
'tField.Attributes = dbAutoIncrField
CantCreate:
db.Close


On Error GoTo BadUpgrade
Set db = DBEngine.Workspaces(0).OpenDatabase((lbldrive _
         & IIf(Right$(lbldrive, 1) <> "\", "\", "") _
         & "PI32PostOffice.MDB"))
Set rs = db.OpenRecordset("Nyms", dbOpenDynaset)

If iFileExists(lbldrive + "\NYMS.TXT") Then
    FileNum = FreeFile
    Open lbldrive + "\NYMS.TXT" For Input As FileNum
    Do Until EOF(FileNum)
        rs.AddNew
        Line Input #FileNum, gNym
        rs("Nym Email") = IIf(gNym = "", "Empty", gNym)
        Line Input #FileNum, gfullNym
        rs("Nym Full Name") = IIf(gfullNym = "", "Empty", gfullNym)
        Line Input #FileNum, gNymServer
        rs("Nym Server") = IIf(gNymServer = "", "Empty", gNymServer)
        rs("Nym Passphrases") = "Empty"
        rs.Update
    Loop
    Close FileNum
    Name lbldrive + "\NYMS.TXT" As lbldrive + "\nyms(old).TXT"
    End If
rs.Close
db.Close
Upgrade_exe:
'Copy updated exe across
FileCopy App.Path & "\pi32.exe", lbldrive & "\pi32.exe"
'Command1.Caption = "Exit"
DoEvents
lblstatus = "Upgrade was successful. "

Exit Sub
BadUpgrade:
    Beep
    MsgBox "Upgrade was not successful.  Error given was: " & Err.Description, vbApplicationModal + vbExclamation, App.Title
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblstatus = ""
lbldrive = "C:\pi32"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUpgrade = Nothing
End
End Sub
Function iFileExists(ByVal sFileName As String) As Boolean
Dim i As Integer
Dim FileNum As Integer
    Err.Clear
    On Error Resume Next
    i = Len(Dir$(sFileName))
    If Err.Number <> 0 Or i = 0 Then
        iFileExists = False
    Else
        If InStr(1, sFileName, "*", vbTextCompare) = 0 Then
            FileNum = FreeFile
     '   Open Filenum For Input As Binary
            Open sFileName For Input As FileNum
            If EOF(FileNum) Then
                iFileExists = False
            Else
                iFileExists = True
            End If
            Close FileNum
        Else
            iFileExists = True
        End If
    End If
End Function
