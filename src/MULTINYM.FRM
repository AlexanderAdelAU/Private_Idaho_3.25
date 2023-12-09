VERSION 5.00
Begin VB.Form frmMultiNyms 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nyms"
   ClientHeight    =   3510
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5610
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form27"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   5610
   Begin VB.CheckBox chkServer 
      Caption         =   "Remove the Nym from the Server?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2985
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
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
      Left            =   4320
      TabIndex        =   3
      Top             =   2580
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
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
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   2580
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
      Height          =   1785
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select a nym to use."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMultiNyms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    
    gCancelAction = False
    Nym.NymState = gNymState
    Nym.ListIndex = List1.ListIndex
    Unload Me
    

End Sub

Private Sub Command2_Click()
    gCancelAction = True
    Unload frmMultiNyms
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    Dim FileNum As Integer
    Dim Item As String
    Dim Item1 As String
    Dim Win As New CWindow
    
    Win.OnTop(Me) = True
    On Error GoTo NymLoadError
    If gNymState = gNYMDEL Then
        Label1(0).Caption = "Select a Nym to delete."
        chkServer.Enabled = True
    End If
    Set rs = DB.OpenRecordset("NYMS", dbOpenDynaset)
    If rs.EOF Then PIForm(gActivePIInstance).ShowStatus 1, "No Nyms in database...."
    Do While Not rs.EOF
        Item1 = rs("Nym Email")
        Item = rs("Nym Full Name")
        If PIForm(gActivePIInstance).IsNewNym(Item1) Then
                List1.AddItem Item & " - " & Item1
        Else
            List1.AddItem Item & " - " & Item1
        End If
        rs.MoveNext
    Loop
    rs.Close
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
   
    Exit Sub

NymLoadError:
    MsgBox Err.Description + "MultiNym.Load. The please ensure that the database is uptodate."
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMultiNyms = Nothing
End Sub

