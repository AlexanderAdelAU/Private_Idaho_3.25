VERSION 5.00
Begin VB.Form frmKeysURL 
   Caption         =   "Keys URL List"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3000
      Width           =   795
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3420
      TabIndex        =   6
      Top             =   3000
      Width           =   795
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   4
      Left            =   420
      TabIndex        =   5
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   3
      Left            =   420
      TabIndex        =   4
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   2
      Left            =   420
      TabIndex        =   3
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Current List of Key Server Addresses.  Edit text or clear to delete from the list."
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   420
      Width           =   4815
   End
End
Attribute VB_Name = "frmKeysURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOkay_Click()
Dim FileNum As Integer
Dim i As Integer

FileNum = FreeFile
Open App.Path + "\keysURL.TXT" For Output As FileNum
For i = 0 To txtURL.Count - 1
    If Not txtURL(i) = "" Then Print #FileNum, txtURL(i)
Next
Close FileNum
Unload Me
End Sub

Private Sub Form_Load()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

FileNum = FreeFile

If iFileExists(App.Path + "\keysURL.TXT") Then
    Open App.Path + "\keysURL.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            txtURL(i) = tmpString
            If i = 5 Then Exit Do
            i = i + 1
        Loop
    Else
       txtURL(0) = gPGPKeysURL
    End If
    Close FileNum
Else
    txtURL(0) = gPGPKeysURL
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmKeysURL = Nothing
End Sub
