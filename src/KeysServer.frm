VERSION 5.00
Begin VB.Form frmKeyServer 
   Caption         =   "Key Server List"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSubmit 
      Height          =   315
      Index           =   2
      Left            =   420
      TabIndex        =   10
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtSubmit 
      Height          =   315
      Index           =   1
      Left            =   420
      TabIndex        =   9
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton btnOkay 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   4440
      Width           =   795
   End
   Begin VB.TextBox txtSubmit 
      Height          =   315
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   2
      Left            =   420
      TabIndex        =   3
      Top             =   1680
      Width           =   6735
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   1320
      Width           =   6735
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "E-Mail address of server to which key will be submmited."
      Height          =   315
      Index           =   2
      Left            =   420
      TabIndex        =   8
      Top             =   2460
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "URL from which to obtain key: (Note includes search syntax)"
      Height          =   315
      Index           =   1
      Left            =   420
      TabIndex        =   7
      Top             =   660
      Width           =   6435
   End
   Begin VB.Label Label1 
      Caption         =   "Current List of Key Server Addresses.  Edit text or clear to delete from the list."
      Height          =   315
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   6435
   End
End
Attribute VB_Name = "frmKeyServer"
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
Open App.Path + "\KeyServerURL.TXT" For Output As FileNum
For i = 0 To txtURL.Count - 1
    If Not txtURL(i) = "" Then Print #FileNum, txtURL(i)
Next
Close FileNum
'
FileNum = FreeFile
Open App.Path + "\SubmitKeyAddr.TXT" For Output As FileNum
For i = 0 To txtSubmit.Count - 1
    If Not txtSubmit(i) = "" Then Print #FileNum, txtSubmit(i)
Next
Close FileNum

Unload Me
End Sub

Private Sub Form_Load()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

FileNum = FreeFile
If iFileExists(App.Path + "\KeyServerURL.TXT") Then
    Open App.Path + "\KeyServerURL.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            txtURL(i) = tmpString
            If i = 5 Then Exit Do
            i = i + 1
        Loop
    Else
       txtURL(0) = gGetKeyURL
    End If
    Close FileNum
Else
    txtURL(0) = gGetKeyURL
End If

' Now do the email server address
FileNum = FreeFile
If iFileExists(App.Path + "\SubmitKeyAddr.TXT") Then
    Open App.Path + "\SubmitKeyAddr.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            txtSubmit(i) = tmpString
            If i = 5 Then Exit Do
            i = i + 1
        Loop
    Else
       txtSubmit(0) = gSubKeyURL
    End If
    Close FileNum
Else
    txtSubmit(0) = gSubKeyURL
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSelectKeyServer = Nothing
End Sub
