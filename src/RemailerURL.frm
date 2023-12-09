VERSION 5.00
Begin VB.Form frmRemailerURL 
   Caption         =   "Remailer URL List"
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
   Begin VB.Label lblListType 
      Caption         =   "Label2"
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   630
      Width           =   3285
   End
   Begin VB.Label Label1 
      Caption         =   "Current List of URLs from which remailer list and information can be obtained.  Edit or clear to delete."
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   90
      Width           =   4815
   End
End
Attribute VB_Name = "frmRemailerURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOkay_Click()
Select Case gRemailerTypeURL
    Case 0
        UpdateRemailerFile
    Case 1
        UpdateMixListFile
    Case 2
        UpdateMixType2File
    Case 3
        UpdateMixPubRingsFile
End Select
Unload Me
End Sub

Private Sub Form_Load()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

Select Case gRemailerTypeURL
    Case 0
        LoadRemailerURLs
    Case 1
        LoadMixListURL
    Case 2
        LoadMixType2URL
    Case 3
        LoadMixPubRingURL
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmRemailerURL = Nothing
End Sub

Public Sub LoadRemailerURLs()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

FileNum = FreeFile

If iFileExists(App.Path + "\InfoURL.TXT") Then
    Open App.Path + "\infoURL.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            If i = 5 Then Exit Do
            txtURL(i) = tmpString
            i = i + 1
        Loop
    Else
       txtURL(0) = gRemailerInfoURL
    End If
    Close FileNum
Else
    txtURL(0) = gRemailerInfoURL
End If
End Sub

Public Sub LoadMixListURL()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

FileNum = FreeFile

If iFileExists(App.Path + "\MixListURL.TXT") Then
    Open App.Path + "\MixListURL.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            txtURL(i) = tmpString
            If i = 5 Then Exit Do
            i = i + 1
        Loop
    Else
       txtURL(0) = gMixListURL
    End If
    Close FileNum
Else
    txtURL(0) = gMixListURL
End If
End Sub

Public Sub LoadMixType2URL()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

FileNum = FreeFile

If iFileExists(App.Path + "\MixType2URL.TXT") Then
    Open App.Path + "\MixType2URL.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            txtURL(i) = tmpString
            If i = 5 Then Exit Do
            i = i + 1
        Loop
    Else
       txtURL(0) = gMixType2URL
    End If
    Close FileNum
Else
    txtURL(0) = gMixType2URL
End If
End Sub

Public Sub LoadMixPubRingURL()
Dim FileNum As Integer
Dim i As Integer
Dim tmpString As String

FileNum = FreeFile

If iFileExists(App.Path + "\MixPubRingURL.TXT") Then
    Open App.Path + "\MixPubRingURL.TXT" For Input As FileNum
    i = 0
    If Not EOF(FileNum) Then
        Do Until EOF(FileNum)
            Line Input #FileNum, tmpString
            txtURL(i) = tmpString
            If i = 5 Then Exit Do
            i = i + 1
        Loop
    Else
       txtURL(0) = gMixPubRingURL
    End If
    Close FileNum
Else
    txtURL(0) = gMixPubRingURL
End If
End Sub

Public Sub UpdateRemailerFile()
Dim FileNum As Integer
Dim i As Integer
Dim bRes As Boolean
Dim List As String
'ires=putfFiletext(App.Path + "\InfoURL.TXT"
'FileNum = FreeFile
'Open App.Path + "\InfoURL.TXT" For Output As FileNum
List = ""
For i = 0 To txtURL.Count - 1
    If Not txtURL(i) = "" Then List = List & txtURL(i) & vbCrLf
Next
bRes = PutFileText(App.Path + "\InfoURL.TXT", List)
'Close FileNum
End Sub

Public Sub UpdateMixListFile()
Dim FileNum As Integer
Dim i As Integer
Dim bRes As Boolean
Dim List As String
'ires=putfFiletext(App.Path + "\InfoURL.TXT"
'FileNum = FreeFile
'Open App.Path + "\InfoURL.TXT" For Output As FileNum
List = ""
For i = 0 To txtURL.Count - 1
    If Not txtURL(i) = "" Then List = List & txtURL(i) & vbCrLf
Next
bRes = PutFileText(App.Path + "\MixListURL.TXT", List)
'Close FileNum

End Sub

Public Sub UpdateMixType2File()
Dim FileNum As Integer
Dim i As Integer
Dim bRes As Boolean
Dim List As String
'ires=putfFiletext(App.Path + "\InfoURL.TXT"
'FileNum = FreeFile
'Open App.Path + "\InfoURL.TXT" For Output As FileNum
List = ""
For i = 0 To txtURL.Count - 1
    If Not txtURL(i) = "" Then List = List & txtURL(i) & vbCrLf
Next
bRes = PutFileText(App.Path + "\MixType2URL.TXT", List)
'Close FileNum


End Sub

Public Sub UpdateMixPubRingsFile()
Dim FileNum As Integer
Dim i As Integer
Dim bRes As Boolean
Dim List As String
'ires=putfFiletext(App.Path + "\InfoURL.TXT"
'FileNum = FreeFile
'Open App.Path + "\InfoURL.TXT" For Output As FileNum
List = ""
For i = 0 To txtURL.Count - 1
    If Not txtURL(i) = "" Then List = List & txtURL(i) & vbCrLf
Next
bRes = PutFileText(App.Path + "\MixPubRingURL.TXT", List)
'Close FileNum
End Sub
