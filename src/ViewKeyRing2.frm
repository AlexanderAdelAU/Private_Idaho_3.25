VERSION 5.00
Begin VB.Form frmKeyRing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PGP Keyring"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "ViewKeyRing2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   4320
      TabIndex        =   19
      Top             =   5820
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   1620
      TabIndex        =   18
      Top             =   5820
      Width           =   1812
   End
   Begin VB.CheckBox ckExportPrivate 
      Caption         =   "Export Signatures && Private Keys"
      Height          =   195
      Left            =   3900
      TabIndex        =   15
      Top             =   3900
      Width           =   3672
   End
   Begin VB.Frame Frame1 
      Caption         =   "Key Properties"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   4140
      Width           =   7512
      Begin VB.TextBox edCreated 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox edValidity 
         Height          =   285
         Left            =   6480
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edTrust 
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edPrivate 
         Height          =   285
         Left            =   6480
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox edBits 
         Height          =   285
         Left            =   2940
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edAlg 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox edFingerprint 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblValidity 
         Caption         =   "Validity:"
         Height          =   252
         Left            =   5700
         TabIndex        =   14
         Top             =   1140
         Width           =   732
      End
      Begin VB.Label lblTrust 
         Caption         =   "Trust:"
         Height          =   252
         Left            =   3900
         TabIndex        =   13
         Top             =   1140
         Width           =   552
      End
      Begin VB.Label lblCreated 
         Caption         =   "Created:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   852
      End
      Begin VB.Label lblPrivate 
         Caption         =   "Private or Public?"
         Height          =   252
         Left            =   4980
         TabIndex        =   8
         Top             =   420
         Width           =   1392
      End
      Begin VB.Label lblKeySize 
         Caption         =   "Key Size:"
         Height          =   252
         Left            =   2040
         TabIndex        =   6
         Top             =   1140
         Width           =   852
      End
      Begin VB.Label lblAlgorithm 
         Caption         =   "Algorithm:"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   852
      End
      Begin VB.Label lblFingerprint 
         Caption         =   "Fingerprint:"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   912
      End
   End
   Begin VB.ListBox List1 
      Height          =   3435
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   7512
   End
   Begin VB.Label lblPrefer 
      Alignment       =   2  'Center
      Caption         =   "(Click checkboxes to indicate preferred keys for encryption)"
      Height          =   252
      Left            =   780
      TabIndex        =   20
      Top             =   3660
      Width           =   5892
   End
   Begin VB.Label lblExport 
      Alignment       =   2  'Center
      Caption         =   "(Double-click on key to export)"
      Height          =   252
      Left            =   360
      TabIndex        =   16
      Top             =   3900
      Width           =   3432
   End
End
Attribute VB_Name = "frmKeyRing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 1999 Potato Software.  All rights reserved.
'See MAIN.BAS for License.

Dim DataChanged As Boolean
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    If DataChanged Then
        PubKeyX = 0: ReDim PubKeys(List1.ListCount - 1)
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then PubKeys(PubKeyX) = List1.List(i): PubKeyX = PubKeyX + 1
        Next i
        WriteData
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long, j As Long
    
    Me.Left = Val(RData(34))
    Me.Top = Val(RData(35))
    
    If Lang <> "en" Then
        Me.Caption = LMsg(317)
        lblPrefer.Caption = LMsg(318)
        lblExport.Caption = LMsg(319)
        ckExportPrivate.Caption = LMsg(320)
        Frame1.Caption = LMsg(321)
        lblFingerprint.Caption = LMsg(322)
        lblCreated.Caption = LMsg(323)
        lblAlgorithm.Caption = LMsg(324)
        lblPrivate.Caption = LMsg(325)
        lblKeySize.Caption = LMsg(326)
        lblTrust.Caption = LMsg(327)
        lblValidity.Caption = LMsg(328)
        cmdCancel.Caption = LMsg(279)
    End If
    
    For i = 0 To KeyArrayCount - 1
        If KeyArray(i).keyid <> "" Then
            List1.AddItem KeyArray(i).keyid & Chr(9) & KeyArray(i).UserID
            For j = 0 To PubKeyX - 1
                If PubKeys(j) = KeyArray(i).keyid & Chr(9) & KeyArray(i).UserID Then List1.Selected(List1.NewIndex) = True: Exit For
            Next j
        End If
    Next i
    If List1.ListCount <> 0 Then List1.ListIndex = 0
    DataChanged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RData(34) = Str(Me.Left)
    RData(35) = Str(Me.Top)
End Sub

' extract the Key ID from the selected list item
' call keyprops to get the key's properties
Private Sub List1_Click()
Dim i As Long
Dim BufferIn, BufferOut As String * 8192  '1024
Dim Key As TKey_Data

DataChanged = True
BufferOut = BufferOut & Chr(0)

' first 10 characters will be the key id (0x12345678)
BufferIn = Mid(List1.List(List1.ListIndex), 1, 10)

' keyprops takes either key id(s) or user id(s)
' and returns the key's properties
i = spgpKeyProps(BufferIn, BufferOut, 8192)  '1024

' parse the returned property-string into a TKey_Data record
Key = ParseKeyData(BufferOut)

  edFingerprint.Text = Key.Fingerprint
  edAlg.Text = Mid(Key.KeyAlgorithm, InStr(Key.KeyAlgorithm, "_") + 1)
  'XXXXXXXXXXX edAlg.Text = Mid(Key.KeyAlgorithm, 14)
  edBits.Text = Key.Bits
  
  If Key.Private = True Then
    edPrivate = "Private"
  Else
    edPrivate = "Public"
  End If
  
  edCreated.Text = Key.DateTimeStr
  edTrust.Text = Mid(Key.Trust, InStr(Key.Trust, "_") + 1)
  edValidity.Text = Mid(Key.Validity, InStr(Key.Validity, "_") + 1)
'XXXXXXXXXXX  edTrust.Text = Mid(Key.Trust, 10)
'XXXXXXXXXXX  edValidity.Text = Mid(Key.Validity, 13)
  
End Sub

Private Sub List1_DblClick()
    Dim i As Long
    Dim BufferIn As String * 1024
    Dim BufferOut As String * 8192  '1024
    Dim Key As TKey_Data
    
    If List1.ListIndex = -1 Then Exit Sub
    List1.Selected(List1.ListIndex) = Not List1.Selected(List1.ListIndex)
    BufferOut = BufferOut & Chr(0)
    
    ' first 10 characters will be the key id (0x12345678)
    BufferIn = Mid(List1.List(List1.ListIndex), 1, 10)
    
    ' keyexport takes either key id(s) or user id(s)
    ' and returns the key
    i = spgpKeyExport(BufferIn, BufferOut, 8192, ckExportPrivate.Value, 1)
    
    If i = 0 Then
        Clipboard.Clear
        Clipboard.SetText BufferOut
        MsgBox LMsg(315) + ":" + vbCr + vbCr + List1.List(List1.ListIndex), vbOKOnly, LMsg(316)
    Else
        MsgBox LMsg(0), vbCritical + vbOKOnly, LMsg(316)
    End If
End Sub
