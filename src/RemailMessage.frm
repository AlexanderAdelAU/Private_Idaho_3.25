VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Private Sub RemailMessage(TheMessage As String)
Dim FileNum As Integer
    Dim FileNum2 As Integer
    Dim IdxFile As Integer
    Dim IdxLen As Long
    Dim Buffer As String
    Dim BoundaryName As String
    Dim cmd As String
    Dim TempInt As Integer
    'Dim TempStr as String
    Dim i As Integer
    Dim ClipText As String
    Dim Cyphertext As String
    Dim Textline As String
    
    
    On Error GoTo RemailMessError
        
    gMessageRecord.Contents = TheMessage
   ' gMessageRecord.zap = 0
    gMessageRecord.PGP = 1
    'gMessageRecord.Read = 0
    gMessageRecord.SentDate = GetMsgSent(TheMessage)
    gMessageRecord.Sender = GetMsgSender(TheMessage)
    gMessageRecord.Subject = GetMsgSubject(TheMessage)
    gMessageRecord.ATTACHMENT = GetMsgAttachment(TheMessage)
    
'Decrypt the Message'
   If frmMain.MessageArea.Text = "" Then Exit Sub
        ClipText = frmMain.MessageArea.Text
        FileNum = FreeFile
        If InStr(1, gPGPFile, ":") = 0 Then
            Open gPGPPath + "\" + gPGPFile + ".out" For Output As FileNum
        Else
            Open gPGPFile + ".out" For Output As FileNum
        End If
        Print #FileNum, ClipText
        Close #FileNum
        
        On Error Resume Next
        Kill gPGPPath & "\" & gPGPFile
        gPassPhrase = "made glorious summer 181147"
        cmd = gPGPPath & "\PGP " & gPGPPath & "\" & gPGPFile & ".out " & " -o " & gPGPPath & "\" & gPGPFile & " -z " & """" & gPassPhrase & """"
        CheckLen (cmd)
        ExecCmd (cmd)
        
        FileNum = FreeFile
        If InStr(1, gPGPFile, ":") = 0 Then
            Open gPGPPath + "\" + gPGPFile For Input As FileNum
        Else
            Open gPGPFile For Input As FileNum
        End If
        While Not EOF(FileNum)
            Line Input #FileNum, Textline
            Cyphertext = Cyphertext & Textline & vbCrLf
        Wend
        Close #FileNum
        If gPOPState = POPDECRYPT Then
            Cyphertext = "Subject: " + gMessageRecord.Subject + gCRLF + gCRLF + Cyphertext
            Cyphertext = "Sent: " + gMessageRecord.SentDate + gCRLF + Cyphertext
            Cyphertext = "From: " + gMessageRecord.Sender + gCRLF + Cyphertext
            gPOPState = 0
        End If
        frmMain.MessageArea.Text = Cyphertext
        'If InStr(1, gPGPFile, ":") = 0 Then
         '   Kill gPGPPath + "\" + gPGPFile
        'Else
         '   Kill gPGPFile
        'End If
        'If InStr(1, gPGPFile, ":") = 0 Then
        '    WipeFile (gPGPPath + "\" + gPGPFile)
        'Else
        '    WipeFile (gPGPFile)
        'End If
        

'Look for the remailer info
    Dim StartPosition As Long
    Dim RemailAddress As String
    
    FileNum = FreeFile
    If InStr(1, gPGPFile, ":") = 0 Then
        Open gPGPPath + "\" + gPGPFile For Input As FileNum
    Else
        Open gPGPFile For Input As FileNum
    End If
    While Not EOF(FileNum)
        Line Input #FileNum, Textline
        StartPosition = InStr(1, Textline, gRequestRemailingTo)
        If StartPosition Then
            StartPosition = InStr(1, Textline, ":")
            RemailAddress = Mid(Textline, StartPosition + 1)
            RemailAddress = Trim(RemailAddress)
            Set EOF(FileNum) = True
        End If
    Wend
    Close #FileNum
    
           
       
    Exit Sub

RemailMessError:
    MsgBox Err.Description & " (WriteMessageRecord)"
    Err.Clear
End Sub
