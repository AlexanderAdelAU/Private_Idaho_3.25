Attribute VB_Name = "EMAIL"
Option Explicit

Private Function CheckConnection()
    Dim ReturnValue As Boolean
    Dim WaitAWhile As Variant
    Dim Req1 As Long
    Dim Response As String

    On Error GoTo BadConnection
    If PIForm(gActivePIInstance).IPInfo1.WinsockLoaded Then
        CheckConnection = True
    Else
        PIForm(gActivePIInstance).IPInfo1.WinsockLoaded = True
        CheckConnection = True
    End If
    DoEvents
    Exit Function
    
BadConnection:
    CheckConnection = False
    Beep
    MsgBox "Can't create connection - ", vbApplicationModal + vbCritical
    Err.Clear
End Function
Public Sub SendToOutBox(Optional EncryptedAttachments As Boolean)
Dim FileNum As Integer
Dim rs As Recordset
Dim FolderId As Long
Dim retStatus As Boolean
Dim AttFileName As String
Dim DBFileName As String
Dim i As Integer
Dim iAttachmentCount As Integer
   On Error GoTo SendToOutBoxMailError
    '
    'If a remailer is used then it is part of the message..
    '
    If gRemailerType = STANDARD_EMAIL Then
       'see if we have other X-headers
        Select Case MailHeader(0).ID
            Case "USENET"
                GetNewsHeaders
            Case "ExtendedHeaders"
                GetExtendedHeaders
        End Select
    End If
    If MailConnector.EmailAddress = "" Then
        MailConnector.EmailAddress = "Anonymous@mail"
    End If
       
    
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "[Folder] =" & "'" & "Outbox" & "'"
    If rs.NoMatch Then
        Err.Raise 10002, , "The database may be corrupt.  Inbox folder is missing from database."
    End If
    FolderId = rs("folder id")
    rs.Close
    
    Set rs = DB.OpenRecordset("messages", dbOpenDynaset)
    If Not rs.EOF Then rs.MoveLast
    
    rs.AddNew
    rs("Folder ID") = FolderId
    DBFileName = DatePart("d", Now())
    DBFileName = DBFileName & DatePart("m", Now())
    'DBFileName = DBFileName & Mid$(DatePart("yyyy", Now()), 3, 2)
    DBFileName = DBFileName & DatePart("yyyy", Now())
   ' Randomize rs("Message ID")
    'DBFileName = CStr(Int(100000000# * Rnd + 10000 * Rnd)) '+
    DBFileName = rs("Message ID") & DBFileName
    DBFileName = DBFileName & ".pim"
    
    rs("From") = MailConnector.EmailAddress
    rs("To") = PIForm(gActivePIInstance).txtTo.Text
    If PIForm(gActivePIInstance).txtSubject = "" Then
        rs("Subject") = "None"
    Else
        rs("Subject") = PIForm(gActivePIInstance).txtSubject
    End If
    rs("CC") = PIForm(gActivePIInstance).txtCC.Text
    rs("Date Sent") = Now() 'Format(Now(), "ddd d/mm/yy h:m ")
    DoEvents
    
    PIForm(gActivePIInstance).MIME1.Action = a_ResetData
    PIForm(gActivePIInstance).MIME1.PartCount = PIForm(gActivePIInstance).lvwAttachments.ListItems.Count + 1
    PIForm(gActivePIInstance).MIME1.PartEncoding(0) = 3  'No encoding 'pe_8Bit 'pe_QuotedPrintable
    PIForm(gActivePIInstance).MIME1.PartContentType(0) = "text/plain"
    'For i = 1 To PIForm(gActivePIInstance).MIME1.PartCount - 1
       ' PIForm(gActivePIInstance).MIME1.PartEncoding(i) = 2  'base 64 'pe_8Bit 'pe_QuotedPrintable
        'PIForm(gActivePIInstance).MIME1.PartContentType(i) = "image/gif"
    'Next
    PIForm(gActivePIInstance).MIME1.PartDecodedString(0) = PIForm(gActivePIInstance).MessageArea.Text
    
    'Note the path was stored in the temporary path location
    iAttachmentCount = PIForm(gActivePIInstance).MIME1.PartCount - 1
    For i = 1 To PIForm(gActivePIInstance).MIME1.PartCount - 1
        If EncryptedAttachments Then
            PIForm(gActivePIInstance).MIME1.PartDecodedFile(i) = TempPathLocation & StripFileName(PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)) '& ".asc"
        Else
            PIForm(gActivePIInstance).MIME1.PartDecodedFile(i) = TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
        End If
    Next i
    
    PIForm(gActivePIInstance).MIME1.Message = App.Path & "\mailbox\" & DBFileName
    PIForm(gActivePIInstance).MIME1.Action = 3 'a_EncodeToFile
    
       
    On Error GoTo SendToOutBoxMailError
    If PIForm(gActivePIInstance).MIME1.PartCount > 1 Then
        rs("Attachment") = True
    Else
        rs("Attachment") = False
    End If
    rs("MIME Message") = StripFileName(PIForm(gActivePIInstance).MIME1.Message)
    If MailHeader(0).Value = "" Then
         'Clipboard.Clear
    'Clipboard.SetText (PIForm(gActivePIInstance).MIME1.MessageHeaders)
        rs("MIME Message Header") = PIForm(gActivePIInstance).MIME1.MessageHeaders  '
    Else
        rs("MIME Message Header") = PIForm(gActivePIInstance).MIME1.MessageHeaders & MailHeader(0).Value & vbCrLf
    End If
       
    rs("Message Read") = False
    rs("Incoming Message") = False
    rs("Message Sent") = False
    rs.Update
    rs.Close
   
    'Kill all temporary files
    On Error Resume Next
    WipeFile TempPathLocation & "*.asc"
    For i = 1 To iAttachmentCount
        WipeFile TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
    Next i

Exit Sub

SendToOutBoxMailError:
    MsgBox " Message was not transferred to your out box. The following error occurred: " & Err.Description, vbApplicationModal + vbCritical, App.Title
    Err.Clear
    On Error Resume Next
    WipeFile TempPathLocation & "*.asc"
    For i = 1 To iAttachmentCount
        WipeFile TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
    Next i
End Sub
Public Sub EncryptToSelftandSendToSentBox(Optional EncryptedAttachments As Boolean)
Dim FileNum As Integer
Dim rs As Recordset
Dim FolderId As Long
Dim retStatus As Boolean
Dim AttFileName As String
Dim DBFileName As String
Dim sTextTemporaryFileName As String
Dim i As Integer
Dim iAttachmentCount As Integer
   On Error GoTo SendToOutBoxMailError
    
    'First encrypt
    gPGPKeyID = ReadProfile("PGP Options", "Default Key ID")
    If gPGPKeyID = "" Then
        MsgBox "The default Key ID is not set.  Set this in PGP ->Options.  A copy of this message will not be saved in the Sent Folders", vbApplicationModal + vbCritical, "PGP Default Key ID not set"
        Exit Sub
    End If
    vb2spgpContext.Initialise
    vb2spgpContext.ConventionalEncrypt = 0
    vb2spgpContext.KeyEncrypt = 1
    vb2spgpContext.TextMode = 0
    vb2spgpContext.Armor = 1
    vb2spgpContext.CryptKeyID = gPGPKeyID
    'Okay we need to copy this to file because we will leave the Message area untouched
    vb2spgpContext.FileIn = GetTemporaryFile
    PutFileText vb2spgpContext.FileIn, PIForm(gActivePIInstance).MessageArea.Text
    spgpEncryptfile
    sTextTemporaryFileName = vb2spgpContext.FileOut
    'PIForm(gActivePIInstance).MessageArea.Text = GetFileText(vb2spgpContext.FileOut)
    
    
    'spgpEncryptMessage
    EncryptAttachments (gPGPKeyID)
    
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "[Folder] =" & "'" & "Sent Items" & "'"
    If rs.NoMatch Then
        Err.Raise 10002, , "The database may be corrupt.  Inbox folder is missing from database."
    End If
    FolderId = rs("folder id")
    rs.Close
    
    Set rs = DB.OpenRecordset("messages", dbOpenDynaset)
    If Not rs.EOF Then rs.MoveLast
    
    rs.AddNew
    rs("Folder ID") = FolderId
    DBFileName = DatePart("d", Now())
    DBFileName = DBFileName & DatePart("m", Now())
    DBFileName = DBFileName & Mid$(DatePart("yyyy", Now()), 3, 2)
    Randomize rs("Message ID")
    DBFileName = CStr(Int(100000000# * Rnd + 10000 * Rnd)) '+  '(DBFileName & rs("Message ID"))))
    DBFileName = DBFileName & ".pim"
    
    rs("From") = MailConnector.EmailAddress
    rs("To") = PIForm(gActivePIInstance).txtTo.Text
    If PIForm(gActivePIInstance).txtSubject = "" Then
        rs("Subject") = "[EncryptToSelfCopy] - None"
    Else
        rs("Subject") = "[EncryptToSelfCopy] - " & PIForm(gActivePIInstance).txtSubject
    End If
    rs("CC") = PIForm(gActivePIInstance).txtCC.Text
    rs("Date Received") = Now()  'Format(Now(), "ddd d/mm/yy h:m ")
    DoEvents
    
    PIForm(gActivePIInstance).MIME1.Action = a_ResetData
    PIForm(gActivePIInstance).MIME1.PartCount = PIForm(gActivePIInstance).lvwAttachments.ListItems.Count + 1
    PIForm(gActivePIInstance).MIME1.PartEncoding(0) = 3  'No encoding 'pe_8Bit 'pe_QuotedPrintable
    PIForm(gActivePIInstance).MIME1.PartContentType(0) = "text/plain"
    'This is the temporary file that is holding the Message Body text
    PIForm(gActivePIInstance).MIME1.PartDecodedString(0) = GetFileText(sTextTemporaryFileName)
    
    'Note the path was stored in the temporary path location
    iAttachmentCount = PIForm(gActivePIInstance).MIME1.PartCount - 1
    For i = 1 To PIForm(gActivePIInstance).MIME1.PartCount - 1
        If EncryptedAttachments Then
            PIForm(gActivePIInstance).MIME1.PartDecodedFile(i) = TempPathLocation & StripFileName(PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)) '& ".asc"
        Else
            PIForm(gActivePIInstance).MIME1.PartDecodedFile(i) = vb2spgpContext.FileOut 'InStripFileName(PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i))
        End If
    Next i
    
    PIForm(gActivePIInstance).MIME1.Message = App.Path & "\mailbox\" & DBFileName
    PIForm(gActivePIInstance).MIME1.Action = 3 'a_EncodeToFile
    
       
    On Error GoTo SendToOutBoxMailError
    If PIForm(gActivePIInstance).MIME1.PartCount > 1 Then
        rs("Attachment") = True
    Else
        rs("Attachment") = False
    End If
    rs("MIME Message") = StripFileName(PIForm(gActivePIInstance).MIME1.Message)
    If MailHeader(0).Value = "" Then
         'Clipboard.Clear
    'Clipboard.SetText (PIForm(gActivePIInstance).MIME1.MessageHeaders)
        rs("MIME Message Header") = PIForm(gActivePIInstance).MIME1.MessageHeaders  '
    Else
        rs("MIME Message Header") = PIForm(gActivePIInstance).MIME1.MessageHeaders & MailHeader(0).Value & vbCrLf
    End If
       
    rs("Message Read") = False
    rs("Incoming Message") = False
    rs("Message Sent") = True
    rs.Update
    rs.Close
   
    'Kill all temporary files
    On Error Resume Next
   ' WipeFile TempPathLocation & "*.asc"
   ' For i = 1 To iAttachmentCount
   '     WipeFile TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
   ' Next i
KillTemporaryFiles
Exit Sub

SendToOutBoxMailError:
    MsgBox " Message was not transferred to your out box. The following error occurred: " & Err.Description, vbApplicationModal + vbCritical, App.Title
    Err.Clear
    On Error Resume Next
    WipeFile TempPathLocation & "*.asc"
    For i = 1 To iAttachmentCount
        WipeFile TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
    Next i
End Sub
Sub ReadMailMessage(EncodedFileName As String)
    Dim FileNum As Integer
    Dim MessageFileName As String
    Dim buffer As String
    
    On Error GoTo ReadMailError
   
    FileNum = FreeFile
    MessageFileName = App.Path & "\mailbox\" & EncodedFileName
    Open MessageFileName For Binary As #FileNum
    buffer = String(LOF(FileNum), " ")
    Get #FileNum, , buffer
    Close #FileNum
    PIForm(gActivePIInstance).MessageArea.Text = buffer
   
Exit Sub

ReadMailError:
    Close #FileNum
    HideStatus
    MsgBox Err.Description & "Error reading message.", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

Public Sub GetExtendedHeaders()
Dim i As Integer
 MailHeader(0).Value = ""
 For i = 6 To UBound(MailHeader)
        If Not MailHeader(i).Value = "" Then
            MailHeader(0).Value = MailHeader(0).Value & MailHeader(i).ID
            MailHeader(0).Value = MailHeader(0).Value & MailHeader(i).Value & vbLf
        End If
    Next
End Sub

Public Sub GetNewsHeaders()
Dim i As Integer
 
MailHeader(0).Value = ""
For i = 1 To 5
        If Not MailHeader(i).Value = "" Then
            MailHeader(0).Value = MailHeader(0).Value & MailHeader(i).ID
            MailHeader(0).Value = MailHeader(0).Value & IIf(i = 3, "<" & MailHeader(i).Value & ">", MailHeader(i).Value) & vbLf
        End If
    Next
    'MailHeader(0).Value = MailHeader(0).Value & vbCrLf
End Sub

Public Sub EncryptAttachments(KeyID As String)
Dim i As Integer
Dim J As Integer
Dim lListItem As ListItem
Dim AttachmentCount As Integer
Dim bEncryptToSelf As String
Static bAttachmentsEncrypted As Boolean

On Error Resume Next
    'Check if we have already encrypted the attachement
bEncryptToSelf = ReadProfile("PGP Options", "EncryptToSelf")
If bEncryptToSelf = "true" Then
    If bAttachmentsEncrypted Then
        bAttachmentsEncrypted = False
        Exit Sub
    End If
    bAttachmentsEncrypted = True
End If
    AttachmentCount = PIForm(gActivePIInstance).lvwAttachments.ListItems.Count
    If AttachmentCount = 0 Then Exit Sub
          
    'Need to check here for conventional encryption....
    For i = 1 To AttachmentCount
        vb2spgpContext.Initialise
        vb2spgpContext.Armor = 1
        If PIForm(gActivePIInstance).mEncryptAttachmentWithKey.Checked Then
           ' On Error Resume Next
            
            PIForm(gActivePIInstance).ShowStatus 1, "Encrypting " & TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
            'If the file already exists get rid of it...
            vb2spgpContext.ConventionalEncrypt = 0
            vb2spgpContext.CryptKeyID = KeyID
            vb2spgpContext.FileIn = TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
            vb2spgpContext.KeyEncrypt = 1
           ' vb2spgpContext.FileOut = TempPathLocation & StripFileName(PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)) & ".asc"
            If spgpEncryptfile Then
                'Remove the temporary files
                Kill TempPathLocation & PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
                PIForm(gActivePIInstance).StatusBar = "Attachment saved as " & vb2spgpContext.FileOut
                PIForm(gActivePIInstance).lvwAttachments.ListItems.Remove (i)
                J = PIForm(gActivePIInstance).GetFileIconNum(StripFileName(vb2spgpContext.FileOut))
                Set lListItem = PIForm(gActivePIInstance).lvwAttachments.ListItems.Add(, , StripFileName(vb2spgpContext.FileOut), J, J)
            Else
                Err.Raise 10023, , "Error encrypting attachment."
            End If
        Else
            'PIForm(gActivePIInstance).mConventionallyEncryptAttachment.Checked
             'On Error Resume Next
             'If the file already exists get rid of it...
             vb2spgpContext.ConventionalEncrypt = 1
             vb2spgpContext.FileIn = PIForm(gActivePIInstance).lvwAttachments.ListItems.Item(i)
             If spgpEncryptfile Then
                PIForm(gActivePIInstance).StatusBar = "Attachment saved as " & vb2spgpContext.FileOut
                PIForm(gActivePIInstance).lvwAttachments.ListItems.Remove (i)
                J = PIForm(gActivePIInstance).GetFileIconNum(StripFileName(vb2spgpContext.FileOut))
                Set lListItem = PIForm(gActivePIInstance).lvwAttachments.ListItems.Add(, , vb2spgpContext.FileOut, J, J)
             Else
                Err.Raise 10023, , "Error encrypting file."
             End If
        End If
    Next

End Sub

Public Function GenerateCommaSeparatedList(sList As String) As String
Dim bDone As Boolean
Dim lPos As Long
Dim l As Long

l = 1
GenerateCommaSeparatedList = sList
Do Until bDone = True
    lPos = InStr(l, GenerateCommaSeparatedList, ";")
    If lPos = 0 Then
        bDone = True
    Else
        Mid(GenerateCommaSeparatedList, lPos, 1) = ","
        l = lPos + 1
    End If
Loop
End Function
