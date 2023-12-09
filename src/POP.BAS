Attribute VB_Name = "POP"
 Option Explicit
 
Function GetMsgSender(TheMessage As String) As String
    Dim TmpSender As String
    Dim i As Integer
    Dim tmpstr As String
    TmpSender = ""
    i = InStr(1, TheMessage, "From: ")
    If i = 0 Then
        GetMsgSender = ""
        Exit Function
    End If
    i = i + 6
    Do Until tmpstr = vbLf Or tmpstr = vbCr
        tmpstr = Mid$(TheMessage, i, 1)
        TmpSender = TmpSender + tmpstr
        i = i + 1
    Loop
    GetMsgSender = LTrim$(Mid$(TmpSender, 1, Len(TmpSender) - 1))
End Function

Function GetMsgSent(TheMessage As String) As String
Dim TmpSent As String
Dim i As Integer
Dim tmpstr As String
    
    TmpSent = ""
    i = InStr(1, TheMessage, "Date: ")
    If i = 0 Then
        GetMsgSent = ""
        Exit Function
    End If
    i = i + 6
    Do Until tmpstr = vbLf Or tmpstr = vbCr
        tmpstr = Mid$(TheMessage, i, 1)
        TmpSent = TmpSent + tmpstr
        i = i + 1
    Loop
    'If Mid$(TmpSent$, Len(TmpSent), 1) = vbLf Then
     '   TmpSent$ = Mid$(TmpSent, 1, Len(TmpSent) - 1)
    'End If
    GetMsgSent = LTrim$(Mid$(TmpSent, 1, Len(TmpSent) - 1))
End Function

Function GetMsgSubject(TheMessage As String) As String
   
   Dim TmpSub As String
   Dim i As Integer
   Dim tmpstr As String
   
   
    TmpSub = ""
    i = InStr(1, TheMessage, "Subject: ")
    If i = 0 Then
        GetMsgSubject = ""
        Exit Function
    End If
    i = i + 9
    Do Until tmpstr = vbLf Or tmpstr = vbCr
        tmpstr = Mid$(TheMessage, i, 1)
        TmpSub = TmpSub + tmpstr
        i = i + 1
    Loop
    'If Mid$(TmpSub, Len(TmpSub$), 1) = vbLf Then
      '  TmpSub = Mid$(TmpSub, 1, Len(TmpSub) - 1)
    'End If
    GetMsgSubject = LTrim$(Mid$(TmpSub, 1, Len(TmpSub) - 1))
End Function


Function GetMsgAttachment(TheMessage As String) As Boolean
    Dim tmpstr As String
    Dim TmpSender As String
    Dim i As Integer
    
    'GetPOPTop (messageNum)
    If InStr(1, TheMessage, "multipart/mixed;") Then
        GetMsgAttachment = True
    Else
        GetMsgAttachment = False
    End If
End Function

Function InsertReplies(TheString As String) As String
    
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim i As Integer
    
    
    tmpstr = ">"
    For i = 1 To Len(TheString)
        tmpstr2 = Mid$(TheString, i, 1)
        tmpstr = tmpstr + tmpstr2
        If tmpstr2 = Chr$(10) Then
            tmpstr = tmpstr + ">"
        End If
    Next
    InsertReplies = tmpstr
End Function

Sub WriteMessageRecord()
    
    Dim rsAttachment As Recordset
    Dim AttachmentFileName As String
   Dim TempFile As String
   Dim sTimeZone As String
    Dim i As Integer
    Dim J As Integer
    Dim lTimeZone As Long
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim MessageFileName As String
    Dim lMessageID As Long
    Dim lFolderID As Long
    Dim retStatus As Boolean
    
    
    On Error GoTo WriteMessError
    'TempFile = App.Path & "\mailbox\_pitemp.tmp"
    gMessage = StripNulls(gMessage)
      
    gMessageRecord.PGP = 1
    'Write to datbase here....
    
    Set rs2 = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs2.FindFirst "[Folder] =" & "'" & "Inbox" & "'"
    If rs2.NoMatch Then
        Err.Raise 1002, , "Inbox folder missing from database."
    End If
    lFolderID = rs2("folder id")
    Set rs = DB.OpenRecordset("Messages", dbOpenDynaset)
    rs.AddNew
    MessageFileName = DatePart("d", Now())
    MessageFileName = MessageFileName & DatePart("m", Now())
    MessageFileName = MessageFileName & Mid$(DatePart("yyyy", Now()), 3, 2)
    MessageFileName = MessageFileName & rs("Message ID") & ".pim"
    retStatus = PutFileText(App.Path & "\mailbox\" & MessageFileName, gMessage)
    
    rs("Folder ID") = lFolderID
    rs("From") = Trim(StripFullName(gMessageRecord.From) & " <" & StripEMailAddress(gMessageRecord.From) & ">") 'gMessageRecord.From
    rs("To") = gMessageRecord.To
    rs("Reply To") = Trim(StripFullName(gMessageRecord.ReplyTo) & " " & StripEMailAddress(gMessageRecord.ReplyTo)) '[AC Fix
    'Get the date format and convert
    i = InStr(1, gMessageRecord.SentDate, ",")
    J = InStr(1, gMessageRecord.SentDate, "+")
    If J = 0 Then
        J = InStr(1, gMessageRecord.SentDate, "-") + 2
    End If
    If J = 0 Then J = Len(gMessageRecord.SentDate)
    
   'Now convert to Local Time
  '  sTimeZone = Mid(gMessageRecord.SentDate, J + 1, Len(gMessageRecord.SentDate))
  '  lTimeZone = CLng(Mid(sTimeZone, 1, 2)) * 60 + CLng(Mid(sTimeZone, 3, 2))
    rs("Date Received") = Now()  'gMessageRecord.SentDate 'Mid(gMessageRecord.SentDate, i + 2, J - i - 2)
    Dim date_parts() As String
    's = gMessageRecord.SentDate
    date_parts = Split(gMessageRecord.SentDate, ",")
    date_parts = Split(date_parts(1), "+")
   ' rs("Date Sent") = CDate(gMessageRecord.SentDate)
    rs("Date Sent") = CDate(date_parts(0)) 'StringToDate(gMessageRecord.SentDate, "mdy")
   ' rs("Date Received") = DateAdd("n", lTimeZone, rs("Date Received"))
 '  rs("Date Received") = gMessageRecord.SentDate
    rs("CC") = gMessageRecord.CC   '" "
    rs("Message Number") = gMessageRecord.MessageNumber
    rs("Message Size") = gMessageRecord.MessageSize
    rs("MIME Message Header") = gMessageRecord.Header
   ' Clipboard.SetText gMessageRecord.Header
    rs("MIME Message") = MessageFileName
    If Not InStr(1, gMessageRecord.Header, "boundary=") = 0 Then
        rs("Attachment") = True
    End If
   
    If gMessageRecord.Subject = "" Then gMessageRecord.Subject = "None"
    rs("Subject") = gMessageRecord.Subject
     
     'Added 3/5/03  - relay code
    'Relay syntax is in subject [Relay to: acameron@itech.net.au] Normal Subject
    '========================
    If Not InStr(1, gMessageRecord.Subject, "[Relay to:") = 0 Then
        'Strip email
        i = InStr(1, gMessageRecord.Subject, ": ")
        J = InStr(i, gMessageRecord.Subject, "]")
        'Now find the outbox
        rs2.FindFirst "[Folder] =" & "'" & "OutBox" & "'"
        If rs2.NoMatch Then
            Err.Raise 1002, , "Outbox folder missing from database."
        End If
        rs("To") = Mid(gMessageRecord.Subject, i + 1, J - i - 1)
        rs("Subject") = Mid(gMessageRecord.Subject, J + 1, Len(gMessageRecord.Subject))
        If rs("Subject") = "" Then rs("Subject") = "None"
        gMessageRecord.Subject = rs("Subject")
        gMessageRecord.To = rs("To")
        'Change to outbox
        lFolderID = rs2("folder id")
        rs("Folder ID") = rs2("folder id")
       ' rs("MIME Message Header") = ""
    End If
    
    '========================
   
    
    'This is crude but last thing is to strip out PGP message
    rs("Message Read") = False
    rs("Incoming Message") = True
    rs("Message Deleted") = False
       
    If gPGPVersion = PGP5x Then
        rs("Message Status") = spgpAnalyseMessage(gMessage)
    Else
        rs("Message Status") = PGPAnalyze_Unknown
    End If
    lMessageID = rs("Message ID")
    rs.Update
    rs.Close
    rs2.Close
    'Success so update contact database
    UpdateContactRecord (gMessageRecord.From)
 Exit Sub

WriteMessError:
    MsgBox "Error writing the message record to the database.  Error returned was: " & Err.Description & " in WriteMessageRecord", vbApplicationModal + vbDefaultButton1, "Write Message Error"
    Err.Clear
End Sub

Public Sub UpdateContactRecord(ContactID As String)
Dim rs As Recordset
Dim ContactEmailAddress As String
Dim ContactFullName As String
On Error GoTo BadUpdate
' Strip out <, quotes etc
ContactEmailAddress = StripEMailAddress(ContactID)
ContactFullName = Trim(StripFullName(ContactID))

Set rs = DB.OpenRecordset("Contacts", dbOpenDynaset)
'If EditContact Then

'Else

rs.FindFirst "[Contact Email] =" & "'" & ContactEmailAddress & "'"
If ContactEmailAddress = "" Then Exit Sub
If rs.NoMatch Then
    rs.AddNew
    rs("Contact Email") = ContactEmailAddress
    rs("Contact Name") = IIf(ContactFullName = "", Null, ContactFullName)
    rs.Update
    rs.Close
End If
'End If
Exit Sub
BadUpdate:
    MsgBox "Could not update your contacts database. Error was: " & Err.Description, vbCritical + vbApplicationModal, "Contact Update Error"
    Err.Clear
End Sub

Public Function StripEMailAddress(EmailID As String)
Dim iFirst As Integer
Dim iNext As Integer

iFirst = InStr(1, EmailID, "<")
iNext = InStr(iFirst + 1, EmailID, ">")
If Not iFirst = 0 Then
    StripEMailAddress = Mid(EmailID, iFirst + 1, iNext - iFirst - 1)
    Exit Function
End If
'Okay now try just the email address
iFirst = InStr(1, EmailID, "@")
If iFirst > 1 Then
    iFirst = InStrRev(EmailID, " ", iFirst)
    StripEMailAddress = Mid(EmailID, iFirst + 1, Len(EmailID) - iFirst)
    Exit Function
End If
End Function

Public Function StripFullName(EmailID As String)
Dim iFirst As Integer
Dim iNext As Integer

iFirst = InStr(1, EmailID, "<")
iNext = InStr(iFirst + 1, EmailID, ">")
If Not iFirst = 0 Then
    StripFullName = Mid(EmailID, 1, iFirst - 1)
    'Exit Function
Else
    StripFullName = EmailID
End If
'Now try for "" field
'Okay now try just the email address
iFirst = InStr(1, StripFullName, """")
iNext = InStr(iFirst + 1, StripFullName, """")
If Not iFirst = 0 Then
'[fix needed]'
    If iNext > iFirst Then
        StripFullName = Mid(StripFullName, iFirst + 1, iNext - iFirst - 1)
        Exit Function
    End If
End If

'Okay no quotes try for plain
'Okay now try just the email address
iFirst = InStr(1, EmailID, "@")
If iFirst > 1 Then
    iFirst = InStrRev(EmailID, " ")
    If Not iFirst = 0 Then
        StripFullName = Mid(EmailID, 1, iFirst - 1)
    Else
        StripFullName = ""
    End If
    
    Exit Function
End If
End Function
