Attribute VB_Name = "EU"
Option Explicit


Sub AppendInfo(Optional bMultipleDestinations As Boolean)
Dim PGPEncryptSelfToggle As Boolean
    Dim PGPWrapToggle As Boolean
    Dim PGPObscurityToggle As Boolean
    Dim tmpPass As String
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim tmpstr3 As String
    Dim tmpstr4 As String
    Dim ClipText As String
    Dim SendName As String
    Dim SendSubject As String
    Dim MixArray As String
    Dim TextLine As String
    Dim Cyphertext As String
    Dim TheMessage As String
    Dim bRes As Boolean
    Dim cmd As String
    Dim tmpchr As String
    Dim SendHeader As String
    Dim FileNum As Integer
    Dim pos1 As Integer
    Dim i As Integer
    Dim J As Integer
    Dim k As Integer
      
    
    On Error GoTo AppendError
    vb2spgpContext.Initialise
    
    
    '---------------------------------------------
    'next verify there is a recipient
    '---------------------------------------------
    If Len(PIForm(gActivePIInstance).txtTo.Text) = 0 Then
        pos1 = 0
        '---------------------------------------------
        'no recipient, so get one
        '---------------------------------------------
        frmSelectUserID.Label2 = "You did not specify a recipient in the To: box."
        frmSelectUserID.Label1 = "Please choose a recipient from your public key ring below."
        '---------------------------------------------
        'display the list of public keys
        '---------------------------------------------
        frmSelectUserID.Show vbModal
        '---------------------------------------------
        'strip off quote marks
        '---------------------------------------------
        pos1 = Len(gKeyID)
        tmpstr = ""
        For i = 1 To pos1
            tmpchr = Mid(gKeyID, i, 1)
            If (tmpchr <> Chr$(34)) Then
                tmpstr = tmpstr + tmpchr
            End If
        Next
        If gCancelAction Then Exit Sub
        PIForm(gActivePIInstance).txtTo.Text = LTrim(tmpstr)
    End If
    
    '---------------------------------------------
    'reset the gPGPKeyID, no signature on anonymous remailing
    '---------------------------------------------
    gPGPKeyID = ""
    '---------------------------------------------
    'temporarily disable encrypt to self, word wrap,
    'and gObscurity
    '---------------------------------------------
    PGPEncryptSelfToggle = False
    PGPWrapToggle = False
    PGPObscurityToggle = False
    If PIForm(gActivePIInstance).PGPSelf.Checked Then
        PGPEncryptSelfToggle = True
        PIForm(gActivePIInstance).PGPSelf.Checked = False
    End If
   
    If PIForm(gActivePIInstance).PGPObscurity.Checked Then
        PGPObscurityToggle = True
        PIForm(gActivePIInstance).PGPObscurity.Checked = False
    End If
    
       
    '---------------------------------------------
    'first BIG process... ANY REMAILER TYPE EXCEPT MIXMASTER
    '---------------------------------------------
    If gRemailerType = REMAILER_CYPHERPUNK Then
        
        If (frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) <> "chain") Then
            If Not KeyOnKeyRing(PIForm(gActivePIInstance).txtTo.Text) Then
                MsgBox "The remailer '" & PIForm(gActivePIInstance).txtTo.Text & "' " & "is not on your keyring.  Run 'Get Remailer Keys' from the Remailer List Dialog.", vbCritical + vbApplicationModal, "Missing Remailer Key"
                gCancelAction = True
                Exit Sub
            End If
            
            '---------------------------------------------
            'start the message selection at the beginning
            '---------------------------------------------
            PIForm(gActivePIInstance).MessageArea.SelStart = 0
            '---------------------------------------------
            'prepend the message with header and who to send to
            '---------------------------------------------
            SendHeader = "Request-Remailing-To: "
            PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf & SendHeader & RTrim(PIForm(gActivePIInstance).txtTo.Text) & vbCrLf & vbCrLf
            PIForm(gActivePIInstance).MessageArea.SelText = "##" & vbCrLf
            PIForm(gActivePIInstance).MessageArea.SelText = "Subject: " & PIForm(gActivePIInstance).txtSubject.Text & vbCrLf
            PIForm(gActivePIInstance).MessageArea.SelText = "CC: " & PIForm(gActivePIInstance).txtCC.Text & vbCrLf
            PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
            PIForm(gActivePIInstance).txtCC.Text = ""
            PIForm(gActivePIInstance).txtSubject.Text = ""
        Else
            'Note this should work because the remailers are global
            If bMultipleDestinations Then frmRemailerChain.Show vbModal
            If gCancelAction Then Exit Sub
            '===============
            ' Let's check to see if the remailers are on the keyring
            '===============
                For i = 1 To gnumRemailers
                    If Not KeyOnKeyRing(RTrim(Remailers(i))) Then
                        MsgBox "The remailer '" & RTrim(Remailers(i)) & "' " & "is not on your keyring.  Run 'Get Remailer Keys' from the Remailer List Dialog.", vbCritical + vbApplicationModal, "Missing Remailer Key"
                        gCancelAction = True
                        Exit Sub
                    End If
                Next
            
            If gnumRemailers <> 0 Then
                PIForm(gActivePIInstance).MessageArea.SelStart = 0
                    PIForm(gActivePIInstance).MessageArea.SelStart = 0
                    PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf & "Request-Remailing-To: " & PIForm(gActivePIInstance).txtTo.Text & vbCrLf & vbCrLf
                    PIForm(gActivePIInstance).MessageArea.SelText = "##" & vbCrLf
                    PIForm(gActivePIInstance).MessageArea.SelText = "Subject: " & PIForm(gActivePIInstance).txtSubject.Text & vbCrLf
                    PIForm(gActivePIInstance).MessageArea.SelText = "CC: " & PIForm(gActivePIInstance).txtCC.Text & vbCrLf
                    PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
                    PIForm(gActivePIInstance).txtCC.Text = "" ' Need to do this otherwise the SMTP will send CC messages
                    PIForm(gActivePIInstance).txtSubject.Text = ""
                    For J = gnumRemailers To 1 Step -1
                        gwhichRemailer = J
                        vb2spgpContext.Armor = 1
                        '[AC - 1/3/2002]
                        'vb2spgpContext.TextMode = 0
                        vb2spgpContext.TextMode = 1
                        vb2spgpContext.KeyEncrypt = 1
                        vb2spgpContext.CryptKeyID = RTrim(Remailers(J))
                        spgpEncryptMessage
                        PIForm(gActivePIInstance).MessageArea.SelStart = 0
                        If J <> 1 Then
                            PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf + "Request-Remailing-To: " & RTrim(Remailers(J)) & vbCrLf & vbCrLf
                        End If
                        PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf & "Encrypted: PGP" & vbCrLf & vbCrLf
                    Next
                    gwhichRemailer = 0
            Else
                If PGPEncryptSelfToggle Then
                    PIForm(gActivePIInstance).PGPSelf.Checked = True
                End If
                If PGPObscurityToggle Then
                    PIForm(gActivePIInstance).PGPObscurity.Checked = True
                End If
                gCancelAction = True
                Exit Sub
            End If
        End If

        If (gRemailerType = REMAILER_CYPHERPUNK) Then
            If frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) = "chain" Then
                PIForm(gActivePIInstance).txtTo.Text = RTrim$(Remailers(1))
            Else
                PIForm(gActivePIInstance).txtTo.Text = GetRemailerID
            End If
       End If
        '---------------------------------------------
        'encrypt for first remailer in the chain if PGP is on
        '---------------------------------------------
        If frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) <> "chain" Then
            vb2spgpContext.Armor = 1
            vb2spgpContext.TextMode = 1
            vb2spgpContext.KeyEncrypt = 1
            vb2spgpContext.CryptKeyID = PIForm(gActivePIInstance).txtTo.Text
            spgpEncryptMessage
            PIForm(gActivePIInstance).MessageArea.SelStart = 0
            PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf + "Encrypted: PGP" & vbCrLf & vbCrLf
        End If
        '---------------------------------------------
        'end of processing for single or chaining remailer
        '---------------------------------------------
    Else
        If iFileExists(gMixPath + "\mixmaste.exe") = False Then
            MsgBox "You need a copy of Mixmaster installed.  Click on the Menu Remailer(Options)to set the path.  Mixmaster is export controlled cryptography.  See http://www.obscura.com/~loki for additional information.", 48, gPiStr
            '---------------------------------------------
            'turn options back on since we are exiting
            '---------------------------------------------
            If PGPEncryptSelfToggle Then
                PIForm(gActivePIInstance).PGPSelf.Checked = True
            End If
            If PGPObscurityToggle Then
                PIForm(gActivePIInstance).PGPObscurity.Checked = True
            End If
            gCancelAction = True
            Exit Sub
        End If
        
        '---------------------------------------------
        'second test, if the mix message is too big,
        'suggest a type 1 and bail
        '---------------------------------------------
        If Len(PIForm(gActivePIInstance).MessageArea.Text) >= 1000000 Then
            MsgBox "Sorry, Private i Mail currently only supports Mixmaster messages no larger than 1M bytes.  Either split the files into smaller chunks using the file splitter or use a Type I remailer instead.", 48, "Mixmaster"
            '---------------------------------------------
            'turn options back on since we are exiting
            '---------------------------------------------
            If PGPEncryptSelfToggle Then
                PIForm(gActivePIInstance).PGPSelf.Checked = True
            End If
            'If PGPWrapToggle Then
             '   PIForm(gActivePIInstance).PGPWrap.Checked = True
            'End If
            If PGPObscurityToggle Then
                PIForm(gActivePIInstance).PGPObscurity.Checked = True
            End If
            gCancelAction = True
            Exit Sub
        End If
        
        '---------------------------------------------
        'got this far, start mixmaster processing
        '---------------------------------------------
        'first, dump the message to a temporary file
        '---------------------------------------------
        'ClipText = PIForm(gActivePIInstance).MessageArea.Text
        'FileNum = FreeFile
        bRes = PutFileText(gMixPath & "\tmx.out", PIForm(gActivePIInstance).MessageArea.Text)
        SendName = PIForm(gActivePIInstance).txtTo.Text
        SendSubject = PIForm(gActivePIInstance).txtSubject
        If SendSubject = "" Then
            SendSubject = "Message"
        End If
        If Not frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) = "chain" Then
            i = InStr(1, frmRemailerList.List3.List(frmRemailerList.List3.ListIndex), " ")
            tmpstr2 = Mid$(frmRemailerList.List3.List(frmRemailerList.List3.ListIndex), 1, i - 1)
            MixArray = GetMixNumber(tmpstr2)
            If MixArray <> "0" Then
                '---------------------------------------------
                'append the remailer to the to line
                '---------------------------------------------
                PIForm(gActivePIInstance).txtTo.Text = GetMixAddress(tmpstr2)
                '---------------------------------------------
                'clear the subject line
                '---------------------------------------------
                PIForm(gActivePIInstance).txtSubject = ""
            Else
                MsgBox "Couldn't find the selected remailer in the Type 2 list.", 48, "Mixmaster"
            End If
        End If
        '---------------------------------------------
        'if a chain, display the chain info and process
        '---------------------------------------------
        If frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) = "chain" Then
            tmpstr3 = ""
            frmRemailerChain.Show vbModal
            If gnumRemailers <> 0 Then
                For i = 1 To gnumRemailers
                   'tmpstr4 = Mid(Remailers(i), 1, InStr(1, Remailers(i), " ") - 1)
                    tmpstr2 = GetMixNumber(Remailers(i))
                    MixArray = MixArray & tmpstr2 & " "
                    If tmpstr2 = "0" And RTrim$(Remailers(i)) <> "random" Then
                        MsgBox "Couldn't find the selected remailer in the Type 2 list.", vbApplicationModal + vbCritical, "Mixmaster"
                    Else
                        If tmpstr3 = "" And RTrim$(Remailers(i)) <> "random" And tmpstr2 <> "0" Then
                            tmpstr3 = RTrim$(Remailers(i))
                        End If
                    End If
                Next
                MixArray = RTrim(MixArray)
                '---------------------------------------------
                'put the first remailer in the chain on the to: line
                '---------------------------------------------
                PIForm(gActivePIInstance).txtTo.Text = Remailers(1)
                '---------------------------------------------
                'clear the subject
                '---------------------------------------------
                PIForm(gActivePIInstance).txtSubject = ""
            Else
                '---------------------------------------------
                'turn options back on since we are exiting
                '---------------------------------------------
                If PGPEncryptSelfToggle Then
                    PIForm(gActivePIInstance).PGPSelf.Checked = True
                End If
                If PGPObscurityToggle Then
                    PIForm(gActivePIInstance).PGPObscurity.Checked = True
                End If
                gCancelAction = True
                'Beep
                Exit Sub
            End If
        End If
        '---------------------------------------------
        'build the commandline prompt
        '---------------------------------------------
        If gNewsgroupType = USENET Then
            cmd = gMixPath & "\mixmaste.exe"
            cmd = cmd & " tmx.out -s " + Chr(34) & SendSubject & Chr(34)
            cmd = cmd & " -v " & Chr(34) & MailHeader(1).ID & MailHeader(1).Value & Chr(34) ' "Newsgroups: " & txtXHeader & Chr(34)
            cmd = cmd & " -to " & Chr(34) & SendName & Chr(34)
            cmd = cmd & " -l " + MixArray
            'reset...
            gNewsgroupType = 0
        Else
           ' cmd = "Set mixpath = " & gMixPath
            'ExecCmd (cmd)
            'Shell (cmd)
            cmd = gMixPath & "\mixmaste.exe"
            cmd = cmd & " tmx.out -s " + Chr(34) & SendSubject & Chr(34)
            'cmd = cmd & " -o " & gMixPath & "\mix.out -to " + Chr(34) & SendName & Chr(34)
            cmd = cmd & " -to " & Chr(34) & SendName & Chr(34)
            cmd = cmd & " -l " & MixArray
        End If
       CheckLen (cmd)
       If gCancelAction Then
            MsgBox "A command line > 128 could have been the result of too many news group or a subject that was too long.", vbApplicationModal + vbCritical
            Beep
        Else
            DoEvents
            ExecCmd (cmd)
        End If
        '[debug 1]
        DoEvents
        '---------------------------------------------
        'wipe the plaintext
        '---------------------------------------------
        WipeFile gMixPath + "\tmx.out"
        '---------------------------------------------
        'move the ciphertext to the message box
        '---------------------------------------------
        
        If Not iFileExists(gMixPath & "\outfile.*") Then
            MsgBox "Mixmaster did not produce any valid files in the the directory " & gMixPath & ".  Couldn't complete Mix operation.", 48, gPiStr
            gCancelAction = True
        Else
            J = 0
            For i = 0 To 50
                If iFileExists(gMixPath & "\outfile." & i) Then J = J + 1
            Next
            PIForm(gActivePIInstance).ShowStatus 1, "Multipart message being generated..."
            DoEvents
            For i = 0 To J - 1
              TheMessage = ""
              If iFileExists(gMixPath & "\outfile." & i) Then
                FileNum = FreeFile
                Open gMixPath & "\outfile." & i For Input As FileNum
                Do While Not EOF(FileNum)
                    Line Input #FileNum, TextLine
                    If InStr(1, TextLine, "To:", vbTextCompare) > 0 Then
                        k = InStr(1, TextLine, "To: ", vbTextCompare)
                        PIForm(gActivePIInstance).txtTo.Text = Mid(TextLine, k + 4, Len(TextLine) - k - 3)
                        Exit Do
                    End If
                Loop
                DoEvents
                Do While Not EOF(FileNum)
                    Line Input #FileNum, TextLine
                    If InStr(1, TextLine, "::", vbTextCompare) > 0 Then
                        TheMessage = TheMessage & TextLine & vbCrLf
                        Do While Not EOF(FileNum)
                            Line Input #FileNum, TextLine
                            TheMessage = TheMessage & TextLine & vbCrLf
                        Loop
                    End If
                Loop
                PIForm(gActivePIInstance).MessageArea.SelLength = Len(PIForm(gActivePIInstance).MessageArea)
                PIForm(gActivePIInstance).MessageArea.SelStart = 0
                PIForm(gActivePIInstance).MessageArea.Text = ""
                PIForm(gActivePIInstance).MessageArea.Text = TheMessage
                
                Close FileNum
                Kill gMixPath & "\outfile." & i
                DoEvents
                PIForm(gActivePIInstance).ShowStatus 1, "Sending part " & i + 1 & " of " & J & " mixmaster messages."
                SendToOutBox
                'SendMailMessage
                DoEvents
                gMixSent = True
            End If
        Next
        End If
        'This will stop send message sending the final one
        GoTo temp:
            FileNum = FreeFile
            Open gMixPath + "\mix.out" For Input As FileNum
            '---------------------------------------------
            'get ride of the first two blank lines
            '---------------------------------------------
            Line Input #FileNum, TextLine
            Line Input #FileNum, TextLine
            While Not EOF(FileNum)
                Line Input #FileNum, TextLine
                Cyphertext = Cyphertext + TextLine & vbCrLf
            Wend
            Close #FileNum
            PIForm(gActivePIInstance).MessageArea.Text = Cyphertext
            '---------------------------------------------
            'del the ciphertext file
            '---------------------------------------------
            Kill gMixPath + "\mix.out"
        'Else
            MsgBox "Mixmaster did not produce any valid files in the the directory " & gMixPath & ".  Couldn't complete Mix operation.", 48, gPiStr
            gCancelAction = True
        'End If
temp:
    '---------------------------------------------
    'done processing *****
    '---------------------------------------------
    
    '---------------------------------------------
    'turn options back on since we are exiting
    '---------------------------------------------
    If PGPEncryptSelfToggle Then
        PIForm(gActivePIInstance).PGPSelf.Checked = True
    End If
    'If PGPWrapToggle Then
    '    PIForm(gActivePIInstance).PGPWrap.Checked = True
   ' End If
    If PGPObscurityToggle Then
        PIForm(gActivePIInstance).PGPObscurity.Checked = True
    End If
End If
    Exit Sub
AppendError:
    MsgBox Err.Description & " in AppendInfo", vbCritical + vbApplicationModal, App.Title
    Err.Clear
    gCancelAction = True
    
End Sub

Sub CheckMultipleKey()
    Dim tmpstr As String
    'check if the user wants to use multiple keys for signing
    Dim SectionName As String
    SectionName = "PGP Options"
    tmpstr = ReadProfile(SectionName, "Multiple")
    If tmpstr = "true" Then
        Form26.Show 1
    End If
End Sub



Sub EncryptFile(msgKey As String, pgpParam As String, theFile As String)

    Dim FileNum
    Dim ErrorNum As Integer
    Dim i As Integer
    Dim J As Integer
    Dim cmd As String
    Dim SectionName As String
    Dim SignKey As String
    Dim notfound As Boolean
    Dim ErrOK As String
    Dim tempstr As String
    Dim tmpstr2 As String
    Dim TextLine As String
    Dim ErrorString As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim Recipient As String
    Dim PGPSelfKey As String
    Dim PGPString As String
    
    On Error GoTo EncryptFileError
    
    '---------------------------------------------
    'initialize
    '---------------------------------------------
    notfound = True
    ErrOK = False
    tempstr = ""
    pos1 = 0
    pos2 = 0
    gKeyID = ""
    SignKey = ""
    Recipient = ""
    PGPSelfKey = ""
    PGPString = ""
    ErrorNum = 1
    'ClipText = PIForm(gActivePIInstance).MessageArea.Text
    
    '---------------------------------------------
    'handle the signing option, if the user chose
    'encrypt and sign from the menu, the key should
    'be stored in gPGPKeyID.  This is not the best
    'programming practice, too many global vars in PI @@.
    '---------------------------------------------
    
    If gPGPKeyID <> "" Then
        '---------------------------------------------
        'strip out to keep only the bare e-mail address
        '---------------------------------------------
        pos1 = InStr(gPGPKeyID, "<") + 1
        pos2 = InStr(gPGPKeyID, ">")
        vb2spgpContext.CryptKeyID = gPGPKeyID
        vb2spgpContext.Armor = 1
        vb2spgpContext.TextMode = 0
        vb2spgpContext.KeyEncrypt = 1
        If pos1 = 1 And pos2 > 1 Then
            SignKey = " -u " & gPGPKeyID
            '[??]
            gPGPKeyID = ""
        Else
            SignKey = " -u " & Mid(gPGPKeyID, pos1, pos2 - pos1)
        End If
    End If
    
    '---------------------------------------------
    'handle the encrypt to self option
    '---------------------------------------------
    If PIForm(gActivePIInstance).PGPSelf.Checked Then
        '---------------------------------------------
        'load the preferred key for encrypting to self
        '---------------------------------------------
        If gPGPVersion = PGP5x Then
            MsgBox "Encrypted to self...Not yet implemented...."
            Exit Sub
        End If
        
        SectionName = "PGP Options"
        PGPSelfKey = ReadProfile(SectionName, "KeyID")
        tempstr = gPGPKeyID
        If PGPSelfKey = "" Then
            '---------------------------------------------
            'no preferred key, so select from list
            '---------------------------------------------
            Form26.Label1.Caption = "Select a key for encrypting to self."
            CheckMultipleKey
            PGPSelfKey = gPGPKeyID
            tmpstr2 = PGPSelfKey
            i = InStr(1, tmpstr2, " <")
            If i > 0 Then
                '---------------------------------------------
                'strip out to keep only the bare e-mail address
                '---------------------------------------------
                PGPSelfKey = Mid$(Mid$(tmpstr2, i + 1, Len(tmpstr2) - i), 1, Len(tmpstr2) - 1)
            End If
        Else
            '---------------------------------------------
            'there is a preferred key, fetch it from public key ring
            '---------------------------------------------
            tempstr = PGPSelfKey
            FileNum = FreeFile
            Open App.Path + "\pubkeys.out" For Input As FileNum
            Do While Not EOF(FileNum)
                Line Input #FileNum, TextLine
                If InStr(LCase(TextLine), LCase(PGPSelfKey)) Then
                    '---------------------------------------------
                    'strip out to keep only the bare e-mail address
                    '---------------------------------------------
                    pos1 = InStr(TextLine, "<") + 1
                    pos2 = InStr(TextLine, ">")
                    PGPSelfKey = Mid(TextLine, pos1, pos2 - pos1)
                    Exit Do
                End If
            Loop
            Close #FileNum
        End If
        gPGPKeyID = tempstr
        '---------------------------------------------
        'could not find the key in pubkeys.out, so advise and bail
        '---------------------------------------------
        If gPGPKeyID = "" Then
            MsgBox tmpstr2 + " does not appear in PUBKEYS.OUT.  Either there is no such key in your public key ring, or PUBKEYS.OUT needs updating (use the option on the KEYS menu)."
            Exit Sub
        Else
            '---------------------------------------------
            'key found so embed in quotes
            '---------------------------------------------
            PGPSelfKey = Chr$(34) + gPGPKeyID + Chr$(34)
        End If
    End If 'End self option
    
    '---------------------------------------------
    'handle the recipient in the "to:" box
    '---------------------------------------------
    If InStr(1, msgKey, "@") <> 0 Then
                   
            If Not spgpKeyIsOnRing(msgKey) Then
                MsgBox "The receipient is not on your keyring.", vbCritical
                Exit Sub
            End If
            vb2spgpContext.Armor = 1
            vb2spgpContext.TextMode = 0
            vb2spgpContext.KeyEncrypt = 1
            vb2spgpContext.CryptKeyID = msgKey
    Else
        '---------------------------------------------
        'handle blank recipient in the to: box
        '---------------------------------------------
        'bypass if this is a clear signed msg
        '---------------------------------------------
        
         If vb2spgpContext.Clear = 0 And vb2spgpContext.ConventionalEncrypt = 0 Then
                PIForm(gActivePIInstance).GetRecipient
                Recipient = PIForm(gActivePIInstance).txtTo.Text
        End If
    End If
    
    '---------------------------------------------
    'handle other UserIDs (for encrypting with multiple keys)
    '---------------------------------------------
    
    'If gPGPVersion = PGP26x Then
      '  Form11.Text1 = gKeyID
      '  Form11.Show 1
   ' Else
     '   MsgBox "PGP5/6 - Use of multiple keys has not yet been implemented.", vbCritical
     '   Exit Sub
    'End If
    '---------------------------------------------
    'start building up pgp command line, tack on encrypt-to-self
    '---------------------------------------------
    If Len(PGPSelfKey) > 0 Then
        gKeyID = gKeyID & " " & PGPSelfKey
    End If
        
    '---------------------------------------------
    'pgp command line, tack on signing, if there
    '---------------------------------------------
    If Len(SignKey) > 0 Then
        gKeyID = gKeyID & SignKey
    End If
    '
    '---------------------------------------------------------------
    'All options and flags must have been set by the time we get here
    '--------------------------------------------------------------
    '
    If Not gCancelAction Then
      
        vb2spgpContext.FileIn = theFile
        spgpEncryptfile
   End If
   'Need to reset this
   gCancelAction = False
   
    Exit Sub
    
EncryptFileError:
    gCancelAction = False
    MsgBox Err.Description & " (EncryptFile) -> " & ErrorString, vbApplicationModal, gPiStr
    Err.Clear
End Sub
Sub EncryptMessage(EncryptKeyID As String, pgpParam As String)
         
   Dim FileNum As Integer
    Dim i As Long
   ' Dim cmd As String
    Dim ClipText As String
    Dim Cyphertext As String
    Dim SectionName As String
    Dim SignKey As String
    Dim ErrorString As String
    Dim bRes As Long
    Dim tmpstr1 As String
    Dim tmpstr2 As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim Recipient As String
    Dim PGPSelfKey As String
    Dim PGPString As String
    Dim TextLine As String
    
        On Error GoTo EncryptError
        '---------------------------------------------
        'first verify there is something to encrypt
        '---------------------------------------------
       ' If Len(PIForm(gActivePIInstance).MessageArea.Text) = 0 Then
           ' PIForm(gActivePIInstance).MessageArea.Text = "Nothing to Encrypt!"
            'PIForm(gActivePIInstance).MessageArea.SelStart = 0
            'PIForm(gActivePIInstance).MessageArea.SelLength = Len(PIForm(gActivePIInstance).MessageArea)
            'Beep
            'Exit Sub
        'End If
        '---------------------------------------------
        'there is text
        'set the error handler to section 1
        '---------------------------------------------
         ClipText = PIForm(gActivePIInstance).MessageArea.Text
        'If PIForm(gActivePIInstance).PGPWrap.Checked Then
         '   PIForm(gActivePIInstance).MessageArea.Text = InsertCRLFs()
        'End If
        '---------------------------------------------
        'handle the signing option, if the user
        'chose encrypt and sign from the menu, the
        'key should be stored in gPGPKeyID
        'This is not the best programming practice,
        'so it should be revised.  Way too many global
        'variables in PI.  @@
        ''''do this in EnSign
        
        '
        '---------------------------------------------
        'handle the recipient in the "to:" box
        'What we are assuming here is that the user has typed in the receipient in the to: box
        '---------------------------------------------
        If EncryptKeyID <> "" Then
        
                 
                If Not KeyOnKeyRing(EncryptKeyID) Then
                    MsgBox "The receipient is not on your keyring.", vbCritical
                    Exit Sub
                End If
                vb2spgpContext.KeyEncrypt = 1
                vb2spgpContext.CryptKeyID = EncryptKeyID
        Else
            '---------------------------------------------
            'handle blank recipient in the to: box
            '---------------------------------------------
            
                If Not vb2spgpContext.Clear = 1 And Not vb2spgpContext.ConventionalEncrypt = 1 Then
                    PIForm(gActivePIInstance).GetRecipient
                    Recipient = PIForm(gActivePIInstance).txtTo.Text
                    vb2spgpContext.CryptKeyID = PIForm(gActivePIInstance).txtTo.Text
                End If
        End If
        
        '---------------------------------------------
        'handle other UserIDs (for encrypting with multiple keys)
        '---------------------------------------------
        If PIForm(gActivePIInstance).PGPMultiple.Checked Then
                MsgBox "PGP5/6 - Use of multiple keys has not yet been implemented.", vbCritical
                Exit Sub
        End If
        
        '---------------------------------------------
        'start building up pgp command line, tack on encrypt-to-self
        '---------------------------------------------
       ' If Len(PGPSelfKey) > 0 Then
         '   gKeyID = gKeyID + " " + PGPSelfKey
       ' End If
        
        '---------------------------------------------
        'pgp command line, tack on signing, if there
        '---------------------------------------------
       ' If Len(SignKey) > 0 Then
            'gKeyID = gKeyID + SignKey
        'End If
                
        If Not gCancelAction Then
            'If Len(PIForm(gActivePIInstance).MessageArea.Text) > 50000 Then
                   ' vb2spgpContext.FileIn = GetTemporaryFile
                    'bRes = PutFileText(vb2spgpContext.FileIn, PIForm(gActivePIInstance).MessageArea.Text)
                   ' vb2spgpContext.FileOut = GetTemporaryFile
                    'If Not spgpEncryptfile Then
                      '  KillTemporaryFiles
                       ' Err.Raise 10004, , "Error encrypting message."
                    'Else
                     '   PIForm(gActivePIInstance).MessageArea.Text = GetFileText(vb2spgpContext.FileOut)
                    'End If
                    'KillTemporaryFiles
           ' Else
                spgpEncryptMessage
            'End If
        End If
    

gCancelAction = False
Exit Sub
    
EncryptError:
        gCancelAction = False
        MsgBox Err.Description & "Error encrypting message.  Error returned was: " & Err.Description, vbApplicationModal + vbCritical
        Err.Clear

End Sub

Function GetKeyID(UserID As String) As String
        
    Dim FileNum As Integer
    Dim KeyValue As String
    Dim TextLine As String
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim i As Integer
    Dim J As Integer
    
    FileNum = FreeFile
    On Error GoTo IDError
    GetKeyID = ""
    Open App.Path + "\pubkeys.out" For Input As FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, TextLine
        tmpstr = TextLine
        'check if this line contains pub
        tmpstr2 = Mid$(tmpstr, 1, 3)
        If tmpstr2 = "pub" Then
            'if it does grab the key
            'find the first /
            J = InStr(1, tmpstr, "/")
            'now find the the first space after the /
            i = InStr(J + 1, tmpstr, " ")
            'everything between is the key
            KeyValue = Mid$(tmpstr, J + 1, i - J)
        End If
        'if not, we have a multi-user ID, keep the old key value
        'parse the string for the user ID
        If InStr(1, tmpstr, UserID) > 0 Then
            'if it's there, return the key value and exit the sub
            GetKeyID = KeyValue
            Close #FileNum
            Exit Function
        End If
        'if not, check the next line, repeat
    Wend
    Close #FileNum
    Exit Function
    
IDError:
    MsgBox Err.Description & " in (GetKeyID).", vbApplicationModal, gPiStr
    Err.Clear
End Function

Function GetMixAddress(mixName As String) As String
    Dim FileNum As Integer
    Dim Counter As Integer
    Dim TextLine As String
    Dim tmpstr As String
    Dim i As Integer
    
    If iFileExists(gMixPath & "\type2.lis") Then
        FileNum = FreeFile
        Open gMixPath & "\type2.lis" For Input As FileNum
        While Not EOF(FileNum)
            Line Input #FileNum, TextLine
            'grab the remailer name
            tmpstr = Mid$(TextLine, 1, Len(mixName))
            'see if the two match
            If tmpstr = mixName Then
                'grab the address
                i = InStr(Len(mixName) + 2, TextLine, " ")
                GetMixAddress = Mid$(TextLine, Len(mixName) + 2, i - (Len(mixName) + 2))
                Close #FileNum
                Exit Function
            End If
        Wend
        Close #FileNum
    Else
        MsgBox "TYPE2.LIS not present in the same directory as " + gMixPath + ".  Can't complete Mix operation.", 48, gPiStr
    End If
End Function

Function GetMixNumber(theMixName As String) As String
    
    Dim FileNum
    Dim Counter As Integer
    Dim TextLine As String
    
    If iFileExists(gMixPath & "\type2.lis") Then
        FileNum = FreeFile
        Counter = 1
        GetMixNumber = "0"
        Open gMixPath & "\type2.lis" For Input As FileNum
        While Not EOF(FileNum)
            Line Input #FileNum, TextLine
           If Not InStr(1, TextLine, theMixName) = 0 Then
           'If Mid$(TextLine, 1, Len(theMixName)) = theMixName Then
                GetMixNumber = Format(Counter)
                Close #FileNum
                Exit Function
            End If
            If Not Mid$(TextLine, 1, Len(theMixName)) = "" Then Counter = Counter + 1
        Wend
        Close #FileNum
    Else
        MsgBox "TYPE2.LIS not present in the same application directory as " & gMixPath + ".  Can't complete Mix operation.", vbApplicationModal, gPiStr
        Err.Clear
    End If
End Function

Sub GetURL(URLString As String)
Dim CurrentWindow As Long
Dim Length As Integer
Dim ListItem As String
Dim thwnd As Long

    On Error GoTo URLError
    'added for NT support
    CurrentWindow = GetWindow(PIForm(gActivePIInstance).hWnd, GW_HWNDFIRST)
    ' Loop while the hWnd returned by GetWindow is valid.
    While CurrentWindow <> 0
        'Get the length of the task name identified by
        'CurrentWindow in the list.
        Length = GetWindowTextLength(CurrentWindow)
        'Get the task name of the task in the master list.
        ListItem = Space(Length + 1)
        Length = GetWindowText(CurrentWindow, ListItem, Length + 1)
        'If there is an actual task name in the list, add the
        'item to the list.
        If Length > 0 Then
            If InStr(1, ListItem$, gBrowserString) Then
                thwnd = SetFocusAPI(CurrentWindow)
                'got rid of AppActivate because of string limitation
                'since the browser isn't launched, this should work
                'replaced with SetFocus
                'AppActivate ListItem$
                'added for NT support
                thwnd = FindWindow(0, Mid$(ListItem$, 1, Len(ListItem$) - 1))
                SetWindowPos thwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                SendKeys "% thwnd", True
                SendKeys gURLStart, True
                Clipboard.SetText URLString
                SendKeys "^v", True
                SendKeys gURLEnd, True
                Exit Sub
            End If
        End If
        'Get the next task list item in the master list.
        CurrentWindow = GetWindow(CurrentWindow, GW_HWNDNEXT)
        'Process Windows events.
        DoEvents
    Wend
    MsgBox "Your Web browser doesn't appear to be running.  Click OK to run the browser, then select the menu item again after the browser has loaded.  If that does not work, please check that your browser string is correct under Web Options.", 48
    thwnd = Shell(gBrowserPath, 2)
    Exit Sub
URLError:
    MsgBox Err.Description & " Couldn't activate or load the Web browser.  Make sure the settings in the Web Options dialog are correct.", 16
    Err.Clear
End Sub
Sub InsertKey(selector As Integer)

Dim Storestring1 As String
Dim Storestring2 As String
Dim cmd As String
Dim iResult As Integer
Dim Ecount As Integer
Dim TextLine As String
Dim FileNum As Integer
Dim Cyphertext As String

On Error GoTo BadInsertKey

    If selector = IDKEY Then
        vb2spgpContext.SelectPrivateKeys = False
        frmViewKeyRing.lblContext = "Select which key you wish to insert into the message."
        frmViewKeyRing.Show vbModal
        PIForm(gActivePIInstance).MessageArea.SelText = GetKey(Key.KeyID, vb2spgpContext.SelectPrivateKeys)
    Else
        PIForm(gActivePIInstance).MessageArea.SelText = GetKey(Nym.ID, vb2spgpContext.SelectPrivateKeys)
    End If
    Exit Sub

    
   
BadInsertKey:
    MsgBox Err.Description & " in (Insert Key).", vbApplicationModal, App.Title
    Err.Clear
End Sub


Function ReadProfile(pSectionName As String, pKeyName As String) As String
       
    Dim KeyDefault As String
    Dim KeyValue As String
    Dim ReturnCode As Integer
    
    KeyDefault = ""
    ReadProfile = GetSetting(AppName:=App.Title, section:=pSectionName, _
                       Key:=pKeyName, Default:=KeyDefault)
End Function
Sub RestoreSettings()
      
    Dim SectionName As String
    Dim ErrorNum As Integer
    Dim tmpstr As String
    Dim i As Integer
    
    On Error GoTo RestoreError
    
    ErrorNum = 0
    
    SectionName = "PGP Options"
    tmpstr = ReadProfile(SectionName, "PGPSelf")
    If tmpstr = "true" Then
        PIForm(gActivePIInstance).PGPSelf.Checked = True
    End If
    tmpstr = ReadProfile(SectionName, "PGPMultiple")
    If tmpstr = "true" Then
        PIForm(gActivePIInstance).PGPMultiple.Checked = True
    End If
    tmpstr = ReadProfile(SectionName, "PGPObscurity")
    If tmpstr = "true" Then
        PIForm(gActivePIInstance).PGPObscurity.Checked = True
        gObscurity = 1
    End If
    tmpstr = ReadProfile(SectionName, "PGPMin")
    If tmpstr = "true" Then
        PIForm(gActivePIInstance).PGPMin.Checked = True
        gMinState = 2
    End If
    tmpstr = ReadProfile(SectionName, "PGPClose")
        tmpstr = ReadProfile(SectionName, "PGPEyes")
    '
    ' Remailer info
    '
    SectionName = "Remailer Info"
    tmpstr = ReadProfile(SectionName, "RemailersCP")
    If tmpstr = "true" Then
        gRemailerType = REMAILER_CYPHERPUNK
        gSortRemailer.cpunk = 1
    End If
    tmpstr = ReadProfile(SectionName, "EncryptionToRemailers")
    If tmpstr = "true" Then
       ' gEncryptToRemailer = True
        'PIForm(gActivePIInstance).optCPSendOptions(0).Value = True
    Else
        'PIForm(gActivePIInstance).optCPSendOptions(1).Value = True
    End If
    
    tmpstr = ReadProfile(SectionName, "RemailersNone")
    If tmpstr = "true" Then
        gRemailerType = STANDARD_EMAIL
        gSortRemailer.cpunk = 1
        'gEncryptToRemailer = False
        'PIForm(gActivePIInstance).optRemailers(0).Value = True
        'PIForm(gActivePIInstance).OptionsAdv.Enabled = True
    End If
    tmpstr = ReadProfile(SectionName, "RemailersMix")
    If tmpstr = "true" Then
        gRemailerType = REMAILER_MIX
        'gEncryptToRemailer = True
        'PIForm(gActivePIInstance).optRemailers(2).Value = True
        gSortRemailer.mix = 1
        gMixPath = ReadProfile(SectionName, "MixmasterPath")
        If gMixPath = "" Then gMixPath = App.Path
    End If
    
    
    gCutStr = ReadProfile(SectionName, "CutMarks")
    gLatentStr = ReadProfile(SectionName, "Latent")
    '
    ' News Posters
    '
    SectionName = "News Posters"
    tmpstr = ReadProfile(SectionName, "USENETNone")
    'If tmpstr = "true" Then
       ' PIForm(gActivePIInstance).USENETNone.Checked = True
       ' gNewsgroupType = 0
    'End If
    tmpstr = ReadProfile(SectionName, "USENETFI")
    If tmpstr = "true" Then
        'PIForm(gActivePIInstance).USENETFi.Checked = True
        gNewsgroupType = penet
        'PIForm(gActivePIInstance).USENETGate.Visible = False
        'PIForm(gActivePIInstance).USENETFi.Visible = True
        'PIForm(gActivePIInstance).UseNetSoda.Visible = False
    End If
    tmpstr = ReadProfile(SectionName, "USENETGate")
    If tmpstr = "true" Then
        'PIForm(gActivePIInstance).USENETGate.Checked = True
        gNewsgroupType = USENET
        'PIForm(gActivePIInstance).USENETGate.Visible = True
        'PIForm(gActivePIInstance).USENETFi.Visible = False
        'PIForm(gActivePIInstance).UseNetSoda.Visible = False
    End If
    tmpstr = ReadProfile(SectionName, "USENETSoda")
    If tmpstr = "true" Then
       ' PIForm(gActivePIInstance).UseNetSoda.Checked = True
        gNewsgroupType = Soda
       ' PIForm(gActivePIInstance).USENETGate.Visible = False
       ' PIForm(gActivePIInstance).USENETFi.Visible = False
       ' PIForm(gActivePIInstance).UseNetSoda.Visible = True
    End If
    
    '---------------------------------------------
    ' recall the font settings for frmPI
    '---------------------------------------------
    SectionName = "Settings"
        
    ErrorNum = 1
    PIForm(gActivePIInstance).txtSubject.FontName = ReadProfile(SectionName, "selfontname")
    ErrorNum = 2
    PIForm(gActivePIInstance).txtSubject.FontBold = ReadProfile(SectionName, "FontBold")
    ErrorNum = 3
    PIForm(gActivePIInstance).txtSubject.FontItalic = ReadProfile(SectionName, "FontItalic")
    ErrorNum = 4
    PIForm(gActivePIInstance).txtSubject.FontStrikethru = ReadProfile(SectionName, "FontStrikeThru")
    ErrorNum = 5
    PIForm(gActivePIInstance).txtSubject.FontSize = ReadProfile(SectionName, "FontSize")
    
    ErrorNum = 1
    'PIForm(gActivePIInstance).txtCC.FontName = ReadProfile(SectionName, "selfontname")
    'ErrorNum = 2
    'PIForm(gActivePIInstance).txtCC.FontBold = ReadProfile(SectionName, "FontBold")
   ' ErrorNum = 3
   ' PIForm(gActivePIInstance).txtCC.FontItalic = ReadProfile(SectionName, "FontItalic")
   ' ErrorNum = 4
   ' PIForm(gActivePIInstance).txtCC.FontStrikethru = ReadProfile(SectionName, "FontStrikeThru")
   ' ErrorNum = 5
   ' PIForm(gActivePIInstance).txtCC.FontSize = ReadProfile(SectionName, "FontSize")
    
    ErrorNum = 1
    'PIForm(gActivePIInstance).Text4.FontName = ReadProfile(SectionName, "selfontname")
    'ErrorNum = 2
    'PIForm(gActivePIInstance).Text4.FontBold = ReadProfile(SectionName, "FontBold")
    'ErrorNum = 3
    'PIForm(gActivePIInstance).Text4.FontItalic = ReadProfile(SectionName, "FontItalic")
   ' ErrorNum = 4
    'PIForm(gActivePIInstance).Text4.FontStrikethru = ReadProfile(SectionName, "FontStrikeThru")
    'ErrorNum = 5
   ' PIForm(gActivePIInstance).Text4.FontSize = ReadProfile(SectionName, "FontSize")
    
    ErrorNum = 1
    PIForm(gActivePIInstance).MessageArea.SelFontName = ReadProfile(SectionName, "selfontname")
    ErrorNum = 2
    PIForm(gActivePIInstance).MessageArea.SelBold = ReadProfile(SectionName, "FontBold")
    ErrorNum = 3
    PIForm(gActivePIInstance).MessageArea.SelItalic = ReadProfile(SectionName, "FontItalic")
    ErrorNum = 4
    PIForm(gActivePIInstance).MessageArea.SelStrikeThru = ReadProfile(SectionName, "FontStrikeThru")
    ErrorNum = 5
    PIForm(gActivePIInstance).MessageArea.SelFontSize = ReadProfile(SectionName, "FontSize")
    
    ErrorNum = 1
    'PIForm(gActivePIInstance).txtTo.FontName = ReadProfile(SectionName, "selfontname")
    'ErrorNum = 2
    'PIForm(gActivePIInstance).txtTo.FontBold = ReadProfile(SectionName, "FontBold")
    'ErrorNum = 3
    'PIForm(gActivePIInstance).txtTo.FontItalic = ReadProfile(SectionName, "FontItalic")
    'ErrorNum = 4
    'PIForm(gActivePIInstance).txtTo.FontStrikethru = ReadProfile(SectionName, "FontStrikeThru")
    'ErrorNum = 5
    'PIForm(gActivePIInstance).txtTo.FontSize = ReadProfile(SectionName, "FontSize")
    
    '---------------------------------------------
    ' recall the preferred path for saving files
    '---------------------------------------------
    'SectionName = "Options"
    'App.Path = ReadProfile(SectionName, "Preferred Path")
    'If Len(App.Path) = 0 Then
        'WriteProfile "Options", "Preferred Path", App.Path
       ' App.Path = Str(App.Path)
    'End If
   ' frmFileOptions.Text7.Text = App.Path
    Exit Sub

RestoreError:
    If Err.Number = 380 Or Err = 13 Then
        '---------------------------------------------
        'handler for errors reading a value out of the PIDAHO.INI file
        '---------------------------------------------
       ' If iFileExists(App.Path + "\HELLO.TXT") Then
            '---------------------------------------------
            'first time thru on v2.8c save new setting for mit key server
            '---------------------------------------------
          '  WriteProfile "Net Info", "GetKeyURL", "http://swissnet.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
       ' End If
        If ErrorNum = 1 Then
            WriteProfile "Settings", "selfontname", "MS Sans Serif"
        ElseIf ErrorNum = 2 Then
            WriteProfile "Settings", "FontBold", False
        ElseIf ErrorNum = 3 Then
            WriteProfile "Settings", "FontItalic", False
        ElseIf ErrorNum = 4 Then
            WriteProfile "Settings", "FontStrikeThru", False
        ElseIf ErrorNum = 5 Then
            WriteProfile "Settings", "FontSize", 8.25
        End If
        ErrorNum = 0
        Resume Next
   End If
End Sub

Sub SaveSettings()
    
    Dim iResult As Integer
    Dim Ecount As Integer
    Dim cmd As String
    
    Dim KeyValue As String
    Dim SectionName As String
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim tmpstr3 As String
    Dim tmpstr4 As String
    Dim tmpstr5 As String
    Dim tmpstr6 As String
    
    
    '---------------------------------------------
    'general settings
    '---------------------------------------------
    SectionName = "Settings"
    '---------------------------------------------
    'test is window is normal size (not min, not max)
    '---------------------------------------------
    Exit Sub
    If PIForm(gActivePIInstance).WindowState = 0 Then
        '---------------------------------------------
        ' remember the window settings
        '---------------------------------------------
        KeyValue = Format$(PIForm(gActivePIInstance).Top, "0")
        WriteProfile SectionName, "Top", KeyValue
        KeyValue = Format$(PIForm(gActivePIInstance).Left, "0")
        WriteProfile SectionName, "Left", KeyValue
        KeyValue = Format$(PIForm(gActivePIInstance).Width, "0")
        WriteProfile SectionName, "Width", KeyValue
        KeyValue = Format$(PIForm(gActivePIInstance).Height, "0")
        WriteProfile SectionName, "Height", KeyValue
    End If
    SectionName = "Settings"
    '---------------------------------------------
    ' remember the font settings
    '---------------------------------------------
    WriteProfile SectionName, "selfontname", PIForm(gActivePIInstance).MessageArea.SelFontName
    WriteProfile SectionName, "FontBold", PIForm(gActivePIInstance).MessageArea.SelBold
    WriteProfile SectionName, "FontItalic", PIForm(gActivePIInstance).MessageArea.SelItalic
    WriteProfile SectionName, "FontStrikeThru", PIForm(gActivePIInstance).MessageArea.SelStrikeThru
    WriteProfile SectionName, "FontSize", PIForm(gActivePIInstance).MessageArea.SelFontSize
    
    SectionName = "Remailer Info"
    '---------------------------------------------
    ' remember the current remailer pointed to on the list
    '---------------------------------------------
  
    '---------------------------------------------
    ' remember if adv-cypher encrypt to remailers option is checked
    '---------------------------------------------
   ' If gEncryptToRemailer Then
      '  WriteProfile SectionName, "RemailersPGP", "true"
    'Else
      '  WriteProfile SectionName, "RemailersPGP", ""
    'End If
    '---------------------------------------------
    ' remember if cypherpunk-style remailers is checked
    '---------------------------------------------
    If gRemailerType = REMAILER_CYPHERPUNK Then
        WriteProfile SectionName, "RemailersCP", "true"
        'If gEncryptToRemailer Then
         '   WriteProfile SectionName, "EncryptionToRemailers", "true"
       ' Else
        '    WriteProfile SectionName, "EncryptionToRemailers", ""
       ' End If
    Else
        WriteProfile SectionName, "RemailersCP", ""
    End If
    '---------------------------------------------
    ' remember if mixmaster-style remailers is checked
    '---------------------------------------------
    If gRemailerType = REMAILER_MIX Then
        WriteProfile SectionName, "RemailersMix", "true"
    Else
        WriteProfile SectionName, "RemailersMix", ""
    End If
   
    If gRemailerType = STANDARD_EMAIL Then
        WriteProfile SectionName, "StandardEmail", "true"
    Else
        WriteProfile SectionName, "StandardEmail", ""
    End If
    '---------------------------------------------
    ' news posters
    '---------------------------------------------
    SectionName = "News Posters"
    '---------------------------------------------
    ' remember if USENET option is checked as none
    '---------------------------------------------
    'If PIForm(gActivePIInstance).USENETNone.Checked Then
         'riteProfile SectionName, "USENETNone", "true"
    'Else
        'WriteProfile SectionName, "USENETNone", ""
    'End If
    '---------------------------------------------
    ' remember if USENET option anon-penet-FI is checked (disabled)
    '---------------------------------------------
    'If PIForm(gActivePIInstance).USENETFi.Checked Then
       ' WriteProfile SectionName, "USENETFI", "true"
    'Else
      '  WriteProfile SectionName, "USENETFI", ""
   ' End If
    '---------------------------------------------
    ' remember if USENET option e-mail gateway is checked
    '---------------------------------------------
    If PIForm(gActivePIInstance).USENETGate.Checked Then
        WriteProfile SectionName, "USENETGate", "true"
    Else
        WriteProfile SectionName, "USENETGate", ""
    End If
    '---------------------------------------------
    ' remember if USENET option eric is checked (disabled)
    '---------------------------------------------
   ' If PIForm(gActivePIInstance).UseNetSoda.Checked Then
      '  WriteProfile SectionName, "USENETSoda", "true"
    'Else
      '  WriteProfile SectionName, "USENETSoda", ""
    'End If
    '---------------------------------------------
    ' PGP settings
    '---------------------------------------------
    
    SectionName = "PGP Options"
    '---------------------------------------------
    ' remember if PGP encrypt to self is checked
    '---------------------------------------------
    If PIForm(gActivePIInstance).PGPSelf.Checked Then
        WriteProfile SectionName, "PGPSelf", "true"
    Else
        WriteProfile SectionName, "PGPSelf", ""
    End If
    '---------------------------------------------
    ' remember if PGP gObscurity is checked
    '---------------------------------------------
    If PIForm(gActivePIInstance).PGPObscurity.Checked Then
        WriteProfile SectionName, "PGPObscurity", "true"
    Else
        WriteProfile SectionName, "PGPObscurity", ""
    End If
    '---------------------------------------------
    ' remember if run PGP minimized is checked
    '---------------------------------------------
    If PIForm(gActivePIInstance).PGPMin.Checked Then
        WriteProfile SectionName, "PGPMin", "true"
    Else
        WriteProfile SectionName, "PGPMin", ""
    End If
    '---------------------------------------------
    ' remember if PGP auto-close is checked
    '---------------------------------------------
    'If PIForm(gActivePIInstance).PGPClose.Checked Then
      ' WriteProfile SectionName, "PGPClose", "true"
    'Else
     '   WriteProfile SectionName, "PGPClose", ""
    'End If
    '---------------------------------------------
    ' remember if PGP use multiple keys is checked
    '---------------------------------------------
    If PIForm(gActivePIInstance).PGPMultiple.Checked Then
        WriteProfile SectionName, "PGPMultiple", "true"
    Else
        WriteProfile SectionName, "PGPMultiple", ""
    End If
    '---------------------------------------------
    ' remember if PGP eyes only encrryption is checked
    '---------------------------------------------
    'If PIForm(gActivePIInstance).PGPEyes.Checked Then
    '    WriteProfile SectionName, "PGPEyes", "true"
   ' Else
    '    WriteProfile SectionName, "PGPEyes", ""
   ' End If
    '---------------------------------------------
    ' remember if PGP word wrap on encrypt is checked
    '---------------------------------------------
  '  If PIForm(gActivePIInstance).PGPWrap.Checked Then
     '   WriteProfile SectionName, "Wrap", "true"
    'Else
    '    WriteProfile SectionName, "Wrap", ""
   ' End If
    
    '---------------------------------------------
    ' message setting
    '---------------------------------------------
    '---------------------------------------------
    ' remember if message append-on-send is checked
    '---------------------------------------------
    'If gEncryptToRemailer Then
     '   WriteProfile SectionName, "Auto", "true"
   ' Else
     '   WriteProfile SectionName, "Auto", ""
   ' End If
    
    '---------------------------------------------
    ' test to see if we are exiting the program
    '---------------------------------------------
    If gExit = 1 Then
        
    End If
End Sub

Sub SendScript(Script As String)
       
    Dim parseScript As String
    Dim i As Integer
    Dim tmpstr As String
    Dim tmpstr2 As String
    i = 1
    parseScript = Script
    Do
        i = InStr(parseScript, ",")
        If i > 0 Then
            tmpstr = Mid$(parseScript, 1, i - 1)
        Else
            tmpstr = parseScript
        End If
        Select Case tmpstr
        Case "to"
            tmpstr2 = PIForm(gActivePIInstance).txtTo.Text
            Clipboard.SetText tmpstr2
            SendKeys "^v", True
            SendKeys "{TAB}", True
        Case "from"
            Clipboard.SetText MailConnector.EmailAddress
            SendKeys "^v", True
            SendKeys "{TAB}", True
        Case "subject"
            tmpstr2 = PIForm(gActivePIInstance).txtSubject.Text
            Clipboard.SetText tmpstr2
            SendKeys "^v", True
            SendKeys "{TAB}", True
        Case "cc"
            Clipboard.SetText PIForm(gActivePIInstance).txtCC.Text
            SendKeys "^v", True
            SendKeys "{TAB}", True
        'Case "bcc"
          '  Clipboard.SetText PIForm(gActivePIInstance).Text4.Text
           ' SendKeys "^v", True
            'SendKeys "{TAB}", True
        Case "msg"
            Clipboard.SetText PIForm(gActivePIInstance).MessageArea.Text
            SendKeys "^v", True
        Case "tab"
            SendKeys "{TAB}", True
        Case "+tab"
            SendKeys "+{TAB}", True
        Case "enter"
            SendKeys "~", True
        Case "right"
            SendKeys "{RIGHT}"
        Case "left"
            SendKeys "{LEFT}"
        Case "up"
            SendKeys "{UP}"
        Case "down"
            SendKeys "{DOWN}"
        Case "escape"
            SendKeys "{ESC}"
            'document +=shift, ^=ctrl, %=alt
            'otherwise, send whatever keys are specified
        Case Else
            SendKeys tmpstr, True
        End Select
        If i <> 0 Then
            parseScript = Mid$(parseScript, i + 1, Len(parseScript) - i)
        End If
    Loop Until i = 0
End Sub
Sub ShowWebHelp()
        
    If MsgBox("Complete help for Private i Mail is available at http://www.eskimo.com/~joelm/pihelp.html.  Select OK to put the URL in your clipboard.  Paste it in your browser to access the Web help page.  Once there, you can save the page to your hard drive for future reference.", 64 + 1, "Private i Mail Help") = 1 Then
        Clipboard.SetText "http://www.eskimo.com/~joelm/pihelp.html"
    End If
End Sub

Function StripName(theName As String) As String
Dim i As Integer
   
    If InStr(theName, " ") = 0 Then
    i = Len(theName)
Else
    i = InStr(theName, " ") - 1
End If
StripName = Mid$(theName, 1, i)
End Function

Sub TransferInfo(theTrans As String, theScript As String)
Dim CurrentWindow As Long
Dim Length As Integer
Dim ListItem As String
Dim iL As Long


    On Error GoTo TransferError
    
    CurrentWindow = GetWindow(PIForm(gActivePIInstance).hWnd, GW_HWNDFIRST)
    ' Loop while the hWnd returned by GetWindow is valid.
    While CurrentWindow <> 0
        'Get the length of the task name identified by
        'CurrentWindow in the list.
        Length = GetWindowTextLength(CurrentWindow)
        'Get the task name of the task in the master list.
        ListItem = Space$(Length + 1)
        Length = GetWindowText(CurrentWindow, ListItem, Length + 1)
        'If there is an actual task name in the list, add the
        'item to the list.
        If Length > 0 Then
            If InStr(1, ListItem, theTrans) Then
                iL = SetFocusAPI(CurrentWindow)
                iL = FindWindow(0, Mid$(ListItem, 1, Len(ListItem) - 1))
                SetWindowPos iL, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                SendKeys "% iL", True
                SendScript (theScript)
                Exit Sub
            End If
        End If
        'Get the next task list item in the master list.
        CurrentWindow = GetWindow(CurrentWindow, GW_HWNDNEXT)
        'Process Windows events.
        iL = DoEvents()
    Wend
  Exit Sub
TransferError:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " error in the transfer routine.  Application not available to activate.", 48, gPiStr
    End If
    Err.Clear
End Sub



Sub WriteProfile(pSectionName As String, pKeyName As String, pValue As String)
  
    SaveSetting AppName:=App.Title, section:=pSectionName, Key:=pKeyName, setting:=pValue
    
    'ReturnCode = WritePrivateProfileString(ByVal pSectionName, ByVal pKeyName, ByVal pValue, ByVal gAppINI)
    'If ReturnCode = 0 Then
       ' MsgBox "Bad WritePrivateProfileString " & pKeyName & "=" & pValue
    'End If
End Sub


Public Sub AppendNymInfo()
Dim J As Integer
Dim i As Integer
Dim SendHeader As String
On Error GoTo AppendError
    '---------------------------------------------
    'reset the gPGPKeyID, no signature on anonymous remailing
    '---------------------------------------------
    
    gPGPKeyID = ""
    'If Nym.UseNewsGroupReply Then
    If gRemailerType = STANDARD_EMAIL Then
        gnumRemailers = 0
        Exit Sub
    End If
     
    If Not gRemailerType = REMAILER_MIX Then
        
        If (frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) <> "chain") Then
            SendHeader = "::" & vbCrLf & "Request-Remailing-To: " & RTrim$(PIForm(gActivePIInstance).txtTo.Text) + vbCrLf
            If Not Nym.PassPhrase(1) = "" Then
                SendHeader = SendHeader & "Encrypt-Key: " & Nym.PassPhrase(1) & vbCrLf & vbCrLf
            End If
            PIForm(gActivePIInstance).MessageArea.SelStart = 0
            PIForm(gActivePIInstance).MessageArea.SelText = SendHeader
        Else
            frmRemailerChain.lblCaption = "You now have the opportunity to select the remailers that will be used in your reply block. " & vbCrLf
            frmRemailerChain.lblCaption = frmRemailerChain.lblCaption & "In the final step, you will be given the opportunity to send this Nym configuration request through a remailer."
            frmRemailerChain.Caption = "Select a remailer for your reply block"
            frmRemailerChain.Show vbModal
            If gCancelAction Then Exit Sub
            If gnumRemailers <> 0 Then
                    'This first one is to the nym ie remailer 1
                    SendHeader = "::" & vbCrLf & "Request-Remailing-To: " & RTrim$(PIForm(gActivePIInstance).txtTo.Text) & vbCrLf
                    If Not Nym.PassPhrase(gnumRemailers + 1) = "" Then
                        SendHeader = SendHeader & "Encrypt-Key: " & Nym.PassPhrase(gnumRemailers + 1) & vbCrLf & vbCrLf
                    Else
                        SendHeader = SendHeader & vbCrLf
                    End If
                    PIForm(gActivePIInstance).MessageArea.SelStart = 0
                    PIForm(gActivePIInstance).MessageArea.SelText = SendHeader
                    For J = gnumRemailers To 1 Step -1
                        gwhichRemailer = J
                        vb2spgpContext.Initialise
                        vb2spgpContext.Armor = 1
                        vb2spgpContext.KeyEncrypt = 1
                        'vb2spgpContext.TextMode = 0
                        '[AC 2/3/2002]
                        vb2spgpContext.TextMode = 1
                        vb2spgpContext.CryptKeyID = RTrim$(Remailers(J))
                        spgpEncryptMessage
                        'EncryptMessage RTrim$(Remailers(j)), "-eatw"
                        PIForm(gActivePIInstance).MessageArea.SelStart = Len(PIForm(gActivePIInstance).MessageArea)
                        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf & "**" & vbCrLf
                        PIForm(gActivePIInstance).MessageArea.SelStart = 0
                        If J <> 1 Then
                            SendHeader = "::" & vbCrLf & "Request-Remailing-To: " & RTrim$(Remailers(J)) & vbCrLf
                            If Not Nym.PassPhrase(J) = "" Then
                                SendHeader = SendHeader & "Encrypt-Key: " & Nym.PassPhrase(J) & vbCrLf & vbCrLf
                            Else
                                SendHeader = SendHeader & vbCrLf
                            End If
                            PIForm(gActivePIInstance).MessageArea.SelText = SendHeader
                        End If
                        PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf & "Encrypted: PGP" & vbCrLf & vbCrLf
                    Next
                    gwhichRemailer = 0
                'End If
            End If
        End If

        If gnumRemailers <> 0 Then
            If (gRemailerType = REMAILER_CYPHERPUNK) Then
                If frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) = "chain" Then
                    PIForm(gActivePIInstance).txtTo.Text = RTrim(Remailers(1))
                Else
                    PIForm(gActivePIInstance).txtTo.Text = RTrim(StripName(frmRemailerList.List3.List(frmRemailerList.List3.ListIndex)))
                End If
            End If
        '---------------------------------------------
        'encrypt for first remailer in the chain if PGP is on
        '---------------------------------------------
        'If gEncryptToRemailer = True And frmRemailerList.list3.List(frmRemailerList.list3.ListIndex) <> "chain" Then
            If frmRemailerList.List3.List(frmRemailerList.List3.ListIndex) <> "chain" Then
                EncryptMessage PIForm(gActivePIInstance).txtTo.Text, "-eatw"
                PIForm(gActivePIInstance).MessageArea.SelStart = Len(PIForm(gActivePIInstance).MessageArea)
                        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf & "**" & vbCrLf
                PIForm(gActivePIInstance).MessageArea.SelStart = 0
                PIForm(gActivePIInstance).MessageArea.SelText = "::" & vbCrLf + "Encrypted: PGP" & vbCrLf & vbCrLf
            
            End If
        End If
    PIForm(gActivePIInstance).MousePointer = vbDefault
End If
Exit Sub
AppendError:
    MsgBox "An error occured while configuring the Remailer template.  The error returned was " & Err.Description & " in AppendInfo", vbCritical + vbApplicationModal, App.Title
    Err.Clear
    gCancelAction = True
End Sub

Public Sub spgpEncryptMessageArea()
  Dim pEncryptID As String * 256
  Dim pSignID As String * 256
  Dim pConventionalPass As String * 256
  Dim pComment As String * 256
  Dim pSignPass As String * 256
  'Dim Key As TKey_Data
  Dim i As Long
  Dim sResponse As String
  Dim msg As String
  Dim pHashAlg As Long
  Dim BufferOut As String
  Dim BufferIn As String
  Dim pConventionalAlg As Long
  Dim pTextMode As Long
  Dim pClearSign As Long
  Dim pConventionalEncrypt As Long
  Dim pArmor As Long
  Dim pSign As Long
  Dim pEncrypt As Long
  Dim spgperr As String * 256
  Dim pCompress As Long
  Dim pMIME As Long
  Dim pMIMESeparator As String
  Dim pEyesOnly As Long
  
  pMIME = vb2spgpContext.MIME
  pEyesOnly = vb2spgpContext.EyesOnly
  pCompress = vb2spgpContext.Compress
  
  BufferIn = PIForm(gActivePIInstance).MessageArea.Text '& Chr(0)
  BufferOut = String((Len(BufferIn) + 8096), Chr(0))
  pEncryptID = vb2spgpContext.CryptKeyID
  pSignID = vb2spgpContext.SignKeyID
  pSign = vb2spgpContext.Sign
  'pSign = 1
  pClearSign = vb2spgpContext.Clear
  
  If pSign = 1 Then
        
        'sResponse = InputBox("In order to sign the message you need to enter your passphrase. ", "Passphrase required for " & vb2spgpContext.SignKeyID)
        sResponse = myInputBox("In order to sign the message you need to enter your passphrase. ", "Passphrase required for " & vb2spgpContext.SignKeyID)
        If sResponse = "" Then
            gCancelAction = True
            Exit Sub
        End If
        vb2spgpContext.SignKeyPass = sResponse
        pSignPass = vb2spgpContext.SignKeyPass
   End If
  'If pClearSign = 1 Then
      '  vb2spgpContext.SignKeyPass = InputBox("In order to sign the message you need to enter your passphrase. ", "Passphrase required for " & vb2spgpContext.SignKeyID)
       ' pSignPass = vb2spgpContext.SignKeyPass
  'End If
  If vb2spgpContext.ConventionalEncrypt Then
    pConventionalEncrypt = vb2spgpContext.ConventionalEncrypt
    vb2spgpContext.ConventionalPass = myInputBox("You have requested that the message be conventionally encrypted.  Please enter the passphrase you wish to encrypt this message with: ", "Conventional Passphrase")
    pConventionalPass = vb2spgpContext.ConventionalPass
  End If
  
  
  pArmor = vb2spgpContext.Armor
  pTextMode = vb2spgpContext.TextMode
  pHashAlg = vb2spgpContext.PGPHashAlgorithm '   PGPHashAlgorithm_MD5 '
  pConventionalAlg = vb2spgpContext.PGPCipherAlgorithm
  pEncrypt = vb2spgpContext.KeyEncrypt
  '
  ' a fancy touch: put the signing key's key-id in the "Comment:" string
  If pClearSign = 1 Then
    i = spgpKeyProps(pSignID, BufferOut, Len(BufferOut))
    Key = ParseKeyData(BufferOut)
    pComment = "Signed by Key " & Key.KeyID & Chr(0)
    'BufferOut = String(Len(bufferin) * 2, Chr(0))
  End If
  'pSignPass = "dsfa"
  i = spgpEncode(BufferIn, BufferOut, Len(BufferOut), pEncrypt, pSign, pHashAlg, pConventionalEncrypt, pConventionalAlg, pArmor, pTextMode, pClearSign, pCompress, pEyesOnly, pMIME, pEncryptID, pSignID, pSignPass, pConventionalPass, pComment, pMIMESeparator)
  If Not i = 0 Then
    Call spgpGetErrorString(i, spgperr)
    MsgBox "Error encrypting message:  Error returned was: " & StripNulls(spgperr), vbApplicationModal + vbCritical, "Encryt Message"
    Beep
  Else
   PIForm(gActivePIInstance).MessageArea.SelText = 0
   PIForm(gActivePIInstance).MessageArea.Text = StripNulls(BufferOut)
  End If
End Sub

Public Function spgpDecryptMessage() As Integer
' all strings must be of fixed length
  Dim BufferIn As String
  Dim BufferOut As String   ' passes clear-text, receives cipher-text
  Dim pPass As String * 256
  Dim SigProps As String * 256
  Dim spgperr As String * 256
  Dim UserID As String * 256
  'Dim Sig As TSig_Data
  'Dim Key As TKey_Data
  Dim AnalysisData As Long
  Dim i As Long
  Dim J As Integer
 
  On Error GoTo BadDecrypt
 ' BufferIn = String(Len(PIForm(gActivePIInstance).MessageArea.Text & Chr(0)), Chr(0)) ' final null is required
  BufferIn = PIForm(gActivePIInstance).MessageArea.Text & Chr(0)
  BufferOut = String(Len(BufferIn) + 1024, Chr(0))
  SigProps = "" & Chr(0)         ' output buffers should be initialised & terminated
  pPass = GetPassPhrase(0, BufferIn)
  If Trim(pPass) = "" Then
    spgpDecryptMessage = vbCancel
    Exit Function
  End If
  i = spgpDecode(BufferIn, BufferOut, Len(BufferOut), pPass, SigProps)
  If Not i = 0 Then
    Call spgpGetErrorString(i, spgperr)
    Err.Raise 2000, "An error has occurred Decrypting the Message: ", spgperr
    Beep
    PIForm(gActivePIInstance).ShowStatus 1, Err.Description
  Else
    SignatureProperties = ParseSigData(SigProps)
    If SignatureProperties.Status = "SIGNED_BAD" Then
        Err.Raise 1001, "Bad signature.", "Bad signature detected."
        PIForm(gActivePIInstance).ShowStatus 1, Err.Description
    Else
        spgpDecryptMessage = vbOK
        PIForm(gActivePIInstance).MessageArea.SelStart = 0
        PIForm(gActivePIInstance).MessageArea.Text = StripNulls(BufferOut)
    End If
End If
Exit Function
BadDecrypt:
    spgpDecryptMessage = vbCancel
    MsgBox "Error decrytping message: " & Err.Description, vbApplicationModal + vbCritical, "Decrypt Error"
    Err.Clear
End Function



Public Function spgpEncryptfile() As Boolean
Dim pCryptKeyID As String * 256
Dim pSignKeyID As String * 256
Dim pSignKeyPass As String * 256
Dim pConventionalPass As String * 256
Dim pComment As String * 256
Dim spgperr As String * 256
Dim i As Long
Dim pHashAlg As Long
Dim pFileIn As String
Dim pFileOut As String
Dim pEncrypt As Long
Dim pSign As Long
Dim pArmor As Long
Dim pConventionalEncrypt As Long
Dim pConventionalAlg As Long
Dim pTextMode As Long
Dim pClearSign As Long
Dim pMIME As Long
Dim pEyesOnly As Long
Dim pMIMESeparator As Long
Dim pCompress As Long


  pMIME = vb2spgpContext.MIME
  pEyesOnly = vb2spgpContext.EyesOnly
  pCompress = vb2spgpContext.Compress
  pEncrypt = vb2spgpContext.KeyEncrypt
  pSignKeyID = vb2spgpContext.SignKeyID
  pSign = vb2spgpContext.Sign
  pClearSign = vb2spgpContext.Clear
  
  If pSign = 1 Or pClearSign = 1 Then
        vb2spgpContext.SignKeyPass = myInputBox("In order to sign the message you need to enter your passphrase. ", "Passphrase required for " & vb2spgpContext.SignKeyID)
        pSignKeyPass = vb2spgpContext.SignKeyPass
   End If 'If iFileExists(StripExt(pFileIn) & ".asc") Then
  
  If vb2spgpContext.FileOut = "" Then
    pFileOut = vb2spgpContext.FileIn & ".asc"
    If iFileExists(pFileOut) Then
        Kill pFileOut
        PIForm(gActivePIInstance).ShowStatus 1, "Deleting file " & pFileIn & ".asc" & vbNullChar
    End If
    vb2spgpContext.FileOut = pFileOut
  Else
    pFileOut = vb2spgpContext.FileOut & Chr(0)
  End If
  
  
 ' pFileIn = TempPathLocation & vb2spgpContext.FileIn & vbNullChar
 ' pFileOut = StripExt(pFileIn) & ".asc"
 pFileIn = vb2spgpContext.FileIn & vbNullChar
  pCryptKeyID = vb2spgpContext.CryptKeyID
 'pSignKeyID = vb2spgpContext.SignKeyID
  pSignKeyPass = vb2spgpContext.SignKeyPass
  
  If vb2spgpContext.ConventionalEncrypt Then
    If Not pConventionalPass = "" Then
        pConventionalEncrypt = vb2spgpContext.ConventionalEncrypt
        vb2spgpContext.ConventionalPass = myInputBox("This file, message or attachment will be conventionally encrypted.  Please enter the passphrase you wish to Encrypt this item with: ", "Encrypt File")
        pConventionalPass = vb2spgpContext.ConventionalPass
    End If
  End If
  pConventionalAlg = vb2spgpContext.PGPCipherAlgorithm
  pHashAlg = vb2spgpContext.PGPHashAlgorithm '0 'PGPHashAlgorithm_MD5 'PGPHashAlgorithm_SHA
  'If pConventionalAlg = 0 Then pConventionalAlg = 1
  pSign = vb2spgpContext.Sign
  pArmor = vb2spgpContext.Armor
  pTextMode = vb2spgpContext.TextMode
  pComment = vb2spgpContext.Comment & vbNullChar
  

DoEvents

i = spgpEncodeFile(pFileIn, _
                   pFileOut, _
                   pEncrypt, _
                   pSign, pHashAlg, _
                   pConventionalEncrypt, _
                   pConventionalAlg, _
                   pArmor, _
                   pTextMode, _
                   pClearSign, _
                   pCompress, _
                   pEyesOnly, _
                   pMIME, _
                   pCryptKeyID, _
                   pSignKeyID, _
                   pSignKeyPass, _
                   pConventionalPass, _
                   pComment, _
                   pMIMESeparator)
   
 DoEvents
 vb2spgpContext.MIMESeparator = pMIMESeparator
 If i = 0 Then
    spgpEncryptfile = True
 Else
    Call spgpGetErrorString(i, spgperr)
    Err.Raise 1000, " - (Error in pEncrypt Message) - ", spgperr
    Beep
    PIForm(gActivePIInstance).ShowStatus 1, Err.Description
    spgpEncryptfile = False
 End If

End Function
Public Sub spgpDecryptFile()
 ' all strings must be of fixed length
  Dim pSigProps As String * 512
  Dim pPass As String * 256
  Dim pFileIn As String * 256
  Dim pFileOut As String * 256
  Dim iRes As Long
  Dim spgperr As String * 256
  'Dim Sig As TSig_Data
  Dim AnalysisData As Long
  
    pSigProps = "" & Chr(0)
    pFileIn = vb2spgpContext.FileIn & Chr(0)
    pFileOut = vb2spgpContext.FileOut & Chr(0)
    pPass = GetPassPhrase(1, pFileIn)
    iRes = spgpDecodeFile(pFileIn, pFileOut, pPass, pSigProps)
    If Not iRes = 0 Then
        spgperr = spgperr & Chr(0)
        Call spgpGetErrorString(iRes, spgperr)
        Err.Raise 2010, "An error has occurred: ", spgperr
    End If
 
 End Sub

Public Function GetPassPhrase(SourceType As Long, Source As String) As String
  Dim i As Integer
  Dim AnalysisData As Long
  Dim sKeyID As String

  'Need to do this to dimension the array..
  ReDim KeyArray(0)
  Call GetKeyData(SourceType, Source & Chr(0), KeyArray, AnalysisData)
  
   ' if
    For i = 1 To UBound(KeyArray)
        If i = 1 Then
            sKeyID = KeyArray(1).UserID
        Else
            sKeyID = sKeyID & vbCrLf & KeyArray(i).UserID
        End If
    Next
    sKeyID = sKeyID & vbCrLf & vbCrLf
  
  Select Case AnalysisData
    Case 0 'Encrypted
        GetPassPhrase = myInputBox("To be able to decrypt the message, you need to enter your passphrase for on or more of the following ID/IDs: " & vbCrLf & vbCrLf & sKeyID, "UserID = " & Key.UserID)
    Case 1 'Signed
        GetPassPhrase = myInputBox("To be able to decrypt the message, you will need to enter your passphrase for on or more of the following ID/IDs: " & vbCrLf & vbCrLf & sKeyID, "UserID = " & Key.UserID)
    Case 5
        GetPassPhrase = myInputBox("This message appears to be conventionally encrypted.  You need to enter a passphrase to be able to decrypt the message. ", "Passphrase Required")
    Case 6
        Err.Raise 2012, , "You do not have the key on your keyring to decrypt this message."
    End Select
    GetPassPhrase = IIf(GetPassPhrase = "", "", GetPassPhrase & Chr(0))
End Function

Public Function GetRemailerID() As String
GetRemailerID = PIForm(gActivePIInstance).StatusBar.Panels.Item(3)
End Function

Public Function GetTemporaryFile() As String
GetTemporaryFile = App.Path & "\temp\" & CInt(Rnd(Time()) * 10000) & ".tmp"
End Function

Public Sub KillTemporaryFiles()
On Error Resume Next
Kill App.Path & "\temp\" & "*.tmp"
End Sub

Public Function KeyOnKeyRing(msgKey As String) As Boolean
Dim i As Integer
Dim J As Integer
'msgKey = "alex@itech.net.au"
'If Not InStr(1, msgKey, "<") = 0 Then
    KeyOnKeyRing = IIf(spgpKeyIsOnRing(msgKey) = 0, True, False)
'Else
 '   KeyOnKeyRing = IIf(spgpKeyIsOnRing("<" & msgKey & ">") = 0, True, False)
'End If
End Function

Public Sub spgpEncryptMessage()
'First check length if > 22,000 do file encrypt then exit
  Dim bRes As Integer
  If Len(PIForm(gActivePIInstance).MessageArea.Text) > 22000 Then
        vb2spgpContext.FileIn = GetTemporaryFile
        spgpEncryptMessageAreaToFile
        PIForm(gActivePIInstance).MessageArea.Text = GetFileText(vb2spgpContext.FileOut)
        KillTemporaryFiles
  Else
    spgpEncryptMessageArea
  End If
  
End Sub
'---------------------------
' Assumes filein has the file name
'-----------------------------
Public Sub spgpEncryptMessageAreaToFile()
Dim bRes As Boolean

bRes = PutFileText(vb2spgpContext.FileIn, PIForm(gActivePIInstance).MessageArea.Text)
vb2spgpContext.FileOut = GetTemporaryFile
If Not spgpEncryptfile Then
        KillTemporaryFiles
        Err.Raise 10004, , "Error encrypting message - in EncryptMessage."
End If

End Sub
