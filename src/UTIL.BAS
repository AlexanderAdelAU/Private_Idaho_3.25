Attribute VB_Name = "UTIL"
 Option Explicit
Sub SetTabStops(ListWindowHandle As Long, iNoTabStops As Integer, TabPositions() As Long)
   
    '
    '   Definition of all local variables
    '
    Dim iIndex As Integer
    Dim DialogBaseUnits As Long
    Dim DialogUnitsX As Long
    Dim iTabPos() As Integer        '   Tab positions
    '
    '   Obtain the dialog base units
    '
    DialogBaseUnits = GetDialogBaseUnits()
    DialogUnitsX = (DialogBaseUnits And &HFFFF&) / 4
    '
    '   Calculate the "required" tab positions
    '
    ReDim iTabPos(0 To iNoTabStops - 1)
    For iIndex = 0 To iNoTabStops - 1
        iTabPos(iIndex) = TabPositions(iIndex)
    Next iIndex
       
    '
    '   Set the required tab positions in the list box,
    '   by calling the "SendMessage" api and passing the
    '   handle of the listbox and details of the number of
    '   tabs and their positions.
    '
        Call SendMessageX(ListWindowHandle, LB_SETTABSTOPS, iNoTabStops, iTabPos(0))
End Sub

Sub SortCPRemailers()
    
    Dim x
    Dim y As Integer
    'Static matchedRemailers(50) As Remailer
    y = 0
    For x = 1 To gTotalRemailers
        If gRemailerArray(x).cpunk = 1 Then
            If (gRemailerArray(x).Encrypt = gSortRemailer.Encrypt) And (gRemailerArray(x).post = gSortRemailer.post) And (gRemailerArray(x).latent = gSortRemailer.latent) And (gRemailerArray(x).hash = gSortRemailer.hash) And (gRemailerArray(x).cut = gSortRemailer.cut) = True Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        End If
    Next
End Sub


Sub ExecCmd(cmdline As String)
   
    
   Dim pid As Long
  ' Dim sDir As String
   On Error Resume Next
   pid = Shell(cmdline, vbNormalFocus)
   DoEvents
   If Not pid = 0 Then
        WaitForTerm1 pid
   End If
   Exit Sub
'test:
  ' res = CreateProcessByNum(gMixPath & "\mixmaste.exe", cmdline, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0, gMixPath, sinfo, pinfo)
   'g
   'gmixpath & "\mixmaste.exe"
   'res=CreateProcessByNum(gmixpath & "\mixmaste.exe",cmdline,0,0,)
   
   'If res Then
    '    WaitForTermination pinfo
   ' End If
   ' DoEvents
End Sub
Public Sub WaitForTerm1(pid As Long)
Dim phnd As Long
phnd = OpenProcess(SYNCHRONIZE, 0, pid)
DoEvents
If Not phnd = 0 Then
    Call WaitForSingleObject(phnd, 100000)
    Call CloseHandle(phnd)
End If
End Sub

Private Sub WaitForTermination(pinfo As PROCESS_INFORMATION)
Dim res As Long
Dim i As Integer
Call WaitForInputIdle(pinfo.hProcess, INFINITE)
'Call CloseHandle(pinfo.hThread)
'res = WaitForSingleObject(pinfo.hProcess, 0)
'This forces the DOS window to the foreground
'For i = 0 To 3
 '   DoEvents
'Next
'If res Then
    'If gPGPResponse.Count <> 0 Then
       ' For i = 0 To gPGPResponse.Count - 1
          '  MySendKeys gPGPResponse.res(i)
          '  DoEvents
          '  DoEvents
       ' Next
   'End If
'End If
'DoEvents
Do
    res = WaitForSingleObject(pinfo.hProcess, 0)
    If Not res = WAIT_TIMEOUT Then
        Exit Do
    End If
    DoEvents

    
Loop While True
Call CloseHandle(pinfo.hProcess)

End Sub

Sub CheckLen(TheString As String)
      
    If Len(TheString) > 128 Then
        MsgBox "The PGP command line is" + Str(Len(TheString)) + " characters.  DOS will not allow you to send that large a command to PGP.  Try getting the command line down to less than 128 characters.  The PGP command will be processed anyway, but it won't receive all the command options you specified. You may have to re-install into a more suitable directory. See the PI User Manual for more information.", 48, gPiStr
        gCancelAction = True
    End If
End Sub
Function CountChar(TheStr As String, theChar As String) As Integer '
    Dim i As Integer
    Dim J As Integer
    
    i = 0
    For J = 1 To Len(TheStr)
        If Mid$(TheStr, J, 1) = theChar Then i = i + 1
    Next
    CountChar = i
End Function

Function fGetLine(LineNumber As Long) As String
    Dim byteLo As Integer
    Dim byteHi As Integer
    Dim Buffer As String
    Dim Ltmp As Long
    
    byteLo = MAX_CHAR_PER_LINE And (255)
    byteHi = Int(MAX_CHAR_PER_LINE / 256)
    Buffer = Chr$(byteLo) + Chr$(byteHi) + Space$(MAX_CHAR_PER_LINE - 2)
    Ltmp = SendMessageByString(PIForm(gActivePIInstance).MessageArea.hWnd, EM_GETLINE, LineNumber, Buffer)
    'x% = SendMessage(Form1.Text5.hwnd, EM_GETLINE, LineNumber, Buffer$)
    '---------------------------------------------
    'returns the text on the line with no spaces appended
    '---------------------------------------------
    fGetLine = Left$(Buffer, CInt(Ltmp))
End Function

Function fGetLineCount() As String
    Dim lcount As Long
    
    lcount = SendMessage(PIForm(gActivePIInstance).MessageArea.hWnd, EM_GETLINECOUNT, 0, 0)
    fGetLineCount = lcount
End Function




Sub HideStatus()
    Unload frmStatus
End Sub

Function iFileExists(ByVal sFileName As String) As Boolean
'Dim clsFileFunctions As New clsFileFunctions
'Dim objFiles As New Collection
'Dim objFile As File
'On Error Resume Next

'Set objFiles = clsFileFunctions.FindFile(StripFileName(sFileName), StripPath(sFileName))
'For Each objFile In objFiles
   ' PIForm(gActivePIInstance).List1.AddItem objFile.Path
'Next
'If objFiles.Item(1) = "" Then
'    iFileExists = False
'Else
'    iFileExists = True
'End If
'Exit Function


Dim i As Integer
Dim FileNum As Integer
    Err.Clear
    On Error Resume Next
    i = Len(Dir(sFileName))
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

Function InsertCRLFs() As String
    Dim Buffer As String
    Dim NumIndex As Integer
    Dim i As Integer
        
    Buffer = ""
    '---------------------------------------------
    'get the number of lines in the message window
    '---------------------------------------------
    NumIndex = fGetLineCount()
    '---------------------------------------------
    'add CR+LF to all the lines
    '---------------------------------------------
    For i = 1 To NumIndex
        '---------------------------------------------
        'take whatever text there is, strip off trailing
        'blanks, add a CR+LF and append that to any
        'existing formatted text.
        '---------------------------------------------
        Buffer = Buffer & fGetLine(i - 1) '& vbCrLf
    Next i
    '---------------------------------------------
    'return the formatted buffer
    '---------------------------------------------
    InsertCRLFs = Buffer
End Function

Sub NotYet()
      MsgBox "Not implemented yet.", vbApplicationModal, gPiStr
End Sub

Function ParseString(TheStr As String, delimiter As String) As String
      
    TheStr = RTrim$(TheStr)
    If TheStr = "" Then
        ParseString = ""
    Else
        If InStr(1, TheStr, delimiter) = 0 Then
            ParseString = TheStr
        Else
            ParseString = Mid$(TheStr, 1, InStr(1, TheStr, delimiter) - 1)
        End If
    End If
End Function
Sub PrepareFeedback()
    Dim lVer As Long
    Dim iVer As Integer
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim FileNum As Integer
    Dim App As New CApplication
    
    On Error GoTo PIError
    If Len(PIForm(gActivePIInstance).MessageArea.Text) > 0 Then
        '---------------------------------------------
        'test first to see if message area is empty
        '---------------------------------------------
        tmpstr = MsgBox("The message area contains text.  Is it okay to overwrite it?", vbYesNo, "Send Feedback")
        If tmpstr = vbYes Then
            PIForm(gActivePIInstance).MessageArea.Text = ""
        End If
    End If

    If Len(PIForm(gActivePIInstance).MessageArea.Text) = 0 Then
        '---------------------------------------------
        'the rest is systems info to dump in msg area
        '---------------------------------------------
       ' PIForm(gActivePIInstance).txtTo.BackColor = vbBlack
        PIForm(gActivePIInstance).txtTo.Text = ""
        PIForm(gActivePIInstance).txtsubject.Text = ""
'        PIForm(gActivePIInstance).txtTo.Text = "support@ptech.simplenet.com"
'        PIForm(gActivePIInstance).txtSubject.Text = "Private i Mail Automated Feedback"
        PIForm(gActivePIInstance).MessageArea.SelText = "Private i Mail Diagnostics" & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "A1: Private Idaho Email Version = " & App.Version & vbCrLf
        '========Get Serial Number============
        Dim rs As Recordset
        Dim HexSerialNumber As String
        Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
        If Not rs.EOF Then
            HexSerialNumber = IIf(IsNull(rs("Serial Number")), "0000", rs("Serial Number"))
        End If
        rs.Close
        PIForm(gActivePIInstance).MessageArea.SelText = "A2: Database Serial Number = " & HexSerialNumber & vbCrLf
       
        lVer = GetVersion()
        iVer = CInt(lVer / &H10000)
        
        PIForm(gActivePIInstance).MessageArea.SelText = "A3:  DOS: " + Format$(iVer / 256) + "." + Format$(iVer And &HFF) + gCRLF
        iVer = CInt(lVer And &HFFFF&)
        PIForm(gActivePIInstance).MessageArea.SelText = "A4:  Windows Version: " + Format$(iVer And &HFF) + "." + Format$(CInt(iVer / 256)) + gCRLF + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "Winsock Operations Info" + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" + gCRLF
        PIForm(gActivePIInstance).HTTP1.WinsockLoaded = True
        PIForm(gActivePIInstance).MessageArea.SelText = "B1:  Winsock Type: " + PIForm(gActivePIInstance).HTTP1.WinsockInfo + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "B2:  Winsock Status: " + PIForm(gActivePIInstance).HTTP1.WinsockStatus + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "B3:  Winsock Path: " + PIForm(gActivePIInstance).HTTP1.WinsockPath + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "B4:  Winsock Loaded: " + Str(PIForm(gActivePIInstance).HTTP1.WinsockLoaded) + gCRLF + gCRLF
        PIForm(gActivePIInstance).HTTP1.WinsockLoaded = False
        PIForm(gActivePIInstance).MessageArea.SelText = "Other Info" + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" + gCRLF
        tmpstr2 = GetSystemMetrics(SM_CXSCREEN) & "x" & GetSystemMetrics(SM_CYSCREEN)
        PIForm(gActivePIInstance).MessageArea.SelText = "C1:  Screen size: " + tmpstr2 + gCRLF
        If gMixPath = "" Then
            PIForm(gActivePIInstance).MessageArea.SelText = "C3:  Mixmaster path is not set" & vbCrLf
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "C3:  Mixmaster path is " & gMixPath & vbCrLf
        End If
        
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" + gCRLF
        
        If iFileExists(gMixPath & "\MIXMASTE.EXE") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D10:  MIXMASTE.exe is in the Mixmaster path" + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D10:  MIXMASTE.exe IS NOT in the Mixmaster path" + gCRLF
        End If
         If iFileExists(App.Path + "\mixmaster.htm") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D11:  mixmaster.htm is in the application path" + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D11:  mixmaster.htm IS NOT in the application path" + gCRLF
        End If
        If iFileExists(gMixPath & "\type2.lis") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D12:  Type2.lis is in the Mixmaster path" + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D12:  Type2.lis IS NOT in the Mixmaster path" + gCRLF
        End If
        If iFileExists(gMixPath & "\pubring.mix") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D13:  Pubring.mix is in the Mixmaster path" + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D13:  Pubring.mix IS NOT in the Mixmaster path" + gCRLF
        End If
       
        If iFileExists(App.Path + "\Remailer.txt") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D14:  Remailer.txt is in the application path" + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D14:  Remailer.txt IS NOT in the application path" + gCRLF
        End If
                
        If iFileExists(App.Path + "\Remailer.HTM") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D21:  REMAILER.HTM is in the application path" + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D21:  REMAILER.HTM IS NOT in the application path" + gCRLF
        End If
       
       ' If iFileExists(App.Path + "\URLS.TXT") Then
          '  PIForm(gActivePIInstance).MessageArea.SelText = "D26:  URLS.TXT is in the application path" + gCRLF
       ' Else
         '   PIForm(gActivePIInstance).MessageArea.SelText = "D26:  URLS.TXT IS NOT in the application path" + gCRLF
       ' End If
        If iFileExists(App.Path + "\USENET.TXT") Then
            PIForm(gActivePIInstance).MessageArea.SelText = "D27:  USENET.TXT is in the application path" + gCRLF + gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "D27:  USENET.TXT IS NOT in the application path" + gCRLF + gCRLF
        End If
        Dim SectionName As String
        SectionName = "Options"
        PIForm(gActivePIInstance).MessageArea.SelText = "Server Email Settings" + gCRLF
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "E1:  E-Mail Address: " & ReadProfile(SectionName, "EmailAddress") & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "E2:  Real Name: " & ReadProfile(SectionName, "RealName") & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "E3:  SMTP Server Name: " & ReadProfile(SectionName, "SMTPServerName") & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "E4:  Account Name: " & ReadProfile(SectionName, "AccountName") & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "E5:  Server Name: " & ReadProfile(SectionName, "MailServerName") & vbCrLf
        If ReadProfile(SectionName, "AccountPassword") = "" Then
            PIForm(gActivePIInstance).MessageArea.SelText = "E5:  Password: " & "" & gCRLF
        Else
            PIForm(gActivePIInstance).MessageArea.SelText = "E5:  Password: " & "***********" & gCRLF
        End If
        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "==================PGP Status==============" & vbCrLf
        SectionName = "Options"
        tmpstr = ReadProfile(SectionName, "PGPStatus")
        
        On Error Resume Next
        Select Case tmpstr
            Case NoPGP
                PIForm(gActivePIInstance).MessageArea.SelText = "NoPGP - PGP has been deselected."
            Case PGPNotFound
                PIForm(gActivePIInstance).MessageArea.SelText = "PGPNotFound - PGP 5/6 SDK or spgp.dll was not found."
            Case PGP5x
                PIForm(gActivePIInstance).MessageArea.SelText = "PGP5x - PGP 5/6 SDK or spgp.dll was found." & vbCrLf
                 PIForm(gActivePIInstance).MessageArea.SelText = "Version " & spgpVersion & " of spgp.dll detected. " & vbCrLf
                 PIForm(gActivePIInstance).MessageArea.SelText = "Version " & spgpSdkApiVersion & " of PGPSDK.dll detected. " & vbCrLf
                 
        End Select
        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
        
        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = vbCrLf
        
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "Please submit the info you feel comfortable releasing" & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = " when requesting technical support or when submitting" & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "comments.  The message is pre-addressed.  Thanks.. :)" & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SelText = "--------------------------------------------------" & vbCrLf & vbCrLf
        PIForm(gActivePIInstance).MessageArea.SetFocus
        PIForm(gActivePIInstance).MessageArea.SelStart = Len(PIForm(gActivePIInstance).MessageArea.Text)
        PIForm(gActivePIInstance).MessageArea.SelText = "Hi, here are my thoughts concerning PI32" & vbCrLf & vbCrLf
        PIForm(gActivePIInstance).txtTo.Text = "pisupport@itech.net.au"
        
    End If
    Exit Sub
PIError:
       PIForm(gActivePIInstance).MessageArea.SelText = "A1:  Error " + Err + " occurred " + Err.Description + gCRLF
    Err.Clear
End Sub

Sub ShowStatus(TheMessage As String)
    '---------------------------------------------
    'use whatever string is passed for message and show
    '---------------------------------------------
    frmStatus.lblTransferStatus.Caption = TheMessage
    frmStatus.Show
End Sub

Sub StatusMessage(TheString As String)
    
    frmStatus.lblTransferStatus.Caption = TheString
End Sub

Function StripExt(TheStr As String) As String
    Dim i As Integer
    On Error Resume Next
    i = InStrRev(TheStr, ".")
    StripExt = Left(TheStr, IIf(i = 0, Len(TheStr), i - 1))
    'StripExt = ""
End Function

Sub WipeFile(FileName As String)
    Dim Block1 As String
    Dim Block2 As String
    Dim Blocks As Long
    Dim hFile As Integer
    Dim lcounter As Long
    Dim Offset As Long
    Const BLOCKSIZE = 4096
    On Error Resume Next
    
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    hFile = FreeFile
    Open FileName For Binary As hFile
    Blocks = (LOF(hFile) \ BLOCKSIZE) + 1
    For lcounter = 1 To Blocks
        Offset = Seek(hFile)
        Put hFile, , Block1
        Put hFile, Offset, Block2
    Next lcounter
    Close hFile
    Kill FileName
End Sub



'
' Sends a single character using keybd_event
'   Note that this function does not set shift state
'   (By pressing down the shift key or setting the shift keys state)
'   and it doesn't handle extended keys.
'
Public Sub SendAKey(ByVal c$)
    Dim VK_Read As Long
    Dim vk As Integer
    Dim scan As Integer
    Dim oemchar As String
    Dim Shift As Integer
    Dim Control As Integer
   ' Dim dl&
    
    ' Get the virtual key code for this character
    If c$ = Chr(13) Then
       ' Control = 1
        'c$ = "M"
        scan = &H1C
        vk = VkKeyScan(Asc(c$)) And &HFF
        GoTo test
    End If
    
    
    VK_Read = VkKeyScan(Asc(c$)) 'And &HFF
    vk = VK_Read And &HFF
    
    Shift = VK_Read And &H100
    'Control = VK_Read And &H200
    'If Control = 1 Then Shift = 0
    'If VK_Read = 13 Then Control = 1 'VK_Read And &H200

    oemchar = "  " ' 2 character buffer
    ' Get the OEM character - preinitialize the buffer
    CharToOem Left$(c$, 1), oemchar
    ' Get the scan code for this key
    scan% = OemKeyScan(Asc(oemchar$)) And &HFF
    ' Send the key down
    
test:
    If Control > 0 Then keybd_event VK_CONTROL, &H38, 0, 0
    If Shift > 0 Then keybd_event VK_SHIFT, &H2A, 0, 0
    keybd_event vk%, scan%, 0, 0
    ' Send the key up
    keybd_event vk%, scan%, KEYEVENTF_KEYUP, 0
    ' Send the key up
    If Shift > 0 Then keybd_event VK_SHIFT, &H2A, KEYEVENTF_KEYUP, 0
     If Control > 0 Then keybd_event VK_CONTROL, &H38, KEYEVENTF_KEYUP, 0
End Sub

Public Sub MySendKeys(ByVal c$)
    Dim x&
    For x& = 1 To Len(c$)
        SendAKey Mid$(c$, x&)
    Next x&
    SendAKey Chr(13)
   ' SendAKey Chr(9)
End Sub
Function iFileChanged(ByVal sFileName As String, ReferenceDateTime As String) As Boolean
Dim i As Integer
    Err.Clear
    On Error Resume Next
    i = Len(Dir$(sFileName))
    If Err.Number <> 0 Or i = 0 Then
        iFileChanged = False
    Else
        If FileDateTime(sFileName) = ReferenceDateTime Then
            iFileChanged = False
        Else
            iFileChanged = True
        End If
    End If
End Function

Public Function AllocateMemory(BufferSize As Long) As Boolean
Dim GlblAddr2 As Long
Dim glblhnd As Long
Dim di As Long
glblhnd = GlobalAlloc(GMEM_ZEROINIT Or GMEM_MOVEABLE, BufferSize)
If glblhnd = 0 Then
    AllocateMemory = False
    Exit Function
Else
    gBuffer.hMem = glblhnd
    gBuffer.Address = GlobalLock(glblhnd)
    gBuffer.StartAddr = gBuffer.Address
    AllocateMemory = True
End If
End Function

Public Function FreeMemory() As Boolean

Dim GlblAddr2 As Long
Dim glblhnd As Long
Dim di As Long

If Not GlobalUnlock(gBuffer.hMem) = 0 Then
    gBuffer.hMem = 0
    gBuffer.Address = 0
    gBuffer.StartAddr = 0
    FreeMemory = False
    Exit Function
End If

If GlobalFree(gBuffer.hMem) = 0 Then
    FreeMemory = True
Else
    FreeMemory = False
End If
gBuffer.Address = 0
gBuffer.StartAddr = 0
gBuffer.hMem = 0
End Function

Public Function GetExt(psFileName As String) As String
    Dim llPos As Long
    Dim lbFound As Boolean
    
    If Len(psFileName) <= 1 Or InStrRev(psFileName, ".") = 0 Then
        GetExt = ""
        Exit Function
    End If
    GetExt = Mid(psFileName, InStrRev(psFileName, ".") + 1, Len(psFileName) - InStrRev(psFileName, ".") + 1)
End Function

Public Property Get TempPathLocation() As String
Dim sBuffer As String
Dim MAX_PATH As Long
MAX_PATH = 512

TempPathLocation = App.Path & "\temp\"
Exit Sub

sBuffer = Space(MAX_PATH)
If GetTempPath(MAX_PATH, sBuffer) <> 0 Then
    TempPathLocation = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
Else
    TempPathLocation = ""
End If
End Property
Public Function GetFileText(FileName As String) As String
Dim FileNum As Integer
Dim Item As String

    If iFileExists(FileName) Then
        FileNum = FreeFile
        Open FileName For Binary As FileNum
        GetFileText = Input(LOF(FileNum), FileNum)
        Close FileNum
    Else
        GetFileText = ""
    End If
End Function

Public Function PutFileText(FileName As String, WriteData As String) As Boolean
Dim FileNum As Integer
Dim Item As String
On Error Resume Next
Kill FileName
On Error GoTo BadPutFIleText
    FileNum = FreeFile
    Open FileName For Output As FileNum
    Print #FileNum, WriteData
    Close FileNum
    PutFileText = True
    Exit Function
BadPutFIleText:
    Err.Clear
    Close FileNum
    Item = Err.Description
    PutFileText = False
End Function

Public Function StripFileName(TheStr As String) As String
    If InStr(1, TheStr, "\") = 0 Then
        StripFileName = TheStr
    Else
        StripFileName = Mid(TheStr, InStrRev(TheStr, "\") + 1, Len(TheStr) - InStrRev(TheStr, "\") + 1)
    End If
End Function
Public Function StripPath(TheStr As String) As String
    If InStr(1, TheStr, "\") = 0 Then
        StripPath = ""
    Else
        StripPath = Mid(TheStr, 1, InStrRev(TheStr, "\"))
    End If
End Function

Public Sub SetMixMasterPath()
Dim SectionName As String

SectionName = "Remailer Info"
gMixPath = ReadProfile(SectionName, "MixmasterPath")
frmMixmasterHome.Show vbModal
If gMixPath = "" Or Not iFileExists(gMixPath & "\mixmaste.exe") Then
    'gMixPath = App.Path
    MsgBox "Mixmaster was not found in " & gMixPath & ".  Private Idaho is looking for the file mixmaste.exe.", vbApplicationModal + vbQuestion, "Mixmaster"
    Exit Sub
Else
    SetMixmasterShortcuts
End If
On Error Resume Next
'If Not iFileExists(gMixPath & "\mixmaster.htm") Then FileCopy App.Path & "\mixmaster.htm", gMixPath & "\mixmaster.htm"
'If Not iFileExists(gMixPath & "\type2.lis") Then FileCopy App.Path & "\type2.lis", gMixPath & "\type2.lis"
'If Not iFileExists(gMixPath & "\pubring.mix") Then FileCopy App.Path & "\pubring.mix", gMixPath & "\pubring.mix"
End Sub
Public Sub SetMixmasterShortcuts()

On Error GoTo BadMixmasterShortcut
'---------------------------------------------------------------
'Need to check first and exit if okay..
'---------------------------------------------------------------
    Dim slnk As cShellLink                              ' ShellLink Variable
'---------------------------------------------------------------
    
    Dim LnkFile As String
    Dim sLnkFile As String
    Dim ExeName As String
    Dim sExeName As String
    Dim sWorkDir As String
    Dim WorkDir As String
    Dim CmdArgs As String
    Dim IconFile As String
    Dim IconIndex As Long
    Dim lIconIndex As Long
    Dim ShowCmd As Long
    Dim lShowCmd As Long
    
    Set slnk = New cShellLink                           ' Create ShellLink Instance
    sLnkFile = gMixPath & "\mixmaste.pif"
    sExeName = gMixPath & "\mixmaste.exe"
    sWorkDir = gMixPath
    If Not iFileExists(sExeName) Then Exit Sub
    lShowCmd = 7
    lIconIndex = 1
   
     slnk.GetShellLinkInfo sLnkFile, _
                          ExeName, _
                          WorkDir, _
                          CmdArgs, _
                          IconFile, _
                          IconIndex, _
                          ShowCmd                       ' Get Info for shortcut file...
                        
  '  If UCase(Mid(ExeName, 1, InStr(1, ExeName, ".exe") + 3)) = UCase(sExeName) And ShowCmd = lShowCmd Then Exit Sub      'if already set then exit
    slnk.CreateShellLink sLnkFile, _
                         sExeName, _
                         sWorkDir, _
                         CmdArgs, _
                         IconFile, _
                         lIconIndex, _
                         lShowCmd             ' Create a ShellLink (ShortCut)
    
    Set slnk = Nothing                                  ' Destroy object reference
    Exit Sub
BadMixmasterShortcut:
    MsgBox "The library ShellLnk.dll failed to load.  To fix this you may need run 'regsvr32 shelllnk.dll' from the run command line.", vbApplicationModal + vbCritical, "Shortcut Failure"
    'iResponse = ShellExecute(Me.hWnd, "open", sPath, vbNullString, CurDir, SW_SHOW)
    Err.Clear
End Sub
Public Sub LoadMailGroupList(ListName As ListBox)
Dim rs As Recordset


On Error GoTo BadList
Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
ListName.Clear
While Not rs.EOF
    ListName.AddItem rs("List Name")
    rs.MoveNext
    DoEvents
Wend
rs.Close

BadList:
Err.Clear

End Sub

Public Sub OpenDatabase()
Dim spass As String


spass = ""
'spass = spass & "21"
On Error GoTo CanOpenDatabase
'Open exclusive
Set DB = DBEngine.OpenDatabase((App.Path _
         & IIf(Right$(App.Path, 1) <> "\", "\", "") & _
         "PI32" & "PostOffice.MDB"), False, False, ";pwd=" & spass)

Exit Sub
CanOpenDatabase:
    MsgBox "Can't open the Private Idaho Database (Error returned was: " & Err.Description & " ) - the application can't continue.  Please report the error to pisupport@itech.net.au", vbApplicationModal + vbCritical, App.Title
    End
End Sub

Public Sub SavePasswordToDatabase(Password As String)
    Dim rs As Recordset
On Error GoTo BadSave
Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
rs.Edit
rs("AccountPassword") = Password
rs.Update
rs.Close
Exit Sub
BadSave:
    MsgBox "An error was detected: " & Err.Description, vbCritical + vbApplicationModal, "Nym Creation Error."
    Err.Clear
End Sub

Public Function GetPasswordFromDatabase() As String

Dim rs As Recordset
On Error GoTo BadSave
Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
'rs.Edit
If Not rs.EOF Then GetPasswordFromDatabase = rs("AccountPassword")
rs.Close
Exit Function
BadSave:
    MsgBox "An error was detected: " & Err.Description, vbCritical + vbApplicationModal, "Nym Creation Error."
    Err.Clear
End Function

Public Function CreateListTable(TableListName As String) As String
Dim rs As Recordset
Dim tNewTable As TableDef

On Error Resume Next
'First try to open master if not there then create it
Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
If Err.Number > 0 Then
    'rs.Close
    Err.Clear
    On Error GoTo BadTableCreation
    Set tNewTable = DB.CreateTableDef("Master List")
    With tNewTable
        .Fields.Append .CreateField("MasterList ID", dbLong)
        .Fields("MasterList ID").Attributes = dbAutoIncrField
        .Fields.Append .CreateField("List Name", dbText, 255)
        .Fields.Append .CreateField("Member Type", dbInteger)
        ' 0 = No members, 1 = Private Address Book, 2 = Public Keyring
       ' .Fields.Append .CreateField("List ID", dbLong)
    End With
    DB.TableDefs.Append tNewTable
End If

On Error GoTo BadTableCreation
Set tNewTable = DB.CreateTableDef(TableListName)
With tNewTable
    .Fields.Append .CreateField("List ID", dbLong)
    .Fields("List ID").Attributes = dbAutoIncrField
    .Fields.Append .CreateField("Contact Email", dbText, 128)
    .Fields.Append .CreateField("Member ID", dbLong)
    '.Fields.Append .CreateField("Member ID", dbLong)
    '.Fields.Append .CreateField("Password", dbText, 128)
    '.Fields.Append .CreateField("Expired", dbBoolean)
End With
DB.TableDefs.Append tNewTable

    
'Now create the table references

Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
rs.AddNew
rs("List Name") = TableListName
rs("Member Type") = 0 ' Initialise to no members
rs.Update
rs.Close
CreateListTable = ""
Exit Function

BadTableCreation:
    CreateListTable = False
    Err.Clear
    CreateListTable = Err.Description
    Exit Function
End Function
Public Function DeleteListTable(TableListName As String) As Boolean
Dim rs As Recordset
Dim tTable As TableDef

On Error GoTo BadTableDeletion
'First try to open master if not there then create it
Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
rs.FindFirst "[List Name] =" & "'" & TableListName & "'"
If Not rs.NoMatch Then
 'Now delete the reference in Master List and the actual table
    rs.Delete
    rs.Close
    'B.Delete (TableListName)
End If

DeleteListTable = True
Exit Function

BadTableDeletion:
    DeleteListTable = False
    Err.Clear
    Err.Description = Err.Description
    Exit Function
End Function
Public Function RenameListTable(TableListName As String) As Boolean
Dim rs As Recordset
Dim tTable As TableDef
Dim sNewName As String

On Error GoTo BadTableRename

sNewName = myInputBox("Enter the name of the new MailGroup", "MailGroup Name")
If sNewName = "" Then
    Beep
    RenameListTable = False
    Exit Function
End If
'First try to open master if not there then create it
Set rs = DB.OpenRecordset("Master List", dbOpenDynaset)
rs.FindFirst "[List Name] =" & "'" & TableListName & "'"
If Not rs.NoMatch Then
 'Now rename the reference in Master List and the actual tabl
    rs.Edit
    rs("List Name") = sNewName
    rs.Update
    rs.Close
    Set tTable = DB.TableDefs(TableListName)
    tTable.name = sNewName
End If

RenameListTable = True
Exit Function

BadTableRename:
    RenameListTable = False
    MsgBox "An error occurred while trying to rename the mailgroup.  Error was: " & Err.Description, vbApplicationModal + vbCritical, "Rename Error"
    Err.Clear
    Exit Function

End Function
Public Function LookUpGroupName(GroupName As String) As Boolean
Dim rs As Recordset

On Error Resume Next
Set rs = DB.OpenRecordset(Trim(GroupName), dbOpenDynaset)
If Err.Number = 0 Then
    LookUpGroupName = True
Else
    LookUpGroupName = False
End If
Err.Clear
'Exit Function
'rs.FindFirst "[List Name] = " & "'" & Trim(GroupName) & "'"

'rs("List Name") = rs("List Name")
'If rs.EOF Then
    'LookUpGroupName = False
'Else
   ' LookUpGroupName = True
'End If
rs.Close

End Function

Public Function GetMembersOfGroup(GroupName As String) As Variant
Dim rsList As Recordset
Dim rsMaster As Recordset
Dim iMemberType As Integer
Dim i As Integer
Dim MemberArray() As Variant

Set rsMaster = DB.OpenRecordset("Master List", dbOpenDynaset)
rsMaster.FindFirst "[List Name] = " & "'" & GroupName & "'"
If rsMaster.NoMatch Then Exit Function
iMemberType = rsMaster("Member Type") ' ie is it Private mail or PGP
rsMaster.Close

Set rsList = DB.OpenRecordset(GroupName, dbOpenDynaset)
i = 0
While Not rsList.EOF
    i = i + 1
    ReDim Preserve MemberArray(i)
    MemberArray(i) = rsList("Contact Email")
    MemberArray(0) = i
    rsList.MoveNext
Wend
'The last entry in the array after the member details holds the type
ReDim Preserve MemberArray(MemberArray(0) + 1)
Select Case iMemberType
    Case 1
        MemberArray(MemberArray(0) + 1) = CONTACT_IN_DB
    Case 2
        MemberArray(MemberArray(0) + 1) = CONTACT_ON_PGPKEYRING
End Select
GetMembersOfGroup = MemberArray
End Function

Public Function myInputBox(Prompt As String, Caption As String) As String
    frmInputBox.lblPrompt = Prompt
    frmInputBox.Caption = Caption
    frmInputBox.Show vbModal
    myInputBox = frmInputBox.txtInputText
    Set frmInputBox = Nothing
End Function
' Convert a date string into a Date value using a format.
' The format should be "ymd", "mdy", etc.
Public Function StringToDate(ByVal date_string As String, ByVal _
    date_format As String) As Date
Dim date_parts() As String
Dim day_part As String
Dim month_part As String
Dim year_part As String

    date_parts = Split(date_string, ",")

    date_format = LCase$(date_format)
    Select Case Mid$(date_format, 1, 1)
        Case "d"
            day_part = date_parts(0)
        Case "m"
            month_part = date_parts(0)
        Case "y"
            year_part = date_parts(0)
    End Select
    Select Case Mid$(date_format, 2, 1)
        Case "d"
            day_part = date_parts(1)
        Case "m"
            month_part = date_parts(1)
        Case "y"
            year_part = date_parts(1)
    End Select
    Select Case Mid$(date_format, 3, 1)
        Case "d"
            day_part = date_parts(2)
        Case "m"
            month_part = date_parts(2)
        Case "y"
            year_part = date_parts(2)
    End Select

    StringToDate = CDate(day_part & " " & MonthName(month_part) & _
        ", " & year_part)
End Function
