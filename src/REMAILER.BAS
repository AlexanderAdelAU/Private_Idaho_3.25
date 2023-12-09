Attribute VB_Name = "REMAILER"
    Option Explicit


Function CheckForStatus(TheLine As String) As String

    '---------------------------------------------
    'check if a line has status information
    '---------------------------------------------
    If Right(TheLine, 1) = "%" Then
        '---------------------------------------------
        'lets double check and see if there's an @
        '---------------------------------------------
        If InStr(1, TheLine, "@") Then
            '---------------------------------------------
            'if so return the string
            '---------------------------------------------
            CheckForStatus = TheLine
            '---------------------------------------------
            'should probably do more validation...
            '---------------------------------------------
        Else
            CheckForStatus = ""
        End If
    Else
        '---------------------------------------------
        'if not, return null
        '---------------------------------------------
        CheckForStatus = ""
    End If
End Function

Function CheckForType(TheLine As String) As String
    '---------------------------------------------
    'see if the line contains remailer type information
    '---------------------------------------------
    If InStr(1, TheLine, "$remailer{") > 0 Then
        '---------------------------------------------
        'if so, return the string
        '---------------------------------------------
        CheckForType = TheLine
    Else
        '---------------------------------------------
        'if not return null
        '---------------------------------------------
        CheckForType = ""
    End If
End Function
Sub ClearRemailerSort()
   
    
    gSortRemailer.Name = ""
    gSortRemailer.ShortName = ""
    gSortRemailer.Address = ""
    gSortRemailer.latency = ""
    gSortRemailer.uptime = ""
    gSortRemailer.history = ""
    gSortRemailer.cpunk = 0
    gSortRemailer.mix = 0
    gSortRemailer.penet = 0
    gSortRemailer.alpha = 0
    gSortRemailer.newnym = 0
    gSortRemailer.Encrypt = 0
    gSortRemailer.post = 0
    gSortRemailer.latent = 0
    gSortRemailer.hash = 0
    gSortRemailer.cut = 0
    gSortRemailer.reserved1 = 0
    gSortRemailer.reserved2 = 0
End Sub

Sub FillNymList()
    
    Dim maxName
    Dim maxUp
    Dim maxLatent As Integer
    Dim tempString As String
    Dim ListHandle As Long
    Dim i As Integer
    Dim tmpstr As String
    
    ReDim lTabPos(2) As Long
    On Error GoTo FillNymListErr
    lTabPos(0) = 10
    lTabPos(1) = 32
    lTabPos(2) = 45
    maxName = 0
    maxUp = 0
    maxLatent = 0
    ListHandle = frmNymServerStats.List1.hWnd
    Call SetTabStops(CLng(ListHandle), 3, lTabPos())
    frmNymServerStats.List1.Clear
    'get the max lengths for padding
    For i = 1 To gTotalMatchedRemailers
        If Len(gMatchedRemailers(i).Name) > maxName Then
            maxName = Len(gMatchedRemailers(i).Name)
        End If
    Next
    'now create the padded string
    For i = 1 To gTotalMatchedRemailers
        If Len(gMatchedRemailers(i).Name) < maxName Then
            tmpstr = gMatchedRemailers(i).Name + Space$((maxName - Len(gMatchedRemailers(i).Name)) + 3)
            If gRemailerType = REMAILER_MIX Then
                tmpstr = tmpstr + Space$(10)
            End If
        Else
            tmpstr = gMatchedRemailers(i).Name + Space$(3)
        End If
        tempString = tmpstr + Chr(9)
        tmpstr = gMatchedRemailers(i).latency
        Select Case Len(tmpstr)
        Case 3
            tmpstr = Space(9) + tmpstr
        Case 4
            tmpstr = Space(6) + tmpstr
        Case 5
            tmpstr = Space(5) + tmpstr
        Case 7
            tmpstr = Space(2) + tmpstr
        End Select
        tempString = tempString + tmpstr + Chr(9)
        tmpstr = gMatchedRemailers(i).uptime
        tmpstr = Space$(7 - Len(tmpstr)) + tmpstr
        tempString = tempString + tmpstr
        frmNymServerStats.List1.AddItem tempString
    Next
    frmNymServerStats.List1.Selected(0) = 1
    Exit Sub

FillNymListErr:
    MsgBox Err.Description & " (FillNymList)", vbApplicationModal, App.Title
    Err.Clear
End Sub

Sub FillRemailerList()
   
    Dim maxName
    Dim maxUp
    Dim maxLatent As Integer
    Dim tmpstr1 As String
    Dim tmpstr2 As String
    Dim iListHandle As Long
    Dim i As Integer
    ReDim lTabPos(3) As Long
    
    On Error GoTo FillRemailerListErr
    'first sort remailers
    'SortCPRemailers
    maxName = 0
    maxUp = 0
    maxLatent = 0
    
    '---------------------------------------------
    'set remailer list tab stop locations
    '---------------------------------------------
    lTabPos(0) = 15
    lTabPos(1) = 30
    lTabPos(2) = 45
    lTabPos(3) = 55
    '---------------------------------------------
    'get hWnd for list window and set the tab stops
    '---------------------------------------------
    iListHandle = frmRemailerList.List3.hWnd
    '---------------------------------------------
    'initialize list
    '---------------------------------------------
    frmRemailerList.List3.Clear
   ' frmPI.List3.AddItem "none"
    '---------------------------------------------
    'set tab stops for the list
    '---------------------------------------------
    Call SetTabStops(CLng(iListHandle), 4, lTabPos())
    If gSortRemailer.penet = 0 And gSortRemailer.eric = 0 Then
        frmRemailerList.List3.AddItem "chain"
    End If
    '---------------------------------------------
    'determine the length of the longest remailer name
    '---------------------------------------------
    For i = 1 To gTotalMatchedRemailers
        If Len(gMatchedRemailers(i).ShortName) > maxName Then
            maxName = Len(gMatchedRemailers(i).ShortName)
        End If
    Next
    '---------------------------------------------
    'pad remailer names so fields line up in list
    '---------------------------------------------
    For i = 1 To gTotalMatchedRemailers
        '---------------------------------------------
        'build up the remailer data string for each remailer
        '---------------------------------------------
        If Len(gMatchedRemailers(i).ShortName) < maxName Then
            '---------------------------------------------
            'build the name field for the shorter names
            '---------------------------------------------
            tmpstr2 = gMatchedRemailers(i).ShortName
            tmpstr2 = tmpstr2 & Space(maxName - Len(gMatchedRemailers(i).ShortName))
        Else
            '---------------------------------------------
            'build the name field for the longest name(s)
            '---------------------------------------------
            tmpstr2 = gMatchedRemailers(i).ShortName
        End If
        '---------------------------------------------
        'insert a tab
        '---------------------------------------------
        tmpstr1 = tmpstr2 & Space(6) & Chr(9)
        '--------------------------------------------
        'fetch histroy
        '--------------------------------------------
        'If gRemailerType = REMAILER_MIX Then
            tmpstr2 = gMatchedRemailers(i).history
            tmpstr1 = tmpstr1 & InsertWith("_", 12 - Len(tmpstr2)) & tmpstr2 & Space(6) & Chr(9)
        'End If
        '---------------------------------------------
        'fetch latency time
        '---------------------------------------------
        tmpstr2 = gMatchedRemailers(i).latency
        '---------------------------------------------
        'pad spaces for different latency time formats
        '---------------------------------------------
        tmpstr1 = tmpstr1 & Space(10 - Len(tmpstr2)) & tmpstr2 & Space(6) & Chr(9)
        '---------------------------------------------
        'fetch uptime percentage
        '---------------------------------------------
        tmpstr2 = gMatchedRemailers(i).uptime
        '---------------------------------------------
        'prepend with appropriate number of spaces
        '---------------------------------------------
        '---------------------------------------------
        'build remailer data string to add to listbox
        '---------------------------------------------
        tmpstr1 = tmpstr1 & Space(10 - Len(tmpstr2)) & tmpstr2 '& Space(5)
        '---------------------------------------------
        'add the string
        '---------------------------------------------
        frmPI.List3.AddItem tmpstr1
        frmRemailerList.List3.AddItem tmpstr1
    Next
    '---------------------------------------------
    'select the remailer list (equivalent to clicking on it)
    '---------------------------------------------
    frmPI.List3.Selected(0) = 1
    frmRemailerList.List3.Selected(0) = 1
    Exit Sub

FillRemailerListErr:
    MsgBox Err.Description & " You need to Fill the remailer list."
    Err.Clear
End Sub

Function GetRemailerIndex(theName As String) As Integer

    Dim Found As Integer
    Dim i As Integer
    
    Found = 0
    
    '---------------------------------------------
    'returns the remailer array index
    '---------------------------------------------
    If gTotalRemailers = 0 Then
        '---------------------------------------------
        'nothing in the array yet, so put the first remailer name in it
        '---------------------------------------------
        InitializeRemailerArray (1)
        gRemailerArray(1).Name = theName
        gTotalRemailers = 1
        GetRemailerIndex = 1
    Else
        '---------------------------------------------
        'the remailers are there
        '---------------------------------------------
        For i = 1 To gTotalRemailers
            If gRemailerArray(i).Name = theName Then
            GetRemailerIndex = i
            Found = 1
            Exit For
        End If
    Next
    '---------------------------------------------
    'couldnt find the name, so add it to the array
    '---------------------------------------------
    If Found = 0 Then
        InitializeRemailerArray (i)
        gRemailerArray(i).Name = theName
        gTotalRemailers = gTotalRemailers + 1
        GetRemailerIndex = i
    End If
End If
End Function
  Sub InitializeMixRemailers(theFile As String)
    
    Dim filenum
    Dim y As Integer
    Dim j As Integer
    Dim i As Integer
    Dim TheLine As String
   'Dim tempstr As String
    Dim tmpstr1 As String
    
    On Error GoTo InitializeMixRemailersErr
    filenum = FreeFile
    j = 0
    Open theFile For Input As filenum
    Line Input #filenum, TheLine
    Close filenum
    tmpstr1 = TheLine
    If Len(tmpstr1) > 1024 Then
        MsgBox ("Converting Mixmaste.htm file.  This may take a minute.")
        tmpstr1 = ConvertStr(tmpstr1)
        WriteStrToFile gMixPath + "\mixmaste.htm", tmpstr1
        TheLine = ""
    End If
    Open theFile For Input As filenum
    While Not EOF(filenum)
        Line Input #filenum, TheLine
        'see if the line contains Mixmaster info
        If TheLine <> "" Then
            If Mid$(TheLine, Len(TheLine), 1) = "%" Then
                j = j + 1
                'get name
                'grab position of first non-space character after first space
                i = InStr(1, TheLine, " ")
                'everything between is the name
                gMatchedRemailers(j).Name = Mid$(TheLine, 1, i - 1)
                
                               
                'get uptime
                'go to the end of the line and parse left until a space
                y = 0
                For y = Len(TheLine) To 1 Step -1
                    If Mid$(TheLine, y, 1) = " " Then
                        Exit For
                    End If
                Next
                gMatchedRemailers(j).uptime = Mid$(TheLine, y + 1, Len(TheLine) - y)
                
                'get latency
                'now parse to the left until we hit a non-space
                For i = y To 1 Step -1
                    If Mid$(TheLine, i, 1) <> " " Then
                        Exit For
                    End If
                Next
                'we have the left position, now we need the right position
                For y = i To 1 Step -1
                    If Mid$(TheLine, y, 1) = " " Then
                        Exit For
                    End If
                Next
                gMatchedRemailers(j).latency = Mid$(TheLine, y + 1, i - y)
                 
                'now get history
                For y = y To 1 Step -1
                    Select Case Mid$(TheLine, y, 1)
                        Case "+", "-", "*", "_", "."
                            Exit For
                    End Select
           
                    'If Not Mid$(TheLine, y, 1) = " " Then
                        'Exit For
                'End If
                Next
                i = y
                For y = y To 1 Step -1
                    If Mid$(TheLine, y, 1) = " " Then
                        Exit For
                    End If
                Next
                If y <= 8 Then
                    gMatchedRemailers(j).history = ""
                Else
                    gMatchedRemailers(j).history = Mid$(TheLine, y + 1, i - y)
                End If
                
                gTotalMatchedRemailers = j
            End If
        End If
    Wend
    Close filenum
    frmPI.Label6.Caption = "Remailer update as of " + Format$(FileDateTime(theFile), "d mmm,yyyy hh:mm")
    Exit Sub

InitializeMixRemailersErr:
    MsgBox Err.Description & ". Can't find the some mixmaster files.  Ensure that the mixmaster.htm has been updated. You need to Initialize Mix Remailers"
    Err.Clear
End Sub
  
    

Sub InitializeRemailerArray(theElement As Integer)
        
    gRemailerArray(theElement).Name = ""
    gRemailerArray(theElement).ShortName = ""
    gRemailerArray(theElement).Address = ""
    gRemailerArray(theElement).latency = ""
    gRemailerArray(theElement).latency = ""
    gRemailerArray(theElement).history = ""
    gRemailerArray(theElement).cpunk = 0
    gRemailerArray(theElement).eric = 0
    gRemailerArray(theElement).mix = 0
    gRemailerArray(theElement).penet = 0
    gRemailerArray(theElement).alpha = 0
    gRemailerArray(theElement).newnym = 0
    gRemailerArray(theElement).Encrypt = 0
    gRemailerArray(theElement).post = 0
    gRemailerArray(theElement).latent = 0
    gRemailerArray(theElement).hash = 0
    gRemailerArray(theElement).cut = 0
    gRemailerArray(theElement).reserved1 = 0
    gRemailerArray(theElement).reserved2 = 0
End Sub

Sub InitializeRemailers(TheFileName As String)
    Dim filenum As Integer
    Dim FileNum2 As Integer
    Dim FileNum3 As Integer
    Dim TheLine As String
    Dim tmpstr2 As String
    Dim tmpstr1 As String
    Dim Item As String
    Dim i As Integer
    Dim CreateTextFile As Boolean
    
    On Error GoTo InitializeRemailerErr
   ' if not ifexist
   ' FileNum = FreeFile
    'Open TheFileName For Input As #FileNum
    'Line Input #FileNum, TheLine
    'Close FileNum
    'tmpstr1 = TheLine
    '---------------------------------------------
    'found an unconverted html file, uncommon
    '---------------------------------------------
   ' If Len(tmpstr1) > 4096 Then
       ' MsgBox ("Converting Remailer.htm file.  This may take a minute.")
       ' tmpstr1 = ConvertStr(tmpstr1)
        'WriteStrToFile App.Path + "\remailer.htm", tmpstr1
        'TheLine = ""
    'End If
    
    'Check for extension Type
    '----------------------------
    'Initialise the remailers to 0 so a refill is forced
    '-----------------------------
    'gTotalRemailers = 0
    CreateTextFile = False
    If Not iFileExists(App.Path + "\remailer.htm") Then
        frmPI.lblStatus = "Remailer data was not updated."
        Beep
        Exit Sub
    End If
    If InStr(1, TheFileName, ".htm", vbTextCompare) Then
        FileNum2 = FreeFile
        Open App.Path + "\remailer.txt" For Output As FileNum2
        FileNum3 = FreeFile
        Open App.Path + "\remailerStats.txt" For Output As FileNum3
        CreateTextFile = True
    End If
    If InStr(1, TheFileName, "private.txt", vbTextCompare) Then
        FileNum2 = FreeFile
        Open App.Path & "\remailer.txt" For Append As FileNum2
        CreateTextFile = True
    End If
    
    '---------------------------------------------
    'usual case, process the converted html file
    '---------------------------------------------
    filenum = FreeFile
    Open TheFileName For Input As #filenum
    While Not EOF(filenum)
        Line Input #filenum, TheLine
        'see if the line contains type info
        tmpstr1 = TheLine
        tmpstr2 = CheckForType(tmpstr1)
        'line contains type info
        'tmpstr2 = ""
        If tmpstr2 <> "" Then
            '---------------------------------------------
            'valid remailer information, so parse the line
            'and update the array
            '---------------------------------------------
            If CreateTextFile Then Print #FileNum2, tmpstr1
            ParseRemailerLine (tmpstr1)
        Else
                
            'line does not contain type info,
            'so see if the line contains status info
            '---------------------------------------------
            tmpstr2 = CheckForStatus(tmpstr1)
            '---------------------------------------------
            'line contains status info
            '---------------------------------------------
            If tmpstr2 <> "" Then
                If CreateTextFile Then Print #FileNum3, tmpstr1
                If CreateTextFile Then Print #FileNum2, tmpstr1
                ParseRemailerLine (tmpstr1)
            End If
        End If
    Wend
    Close filenum
    If CreateTextFile Then
        Close FileNum2
        Close FileNum3
    End If
    
    '---------------------------------------------
    'Put the time and date of last update on form1
    '---------------------------------------------
    frmPI.Label6.Caption = "Remailer update as of " + Format$(FileDateTime(TheFileName), "mm/dd/yy hh:mm")
    'RestoreSettings
    Exit Sub

InitializeRemailerErr:
    Close filenum
    Close FileNum2
    Close FileNum3
    MsgBox Err.Description & " (in procedure InitializeRemailers)"
    Err.Clear
End Sub
Function IsNewNym(theName As String) As Boolean
Dim i As Integer

    
    
    IsNewNym = False
    For i = 1 To gTotalRemailers
        If gRemailerArray(i).Name = theName Then
            If gRemailerArray(i).newnym = 1 Then IsNewNym = True
            Exit For
        End If
Next
End Function

Sub NymAliasConfig()
        
    frmPI.MessageArea.SelStart = 0
    frmPI.MessageArea.SelText = "Config:" & vbLf
    frmPI.MessageArea.SelText = "From: " & gNym & vbLf
    If gNymState = gNYMCONFIG And gNymState <> gNYMRPLYCHANGE Then
        frmPI.MessageArea.SelText = "Nym-Commands: create +acksend +signsend name=" + Chr$(34) + gfullNym + Chr$(34) + vbLf
        frmPI.MessageArea.SelText = "Public-Key:" & vbLf
        InsertKey (NYMKEY)
    Else
        If gNymState <> gNYMRPLYCHANGE Then
            frmPI.MessageArea.SelText = "nym-Commands: " & gNymConfigStr
        End If
    End If
End Sub

Sub ParseRemailerLine(TheLine As String)
Dim tempName As String
Dim RemailerName As String
Dim remailerIndex As Integer
Dim x As Integer
Dim y As Integer
Dim tmpstr As String
'Exit Sub
    On Error GoTo ParseRemailerLineErr
    
    '---------------------------------------------
    'we have valid remailer information, so parse
    'the line and insert into the array
    '---------------------------------------------
    
    '---------------------------------------------
    'is this info or remailer
    '---------------------------------------------
    If Left$(TheLine, 10) = "$remailer{" Then
        '---------------------------------------------
        'its info - so grab the name
        '---------------------------------------------
        x = InStr(1, TheLine, "&lt;") ' This is the raw html and must be here
        y = InStr(1, TheLine, "&gt;") ' This is the raw html and must be here
        If x = 0 Or y = 0 Then
            x = InStr(1, TheLine, "<")
            y = InStr(x, TheLine, ">")
            x = x + 1 ' Jump over >
        Else
            x = x + 4 ' Jump over &lt;
        End If
        If x = 0 Or y = 0 Then Exit Sub
        tempName = Mid(TheLine, x, y - x)
      
      '----This is new to just get the remailer name not the address----'
        x = InStr(1, TheLine, "{""")
        y = InStr(1, TheLine, """}")
        RemailerName = Mid(TheLine, x + 2, y - x - 2)
        '--end new stuff
        
        '---------------------------------------------
        'see if it exists in the global array
        '---------------------------------------------
        remailerIndex = GetRemailerIndex(tempName)
        gRemailerArray(remailerIndex).Name = tempName
        gRemailerArray(remailerIndex).ShortName = RemailerName
        '---------------------------------------------
        'check for cypherpunk type
        '---------------------------------------------
        If InStr(1, TheLine, " cpunk") Then
            gRemailerArray(remailerIndex).cpunk = 1
        End If
        '---------------------------------------------
        'check for eric type
        '---------------------------------------------
        If InStr(1, TheLine, " eric") Then
            gRemailerArray(remailerIndex).eric = 1
        End If
        '---------------------------------------------
        'check for penet type
        '---------------------------------------------
        If InStr(1, TheLine, " penet") Then
            gRemailerArray(remailerIndex).penet = 1
        End If
        '---------------------------------------------
        'check for mix type
        '---------------------------------------------
        If InStr(1, TheLine, " mix") Then
            gRemailerArray(remailerIndex).mix = 1
        End If
        '---------------------------------------------
        'check for alpha type
        '---------------------------------------------
        If InStr(1, TheLine, " alpha") Then
            gRemailerArray(remailerIndex).alpha = 1
        End If
        '---------------------------------------------
        'check for newnym type
        '---------------------------------------------
        If InStr(1, TheLine, " newnym") Then
            gRemailerArray(remailerIndex).newnym = 1
        End If
        '---------------------------------------------
        'check for pgp
        '---------------------------------------------
        If InStr(1, TheLine, " pgp") Then
            gRemailerArray(remailerIndex).Encrypt = 1
        End If
        '---------------------------------------------
        'check for latent
        '---------------------------------------------
        If InStr(1, TheLine, " latent") Then
            gRemailerArray(remailerIndex).latent = 1
        End If
        '---------------------------------------------
        'check for hash
        '---------------------------------------------
        If InStr(1, TheLine, " hash") Then
            gRemailerArray(remailerIndex).hash = 1
        End If
        '---------------------------------------------
        'check for cut
        '---------------------------------------------
        If InStr(1, TheLine, " cut") Then
            gRemailerArray(remailerIndex).cut = 1
        End If
        '---------------------------------------------
        'check for post
        '---------------------------------------------
        If InStr(1, TheLine, " post") Then
            gRemailerArray(remailerIndex).post = 1
        End If
    Else
        '---------------------------------------------
        'This is status information
        'get name
        'grab position of first non-space character after first space
        '---------------------------------------------
        If Right$(TheLine, 1) = "%" Then
         x = InStr(1, TheLine, " ")
         tempName = Mid$(TheLine, 1, x - 1)
        '---------------------------------------------
        'grab position of first space after beginning position
        '---------------------------------------------
        For y = x To (Len(TheLine))
            If Mid$(TheLine, y, 1) <> " " Then
                Exit For
            End If
        Next
        '---------------------------------------------
        'now we need to get the next space
        '---------------------------------------------
        tmpstr = Mid$(TheLine, y, Len(TheLine) - y)
        x = InStr(1, tmpstr, " ")
        tempName = Mid$(tmpstr, 1, x - 1)
       'tmpstr = Mid$(TheLine, y, Len(TheLine) - y)
        'x = InStr(1, tmpstr, " ")
       
        '---------------------------------------------
        'everything between is the name
        '---------------------------------------------
        remailerIndex = GetRemailerIndex(Mid$(tmpstr, 1, x - 1))
        'tempName = Mid$(tmpstr, 1, x - 1)
        '---------------------------------------------
        'now search if it exists in the global array
        '---------------------------------------------
        '--------------------------
        'Now get history
        '----------------------------
        Dim i, j As Integer
        For i = x To Len(tmpstr)
            Select Case Mid$(tmpstr, i, 1)
                    Case "+", "-", "*", "_", ".", "#"
                    Exit For
            End Select
        Next
        For j = i To Len(tmpstr)
            If Mid$(tmpstr, j, 1) = " " Then
                Exit For
            End If
        Next
        
        If j - i = 0 Or j - i > 12 Then
            gRemailerArray(remailerIndex).history = ""
        Else
            gRemailerArray(remailerIndex).history = Mid(tmpstr, i, j - i)
        End If
        
        '---------------------------------------------
        '
        'get uptime
        'go to the end of the line and parse left until a space
        '---------------------------------------------
        y = j
        For y = Len(tmpstr) To 1 Step -1
            If Mid$(tmpstr, y, 1) = " " Then
                Exit For
            End If
        Next
        gRemailerArray(remailerIndex).uptime = Mid$(tmpstr, y + 1, Len(tmpstr) - y)
        '---------------------------------------------
        'get latency
        'now parse to the left until we hit a non-space
        '---------------------------------------------
        For x = y To 1 Step -1
            If Mid$(tmpstr, x, 1) <> " " Then
                Exit For
            End If
        Next
        '---------------------------------------------
        'we have the left position, now we need the right position
        '---------------------------------------------
        For y = x To 1 Step -1
            If Mid$(tmpstr, y, 1) = " " Then
                Exit For
            End If
        Next
        gRemailerArray(remailerIndex).latency = Mid$(tmpstr, y + 1, x - y)
        
    End If
   End If
    
    
    'End If
Exit Sub

ParseRemailerLineErr:
    MsgBox Err.Description & " (ParseRemailerLine)"
    Err.Clear
End Sub



Sub SetTabStops(ListWindowHandle As Long, iNoTabStops As Integer, TabPositions() As Long)
    
       
    '''''''''''''''''''''''''''''''''''''''''''''''''
    '                                               '
    '   Procedure:  SetTabStops                     '
    '                                               '
    '   Author:     Martyn Dowsett (November 1995)  '
    '               Climate Changer Software        '
    '                                               '
    '   Abstract:   This procedure is called to set '
    '               custom tab stops in a Visual    '
    '               Basic ListBox custom control.   '
    '                                               '
    '''''''''''''''''''''''''''''''''''''''''''''''''
    '                                               '
    '   Parameter Definitions                       '
    '                                               '
    '   lListWindowHandle - Listbox hWnd            '
    '                                               '
    '   iNoTabStops       - Number of tabs wanted   '
    '                       in the listbox.         '
    '                                               '
    '   lTabPositions()   - An array of the tab     '
    '                       positions required.     '
    '                                               '
    '''''''''''''''''''''''''''''''''''''''''''''''''
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
        iTabPos(iIndex) = TabPositions(iIndex) * DialogUnitsX * 2
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

Sub SortRemailers()
      
    

    Dim isMix
    Dim isAlpha
    Dim isPGP
    Dim isLatent
    Dim isCut
    Dim isHeader
    Dim isCP
    Dim isPenet
    Dim isSoda
    Dim isPost As Integer
    Dim filenum As Integer
    Dim y As Integer
    Dim i As Integer
    
    y = 0
    isMix = 0
    isAlpha = 0
    isPGP = 0
    isCut = 0
    isHeader = 0
    isLatent = 0
    isCP = 0
    isSoda = 0
    isPenet = 0
    isPost = 0
    'if we're dealing with mix, need to handle differently
    If Not gRemailerType = REMAILER_MIX Then
        For i = 1 To gTotalRemailers
            If Not gRemailerType = REMAILER_MIX Then
                If gRemailerArray(i).cpunk = 1 Then
                    isCP = -1
                Else
                    isCP = 0
                End If
            End If
            If gRemailerType = REMAILER_MIX Then
                If gRemailerArray(i).mix = 1 Then
                    isMix = -1
                End If
            End If
           
            'now compare and filter
            If ((isCP = gRemailerType = REMAILER_CYPHERPUNK)) Then
                If gRemailerArray(i).alpha = 0 Then
                    y = y + 1
                    gMatchedRemailers(y) = gRemailerArray(i)
                    gTotalMatchedRemailers = y
                    isMix = 0
                    isAlpha = 0
                    isPGP = 0
                    isCut = 0
                    isLatent = 0
                    isHeader = 0
                    isCP = 0
                    isSoda = 0
                    isPenet = 0
                    isPost = 0
                    'EXIT FOR
                End If
            End If
        Next
    End If
End Sub

Sub SortRemailersByType()
    Dim x
    Dim y As Integer

    y = 0
    gTotalMatchedRemailers = 0
    '[debug]
    If gSortRemailer.cpunk Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).cpunk = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.mix Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).mix = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.penet Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).penet = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.eric Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).eric = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.alpha Or gSortRemailer.newnym Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).alpha = 1 Or gRemailerArray(x).newnym = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    End If
End Sub

Function InsertWith(FillChar As String, Length As Long) As String
Dim i As Integer
InsertWith = ""
For i = 1 To Length
    InsertWith = InsertWith & FillChar
Next
End Function

Public Function GetRemailer(theName As String) As String

    Dim i As Integer
    GetRemailer = ""
    For i = 1 To gTotalMatchedRemailers
            If gMatchedRemailers(i).ShortName = theName Then
            GetRemailer = gMatchedRemailers(i).Name
            Exit For
        End If
    Next

End Function
