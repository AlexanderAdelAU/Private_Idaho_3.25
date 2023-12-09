Attribute VB_Name = "sgpgUtils"
Option Explicit

'Used in extracting Key Properties
Global Const KEYPROPS_BUFFER_SIZE As Integer = 512
'*****************
' data structure for signature information
Global Signed As Boolean
Type TSig_Data
    Status As String
    UserID As String
    KeyID As String
    DateTimeStr As String
    DateTimeInt As Long
    Checked As Boolean
    Verified As Boolean
    KeyValidity As String
    KeyRevoked As Boolean
    KeyDisabled As Boolean
    KeyExpired As Boolean
    KeyIsOnRing As Boolean
End Type
' 1st data structure for key information
Type TKey_Data
    Private As Boolean
    KeyID As String
    UserID As String
    Bits As String
    DateTimeStr As String
    DateTimeInt As String
    Fingerprint As String
    KeyAlgorithm As String
    Trust As String
    Validity As String
End Type

Global KeyArray() As TKey_Data
Global Key As TKey_Data
Global DefaultKey As TKey_Data

' One cheesey string-parsing function, coming right up!
' (takes as an argument the tab-delimited string produced by
' decode/decodefile functions & parses it to populate a TSig_Data structure)
Public Function ParseSigData(SigProps As String) As TSig_Data
  Dim pos1 As Integer
  Dim pos2 As Integer
  Dim sublen As Integer
  Dim Sig As TSig_Data

  If Trim(SigProps) = Chr(0) Then
    Signed = False
    Exit Function
  End If

  ' Status - the apparent status of the signature:
  ' SIGNED_NOT      -unsigned
  ' SIGNED_GOOD     -signing key found, data intact
  ' SIGNED_BAD      -data not intact
  ' SIGNED_NO_KEY   -signing key not found, data unverified
'  Global Const SIGNED_GOOD = 0
'  Global Const SIGNED_NOT = 1
'  Global Const SIGNED_BAD = 2
'  Global Const SIGNED_NO_KEY = 3

  pos1 = 1
  pos2 = InStr(1, SigProps, Chr(9))
  sublen = pos2 - 1
  Select Case Left(SigProps, 1)
    Case 0
    Sig.Status = "SIGNED_GOOD"
    Case 1
    Sig.Status = "SIGNED_NOT"
    Case 2
    Sig.Status = "SIGNED_BAD"
    Case 3
    Sig.Status = "SIGNED_NO_KEY"
    Case Else
    Sig.Status = "SIGNED_NOT"
  End Select
  
  ' if there appears to be a signature
  If Sig.Status <> "" And Not Sig.Status = "SIGNED_NOT" Then
  
  Signed = True
  
  ' UserID - primary user id of signing key
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.UserID = Trim(Mid(SigProps, pos1, sublen))
  
  ' KeyID - key id of signing key
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.KeyID = Trim(Mid(SigProps, pos1, sublen))
  
  ' DateTimeStr - date & time as a ctime-format string (local time)
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.DateTimeStr = Trim(Mid(SigProps, pos1, sublen))
  
  ' DateTimeInt - date & time as a ctime-style number (GMT/UTC)
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.DateTimeInt = Int(Trim(Mid(SigProps, pos1, sublen)))
  
  ' Checked - is signing key available/is message properly formatted?
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.Checked = Trim(Mid(SigProps, pos1, sublen))
  
  ' Verified - is Checked true and is the data intact?
  ' ( this is the one to check: if this is true, the sig. is good )
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.Verified = Trim(Mid(SigProps, pos1, sublen))
  
  ' KeyValidity - validity level of signing key:
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Select Case Trim(Mid(SigProps, pos1, sublen))
  Case 0
  Sig.KeyValidity = "Unknown"
  Case 1
  Sig.KeyValidity = "Invalid"
  Case 2
  Sig.KeyValidity = "Marginal"
  Case 3
  Sig.KeyValidity = "Complete"
  Case Else
  Sig.KeyValidity = "Uknown"
  End Select

  ' misc. key problems
  
  ' KeyRevoked
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.KeyRevoked = Trim(Mid(SigProps, pos1, sublen))
  
  ' KeyDisabled
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.KeyDisabled = Trim(Mid(SigProps, pos1, sublen))
  
  ' KeyExpired
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, SigProps, Chr(9))
  sublen = pos2 - pos1
  Sig.KeyExpired = Trim(Mid(SigProps, pos1))
  End If
  
  ParseSigData = Sig
End Function

Public Function ParseKeyData(KeyProps As String) As TKey_Data
  Dim i As Integer
  Dim pos1 As Integer
  Dim pos2 As Integer
  Dim sublen As Integer
  'Dim Key As TKey_Data
  Dim s As String, t As String

  On Error GoTo ParseError
  
  If Len(KeyProps) < 5 Then Exit Function
  
  ' Private?
  pos1 = 1
  pos2 = InStr(1, KeyProps, Chr(9))
  sublen = pos2 - 1
  'If Trim(InStr(pos1, KeyProps, "1")) <> 0 Then
  If Left(KeyProps, 1) = "1" Then
    Key.Private = True
  Else
    Key.Private = False
  End If
  
  ' Key ID
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Key.KeyID = Trim(Mid(KeyProps, pos1, sublen))
    
  ' UserID - primary user id of key
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Key.UserID = Trim(Mid(KeyProps, pos1, sublen))
  
  ' Key length
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Key.Bits = Trim(Mid(KeyProps, pos1, sublen))

  ' DateTimeStr - date & time as a ctime-format string (local time)
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Key.DateTimeStr = Trim(Mid(KeyProps, pos1, sublen))
  
  ' DateTimeInt - date & time as a ctime-style number (GMT/UTC)
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Key.DateTimeInt = Trim(Mid(KeyProps, pos1, sublen))
  
  ' Fingerprint
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Key.Fingerprint = Trim(Mid(KeyProps, pos1, sublen))
  
  ' Key Trust
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Select Case Trim(Mid(KeyProps, pos1, sublen))
    Case "0"
    Key.Trust = "Undefined"
    Case "1"
    Key.Trust = "Unknown"
    Case "2"
    Key.Trust = "Never"
    Case "5"
    Key.Trust = "Marginal"
    Case "6"
    Key.Trust = "Complete"
    Case "7"
    Key.Trust = "Ultimate"
    Case Else
    Key.Trust = "Unknown"
    End Select
  
  ' Key Validity
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, KeyProps, Chr(9))
  sublen = pos2 - pos1
  Select Case Trim(Mid(KeyProps, pos1, sublen))
    Case "0"
    Key.Validity = "Unknown"
    Case "1"
    Key.Validity = "Invalid"
    Case "2"
    Key.Validity = "Marginal"
    Case "3"
    Key.Validity = "Complete"
    Case Else
    Key.Validity = "Unknown"
    End Select
      
  ' Public-key algorithm
  pos1 = pos2 + 1
  pos2 = pos2 + 2 'InStr(pos2 + 1, KeyProps, Chr(13))
  sublen = pos2 - pos1
  Select Case Trim(Mid(KeyProps, pos1, sublen))
    Case "0"
    Key.KeyAlgorithm = "Invalid"
    Case "1"
    Key.KeyAlgorithm = "RSA"
    Case "2"
    Key.KeyAlgorithm = "RSAEncryptOnly"
    Case "3"
    Key.KeyAlgorithm = "RSASignOnly"
    Case "4"
    Key.KeyAlgorithm = "ElGamal"
    Case "5"
    Key.KeyAlgorithm = "DSA"
    Case Else
    Key.KeyAlgorithm = "Invalid"
  End Select
     
  ParseKeyData = Key
  '[AC]ChopKeyProps KeyProps, 2
    Exit Function
ParseError:
    Err.Clear
    Beep
    MsgBox "Error in parsing Key Data: "
End Function
' whack the keyring string into smaller strings,
' push data into global array for later use
Public Sub ChopKeyProps(KeyProps As String, Count As Long)
  Dim pos1 As Integer
  Dim pos2 As Integer
  Dim sublen As Integer
  Dim TmpArray() As String
  Dim i As Long
  Dim J As Long
  On Error Resume Next
  ReDim KeyArray(Count) As TKey_Data
  
' KeyProps strings are tab-delimited and end in CRLF
pos1 = 0
pos2 = 0
  For i = 0 To Count - 1
  ' first, split the string into lines ending in CRLF
    pos1 = pos2
    pos2 = InStr(pos2 + 1, KeyProps, Chr(13) & Chr(10)) + 2 'extra 2 for crlf
    If pos2 = 0 Then pos2 = Len(KeyProps)
    sublen = pos2 - pos1 - 2 'Remove the CRLF
    ReDim Preserve TmpArray(i)
    TmpArray(i) = Trim(Mid(KeyProps, pos1, sublen))
  Next i

  ' now parse the strings
  For i = 0 To Count - 1
  ' Private?
  pos1 = 1
  pos2 = InStr(1, TmpArray(i), Chr(9))
  sublen = pos2 - 1
  'If Trim(Mid(KeyProps, pos1, sublen)) = "Key_Public" Then
'  If Trim(InStr(pos1, TmpArray(i), "Key_Public")) <> 0 Then
  'If Left(KeyProps, 1) = "1" Then
  If Left(TmpArray(i), 1) = "1" Then
    KeyArray(i).Private = True
  Else
    KeyArray(i).Private = False
  End If
  
' Key ID
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, TmpArray(i), Chr(9))
  sublen = pos2 - pos1
  KeyArray(i).KeyID = Trim(Mid(TmpArray(i), pos1, sublen))
    
  ' UserID - primary user id of key
  pos1 = pos2 + 1
  pos2 = InStr(pos2 + 1, TmpArray(i), Chr(9))
  'sublen = pos2 - pos1
  sublen = Len(TmpArray(i))
  KeyArray(i).UserID = Trim(Mid(TmpArray(i), pos1, sublen))
    
 Next i
  
End Sub

Public Function CountCRLF(KeyProps As String) As Long
  Dim i As Long
  Dim Count As Long
  Dim v As String
  
  Count = 0
  For i = 1 To Len(KeyProps)
    v = Mid(KeyProps, i, 1)
    If StrComp(v, Chr(13), 0) = 0 Then Count = Count + 1
  Next i
  CountCRLF = Count
End Function

' NOTES
' When calling a Delphi/ObjectPascal DLL from VB, strings must be declared
' "ByVal" and must be of fixed length. In my experience, they must also be
' manually null-terminated, e.g., MyString = 'privacy' & Chr(0), though only
' one of my sources confirms this. At any rate, this project won't execute
' properly on my computer without the final null (in fact it GPFs violently
' without it).


Public Function GetKey(KeyID As String, SelectPrivateKey As Boolean) As String
Dim i As Long
Dim BufferIn As String
Dim KeyBuffer As String
Dim KeyProperties As String

On Error Resume Next
'Find out if the keys are on file
'Get the properties so we can work out the buffer size
KeyProperties = String(KEYPROPS_BUFFER_SIZE, Chr(0))

' first 10 characters will be the key id (0x12345678)
BufferIn = KeyID & Chr(0)

' keyprops takes either key id(s) or user id(s)
' and returns the key's properties
i = spgpKeyProps(BufferIn, KeyProperties, Len(KeyProperties))

' parse the returned property-string into a TKey_Data record
Key = ParseKeyData(KeyProperties)
KeyBuffer = String(Mid(Key.Bits, 1, 4), Chr(0))

i = spgpKeyExport(BufferIn, KeyBuffer, Len(KeyBuffer), IIf(SelectPrivateKey = True, 1, 0), 0)

GetKey = KeyBuffer
End Function
Public Sub GetKeyData(SourceType As Long, Source As String, KeyArray() As TKey_Data, Status As Long)
Dim NumID As Integer
Dim lIndex As Integer
Dim i As Integer
Dim J As Integer
Dim KeyID As String
Dim BufferOut As String
Dim KeyProperties As String
Dim lResponse As Long
Dim spgperr As String
On Error Resume Next
'Find out if the keys are on file
'Get the properties so we can work out the buffer size
'KeyProperties = String(2048, Chr(0))
BufferOut = String(4096, Chr(0))

If SourceType = 0 Then
    lResponse = spgpAnalyzeEx(Source, BufferOut, Len(BufferOut))
Else
    lResponse = spgpAnalyzeFileEx(Source, BufferOut, Len(BufferOut))
End If
Select Case lResponse
    Case 0
        Status = 0
        i = InStr(BufferOut, "Keys_Known: ")
        J = InStr(i, BufferOut, vbCrLf)
        If i > 0 Then
            
            NumID = CInt(Mid(BufferOut, i + Len("Keys_Known: "), J - i - Len("Keys_Known: ")))
            lIndex = 0
            'THis will clear the array and size it to numid
            ReDim KeyArray(NumID)
            For i = 1 To NumID
                If NumID >= 1 Then
                    lIndex = InStr(lIndex + 2, BufferOut, "0x") ' the 2 jumps over "0x"
                    KeyProperties = String(1024, Chr(0))
                    lResponse = spgpKeyProps(Mid(BufferOut, lIndex, 10), KeyProperties, Len(KeyProperties))
                    KeyArray(i) = ParseKeyData(KeyProperties)
                End If
            Next
        End If
    Case 1 'Signed
        Status = lResponse
    Case 2, 3, 4, 5, 6
        Status = lResponse
    Case Else
        Call spgpGetErrorString(i, spgperr)
        Err.Raise 1003, " - (Error in Analysing the message): ", spgperr
    End Select
End Sub

Public Function PGP_SDKPresent() As Boolean
Dim Windir As String * 256
Dim dirLen As Long
Dim strFile As String
  Dim udtFileInfo As FILEINFO
    On Error Resume Next
    '-----
    ' Try new function
    '---------
    
    'If spgpSdkApiVersion >= &H1000000 Then
     '   PGP_SDKPresent = True
    'Else
    '    PGP_SDKPresent = False
    'End If
    'Exit Function
    
    PGP_SDKPresent = False
    dirLen = GetSystemDirectory(Windir, 255)
    
    If iFileExists(Mid(Windir, 1, dirLen) & "\PGP_SDK.DLL") Then
        strFile = Mid(Windir, 1, dirLen) & "\PGP_SDK.DLL" 'version 1.7.8
        If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
            PGP_SDKPresent = False
        Else
            'MsgBox ("version: " & udtFileInfo.FileVersion)
            'If udtFileInfo.FileVersion = "3.5.2" And iFileExists(App.Path & "\spgp.dll") Then
              If iFileExists(App.Path & "\spgp.dll") Then
                PGP_SDKPresent = True
            End If
        End If
    End If

  '   If iFileExists(Mid(Windir, 1, dirLen) & "\PGPSDK.DLL") Then
  '      strFile = Mid(Windir, 1, dirLen) & "\PGPSDK.DLL"
   '     If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
   '         PGP_SDKPresent = False
   '     Else
    '        MsgBox ("version: " & udtFileInfo.FileVersion)
    '        If udtFileInfo.FileVersion = "3.5.2" And iFileExists(App.Path & "\PGPDLL.dll") Then
    '            PGP_SDKPresent = True
    '        End If
   '     End If
    
  '  End If
   
    
        'Label3 = Label3 & "File Version:" & udtFileInfo.FileVersion & vbCrLf
   ' if Label3 = "Company Name: " & udtFileInfo.CompanyName & vbCrLf
   ' Label3 = Label3 & "File Description:" & udtFileInfo.FileDescription & vbCrLf
    'Label3 = Label3 & "File Version:" & udtFileInfo.FileVersion & vbCrLf
   ' Label3 = Label3 & "Internal Name: " & udtFileInfo.InternalName & vbCrLf
    'Label3 = Label3 & "Legal Copyright: " & udtFileInfo.LegalCopyright & vbCrLf
   ' Label3 = Label3 & "Original FileName:" & udtFileInfo.OriginalFileName & vbCrLf
   ' Label3 = Label3 & "Product Name:" & udtFileInfo.ProductName & vbCrLf
    'Label3 = Label3 & "Product Version: " & udtFileInfo.ProductVersion & vbCrLf


   ' If Not iFileExists(Mid(Windir, 1, dirLen) & "\spgp.dll") Then

    '    PGP_SDKPresent = False
   ' End If

End Function


Public Function StripNulls(buffer As String) As String
Dim i As Long
i = InStr(buffer, Chr(0))
If i = 0 Then
    StripNulls = buffer
Else
    StripNulls = Mid(buffer, 1, i - 1)
End If

End Function
Public Function sPGPInsertDetachedSignature(sFileName As String) As String
Dim lResponse As Long
Dim pFileIn As String
Dim pSigFile As String
Dim pSignKeyID As String
Dim pSignKeyPass As String
Dim pComment As String
Dim pSignAlg As Long
Dim pArmor As Long
Dim bRes As Boolean
Dim sResponse As String
Dim SigProps As String * 1024
Dim TheFileName As String
    
    On Error GoTo FileEnSError
    gCancelAction = False
    
    'First set the sign parameters as this is common for both routines
    vb2spgpContext.Initialise
    vb2spgpContext.KeyEncrypt = 1
    vb2spgpContext.Armor = 1
  
    '----------------------------------
    ' Now check if we need to sign or ask for a key
    '-----------------------------------
     'If SignMessage Or ClearSign Then
        gPGPKeyID = ReadProfile("PGP Options", "Default Key ID")
        If gPGPKeyID = "" Then
            vb2spgpContext.SelectPrivateKeys = True
            frmViewKeyRing.lblContext = "You need to sign this message.  Please select a key from private key ring."
            frmViewKeyRing.Caption = "Select Key to sign the message"
            frmViewKeyRing.Show vbModal
        End If
            
        If gPGPKeyID = "" Then
            Beep
            Exit Function
        Else
            vb2spgpContext.SignKeyID = gPGPKeyID
        End If
    
    pFileIn = sFileName
    pSigFile = GetTemporaryFile()
    pSignKeyID = vb2spgpContext.SignKeyID
    pArmor = vb2spgpContext.Armor
    sResponse = myInputBox("In order to sign the message you need to enter your passphrase. ", "Passphrase required for " & vb2spgpContext.SignKeyID)
    If sResponse = "" Then
        gCancelAction = True
        Exit Function
    End If
        vb2spgpContext.SignKeyPass = sResponse
        pSignKeyPass = vb2spgpContext.SignKeyPass
   
    pComment = "Signed by " & pSignKeyID
    lResponse = spgpDetachedSigCreate(pFileIn, pSigFile, pSignKeyID, pSignKeyPass, pComment, pSignAlg, pArmor)
    
    lResponse = spgpDetachedSigVerify(pSigFile, pFileIn, SigProps)

    
    sPGPInsertDetachedSignature = GetFileText(pSigFile)
    'PGPInSertDetachedSignature = True
    KillTemporaryFiles
Exit Function
FileEnSError:
        gCancelAction = False
        sPGPInsertDetachedSignature = ""
        MsgBox "There was an error.  The reason given by the operating system is: " & Err.Description, vbApplicationModal, App.Title
        Err.Clear
    
    
End Function
Public Function spgpAnalyseMessage(inBuffer As String) As Long
Dim BufferIn As String
Dim BufferOut As String
'Dim i As Long
' Private Const PGPAnalyze_Encrypted = 0            ' Encrypted message
'''[AC]
'Private Const PGPAnalyze_Signed = 1               ' Signed message
'Private Const PGPAnalyze_DetachedSignature = 2    ' Detached signature
'Private Const PGPAnalyze_Key = 3                  ' Key data
'Private Const PGPAnalyze_Unknown = 4              ' Non-pgp message

' these are for spgpAnalyzeEx only
'Private Const PGPAnalyze_EncryptedConventional = 5 ' Like it says
'Private Const PGPAnalyze_EncryptedNoKeys = 6



    BufferIn = inBuffer & Chr(0)
    'BufferIn = Space(1024) & Chr(0)
    'BufferIn = inBuffer & Chr(0)
   ' BufferIn = "2134444444444444444444444444444444" & Chr(0)
    'BufferOut = Space(1024) 'String(2048, Chr(0))
 'End If
  spgpAnalyseMessage = spgpAnalyze(BufferIn) 'spgpAnalyzeEx(BufferIn, BufferOut, Len(BufferOut))
'spgpAnalyseMessage = spgpAnalyzeEx(BufferIn, BufferOut, Len(BufferOut))

 ' Select Case i
   ' Case PGPAnalyze_Encrypted ' Encrypted message
    '    spgpAnalyseMessage = "Encrypted"
    'Case PGPAnalyze_Signed ' Signed message
    '    spgpAnalyseMessage = "Signed"
    'Case PGPAnalyze_DetachedSignature ' Detached signature
     '   spgpAnalyseMessage = "Detached Signature"
    'Case PGPAnalyze_Key ' Key data
     '   spgpAnalyseMessage = "Key"
    'Case PGPAnalyze_Unknown ' Key data
    '    spgpAnalyseMessage = "Unkown"
    'Case PGPAnalyze_EncryptedConventional
     '   spgpAnalyseMessage = "Conventional Encryption"
    'Case PGPAnalyze_EncryptedNoKeys ' Key data
     '   spgpAnalyseMessage = "Encrypted No Key"
  'End Select
  
End Function
