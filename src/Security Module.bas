Attribute VB_Name = "ApplicationUtilities"
Option Explicit
Public Function BinHex(BinStr As String)
Dim Result As String, i As Integer
For i = 1 To Len(BinStr)
    Result = Result & Right("00" & Hex(Asc(Mid(BinStr, i, 1))), 2)
Next
BinHex = Result

End Function

Public Function MaxValue(Val1 As Single, Val2 As Single, Val3 As Single) As Single
    'Dim Max As Single
    If Val1 > Val2 Then
        MaxValue = Val1
    Else
        MaxValue = Val2
    End If
    If Val3 = INVALID Then _
        Exit Function
    If Val3 > MaxValue Then _
        MaxValue = Val3
   
      
End Function


Public Function MinValue(Val1 As Single, Val2 As Single, Val3 As Single) As Single
   
    If Val1 < Val2 Then
        MinValue = Val1
    Else
        MinValue = Val2
    End If
    If Val3 = INVALID Then _
        Exit Function
    If MinValue > Val3 Then _
        MinValue = Val3
   
End Function

Public Sub SimpleCrypt(PlainText As String, CipherText As String, KeyValue As String)

Dim i As Long, Prev As Integer, Result As String
Dim Char As Integer, KeyIndex As Integer
Dim KeyLen As Integer, TextValue As String
Dim NewChar As Integer, fEncrypting As Integer
ReDim keychar(255) As Integer

'Magic Values - usded fr en/decryption
Const MAGIC1 = 187
Const MAGIC2 = 9
'Determine if we're encrypting or decrypting
If Len(PlainText) Then
    fEncrypting = True
    TextValue = PlainText
Else
    TextValue = CipherText
End If

'Initialise 'previous character' value index into key string and lenght of key
Prev = MAGIC1: KeyIndex = 1
KeyLen = Len(KeyValue)

'Convert key string to array
For i = 1 To Len(KeyValue)
    keychar(i) = Asc(Mid(KeyValue, i, 1))
Next i

'/Actual en/decryption loop
For i = 1 To Len(TextValue)
    Char = Asc(Mid(TextValue, i, 1))
    NewChar = Char Xor keychar(KeyIndex) Xor Prev Xor _
        ((i / MAGIC2) Mod 255)
    Result = Result & Chr(NewChar)
    If fEncrypting Then
        Prev = Char
    Else
        Prev = NewChar
    End If
    KeyIndex = KeyIndex + 1
    If KeyIndex > KeyLen Then KeyIndex = 1
Next i

'Reture result to caller
If fEncrypting Then
    CipherText = Result
Else
    PlainText = Result
End If


End Sub


Public Function HexBin(HexStr As String)
Dim Result As String, i As Integer
For i = 1 To Len(HexStr) Step 2
    Result = Result & Chr(Val("&H" & Mid(HexStr, i, 2)))
Next
HexBin = Result

End Function

Public Function CheckSerialNumber(SerialNumber As String) As Long
Dim i As Long, J As Integer
Dim SI As Long
Dim RI As Long
Dim iRVCount As Integer
Dim UpperBound As Long
Dim LowerBound As Long
Dim Range As Integer
Dim RangeIncrements As Integer
Dim iVer As Integer
Dim ReleaseIndex As Long
Dim SerialIndex As String
Dim rs As Recordset
'FitTest Constants

On Error GoTo BadCheck
'Initial state
CheckSerialNumber = INVALID
gFullRelease = False 'demo

'Freeware version

Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
            rs.Edit
            rs("Expired") = False
            rs.Update
            rs.Close
            gFullRelease = True
            CheckSerialNumber = 8909512
Exit Function
'This was given out and returned
'If SerialNumber = "13323-6337613" Then End
'If SerialNumber = "13323-6340623" Then End
'I'f SerialNumber = "13323-6374294" Then End
'If SerialNumber = "13323-6397127" Then End
'If SerialNumber = "13323-8894398" Then End
'If SerialNumber = "13323-6604756" Then End
'If SerialNumber = "13323-2516992" Then End
'If SerialNumber = "13323-4278637" Then End
If SerialNumber = "15187-8909512" Then End

'If Not gFullRelease Then
  '  On Error Resume Next
   ' DB.Close
    'WipeFile (App.Path & "\PI32PostOffice.MDB")
 '   End
'End If

i = InStr(1, SerialNumber, "-", 1)
ReleaseIndex = Val(Mid(SerialNumber, 1, 5))
SerialIndex = Mid(SerialNumber, i + 1)
'Version First
LowerBound = 12034 '10101
UpperBound = 99999
Range = 5

'Comment out as new serial numbers are generated
'gFullRelease = 0
i = Rnd(-1)
iRVCount = 0
For i = 0 To Range
        RI = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
        If ReleaseIndex = RI Then
            Exit For
        End If
        iRVCount = iRVCount + 1
Next

LowerBound = 1123131  '1000101
UpperBound = 9999999
Range = 500 * 10

'Comment out as new serial numbers are generated
i = Rnd(-1)
'SerialIndex = ""
For i = 0 To Range
        SI = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
        'SerialIndex = SerialIndex & "15187-" & CStr(SI) & vbCrLf
        'If CStr(SI) = "7593566" Then
       ' SerialIndex = SI
       ' End If
        If SerialIndex = CStr(SI) Then
            'CheckSerialNumber = "6651169"
            CheckSerialNumber = SI
            'on Error Resume Next
            Set rs = DB.OpenRecordset("Users", dbOpenDynaset)
            rs.Edit
            rs("Expired") = False
            rs.Update
            rs.Close
            gFullRelease = True
            Exit Function
        End If
Next
'Call PutFileText(App.Path & "\sn.txt", SerialIndex)
Exit Function
BadCheck:
    'SerialIndex = Err.Description
    CheckSerialNumber = INVALID
    gFullRelease = False
    Exit Function
End Function



Public Function SecurityCheck() As String
SecurityCheck = "15187-9683220"
        

End Function

'**************************************
' Name: Get Version Number for EXE, DLL
'     or OCX files
' Description:This function will retriev
'     e the version number, product name, orig
'     inal program name (like if you right cli
'     ck on the EXE file and select properties
'     , then select Version tab, it shows you
'     all that information) etc
' By: Serge
'
' Returns:FileInfo structure
'
' Assumes:Label (named Label1 and make i
'     t wide enough, also increase the height
'     of the label to have size of the form),
'     Common Dilaog Box (CommonDialog1) and a
'     Command Button (Command1)
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=4976&lngWId=1'for details.'**************************************



Public Function GetFileVersionInformation(ByRef pstrFieName As String, ByRef tFileInfo As FILEINFO) As VerisonReturnValue

    Dim lBufferLen As Long, lDummy As Long
    Dim sBuffer() As Byte
    Dim lVerPointer As Long
    Dim lRet As Long
    Dim Lang_Charset_String As String
    Dim HexNumber As Long
    Dim i As Integer
    Dim strTemp As String
    'Clear the Buffer tFileInfo
    tFileInfo.CompanyName = ""
    tFileInfo.FileDescription = ""
    tFileInfo.FileVersion = ""
    tFileInfo.InternalName = ""
    tFileInfo.LegalCopyright = ""
    tFileInfo.OriginalFileName = ""
    tFileInfo.ProductName = ""
    tFileInfo.ProductVersion = ""
    lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)


    If lBufferLen < 1 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    ReDim sBuffer(lBufferLen)
    lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))


    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)


    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    Dim bytebuffer(255) As Byte
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)
    'Pull it all apart:
    '04------= SUBLANG_ENGLISH_USA
    '--09----= LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows
    '     :Multilingual


    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop

    Dim strVersionInfo(7) As String
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    strVersionInfo(3) = "InternalName"
    strVersionInfo(4) = "LegalCopyright"
    strVersionInfo(5) = "OriginalFileName"
    strVersionInfo(6) = "ProductName"
    strVersionInfo(7) = "ProductVersion"
    Dim buffer As String


    For i = 0 To 7
        buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String _
        & "\" & strVersionInfo(i)
        lRet = VerQueryValue(sBuffer(0), strTemp, _
        lVerPointer, lBufferLen)


        If lRet = 0 Then
            GetFileVersionInformation = eNoVersion
            Exit Function
        End If

        lstrcpy buffer, lVerPointer
        buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)


        Select Case i
            Case 0
            tFileInfo.CompanyName = buffer
            Case 1
            tFileInfo.FileDescription = buffer
            Case 2
            tFileInfo.FileVersion = buffer
            Case 3
            tFileInfo.InternalName = buffer
            Case 4
            tFileInfo.LegalCopyright = buffer
            Case 5
            tFileInfo.OriginalFileName = buffer
            Case 6
            tFileInfo.ProductName = buffer
            Case 7
            tFileInfo.ProductVersion = buffer
        End Select

Next i

GetFileVersionInformation = eOK
End Function

