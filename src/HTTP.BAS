Attribute VB_Name = "HTTP"
Option Explicit
Function ConvertStr(TheStr As String) As String
    Dim FileOutBuffer As String
    Dim y As String
    Dim x As Long

    On Error GoTo CovertStrErr
    
    For x = 1 To Len(TheStr)
        y = Mid$(TheStr, x, 1)
        If x > 2 Then
            '---------------------------------------------
            'second character and thereafter
            '---------------------------------------------
            If (y = Chr$(10)) And (Mid$(TheStr, x - 1, 1) <> Chr$(13)) Then
                '---------------------------------------------
                'ready to append CR+LF
                '---------------------------------------------
                FileOutBuffer = FileOutBuffer & vbCrLf
            Else
                '---------------------------------------------
                'add another character to buff
                '---------------------------------------------
                FileOutBuffer = FileOutBuffer & y
            End If
        Else
            '---------------------------------------------
            'first character in stream
            '---------------------------------------------
            FileOutBuffer = FileOutBuffer + y
        End If
    Next
    '---------------------------------------------
    'dump the converted stream
    '---------------------------------------------
    ConvertStr = FileOutBuffer
    Exit Function

CovertStrErr:
    MsgBox Err.Description + "ConvertStr"
    ConvertStr = ""
    Exit Function
End Function
Function GetKeyBlock(TheStr As String) As String
    Dim x As String
    Dim y As String
    
    x = InStr(1, TheStr, "-----BEGIN PGP PUBLIC KEY BLOCK-----")
    y = InStr(1, TheStr, "-----END PGP PUBLIC KEY BLOCK-----")
    If (x > 0) And (y > 0) Then
        GetKeyBlock = Mid$(TheStr, x, (y + 34) - x)
    Else
        GetKeyBlock = "No PGP public key found." + gCRLF
        'GetKeyBlock = TheStr
    End If
End Function

Sub GetWebURL(URLName As String)

   ' On Error GoTo GetURLError
    On Error Resume Next
    '---------------------------------------------
    'set the URL of the document to fetch
    '---------------------------------------------
    PIForm(gActivePIInstance).HTTP1.Action = 4 'reset headers
    PIForm(gActivePIInstance).HTTP1.URL = URLName
    '---------------------------------------------
    'fetch the document using the HTTP GET method
    '---------------------------------------------
    
    PIForm(gActivePIInstance).HTTP1.Action = 1
    DoEvents
    
    Exit Sub

GetURLError:
    HideStatus
    Select Case Err.Number
        Case 25066
            '---------------------------------------------
            '[10065] No route to host.
            '---------------------------------------------
'            MsgBox "This is an ISP connection error.  Apparently, you have become disconnected from your Internet Service Provider.  Please connect to your ISP and retry."
        Case 26005
            '---------------------------------------------
            '[11004] Valid name, no data record (check DNS setup)
            '---------------------------------------------
'            #If win16 = true Then
'                MsgBox "FATAL INTERNET CONNECTION ERROR.  You tried to perform an internet operation without connecting to your Internet Service Provider first.  The only way to clear this error is to REBOOT your machine.  Please try again later after you reboot and connect to your ISP first."
'            #Else
'                Resume
'            #End If
        Case Else
            MsgBox Err.Number & " (" & Err.Description & ")" & " in GetWebUrl."
    End Select
    Err.Clear
    gCancelAction = True
End Sub

Sub WriteRemailerInfo()
   
    Dim FileNum As Integer
    On Error GoTo WriteRemailerError
    FileNum = FreeFile
    If iFileExists(App.Path + "\remailer.old") Then
        Kill App.Path + "\remailer.old"
    End If
    If iFileExists(App.Path + "\remailer.htm") Then
        Name App.Path + "\remailer.htm" As App.Path + "\remailer.old"
    End If
    If Not gWebPage = "" Then
        Open App.Path + "\remailer.htm" For Output As FileNum
        Print #FileNum, gWebPage
        Close FileNum
    End If
    Exit Sub

WriteRemailerError:
    MsgBox Err.Description & " :-> WriteRemailerInfo"
    Err.Clear
End Sub
Sub WriteStrToFile(theFile As String, TheStr As String)
       
    'writes a string to a file
    'if the file exists, renames with .old extension
    'if the .old exists, deletes it first
    Dim FileNum As Integer
    Dim FileBuffer
    Dim FileOutBuffer As String
    On Error GoTo WriteStrToFileErr
    FileNum = FreeFile
    If iFileExists(StripExt(theFile) + "old") Then
        Kill StripExt(theFile) + "old"
    End If
    If iFileExists(theFile) Then
        Name theFile As StripExt(theFile) + "old"
    End If
    Open App.Path + "\remzysd.out" For Output As #FileNum
    Print #FileNum, TheStr
    Close FileNum
    Name App.Path + "\remzysd.out" As theFile
    Exit Sub

WriteStrToFileErr:
    MsgBox Err.Description + "WriteStrToFile"
    Exit Sub
End Sub
