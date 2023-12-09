Attribute VB_Name = "NEWSAG1"

#If Win32 Then
Declare Function GetPrivateProfileString% Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Declare Function WritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String)
#Else
Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
Declare Function WritePrivateProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$, ByVal lpFileName$)
#End If

Function GetProfileString$(Heading$, Key$)
   
Dim ReturnString$: ReturnString$ = Space$(255)
Dim Size%: Size% = Len(ReturnString$)

Dim FileName$: FileName$ = App.Path & "\newsag.ini"

Dim x%: x% = GetPrivateProfileString(Heading$, Key$, "", ReturnString$, Size%, FileName$)

'trim it
ReturnString$ = Trim$(ReturnString$)
'remove last null character
ReturnString$ = Left$(ReturnString$, Len(ReturnString$) - 1)
'finally, return it!!
GetProfileString$ = ReturnString$

End Function

Sub SetProfileString(Heading$, Key$, Value$)

Dim FileName$: FileName$ = App.Path & "\newsag.ini"

Dim x%: x% = WritePrivateProfileString%(Heading$, Key$, Value$, FileName$)

End Sub

