VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Name:   CApplication
' Author:       Deborah Kurata, InStep Technologies
' Date:         June 1, 1995
' Description:  Provides additional application
'               properties and methods not
'               provided by the App object.
'
' Revisions:
'
Option Explicit

' PUBLIC: *****************************************
' Public data members

' PRIVATE: *****************************************
' Private data members

' Wrapper property for the copyright of the application
Public Property Get Copyright() As String
    Copyright = App.LegalCopyright
End Property

' Wrapper property for help file
Public Property Get HelpFile() As String
    
        HelpFile = Path & App.HelpFile
   
End Property

' Wrapper property for the Path
' Ensures the path has an appended "\"
Public Property Get Path() As String
    If Right(App.Path, 1) = "\" Then
        Path = App.Path
    Else
        Path = App.Path & "\"
    End If
End Property

' Wrapper property for the title of the application
Public Property Get Title() As String
    Title = App.Title
End Property

' Wrapper property for the version of the application
Public Property Get Version() As String
    Version = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Property


' Standard error routine
' Used for processing all system type errors.
' Any error that requires a log entry should call this routine.
' Parameters:
'  sUserMsg       data to be displayed to the user
'  sInternalMsg   internal message string
'  iInternalCode  error code to be displayed internally, 0 if none
Public Sub ErrorDisplay(sUserMsg As String, sInternalMsg As String, iInternalCode As Integer)
Dim iFileNum
Dim sPrintMsg

   ' Output detailed internal string to a file
   iFileNum = FreeFile
   Open "ERROR.LOG" For Append As iFileNum

   ' Print the date and message displayed to the user.
   sPrintMsg = Now & vbCrLf & sUserMsg
   sPrintMsg = sPrintMsg & vbCrLf & "  Error: [" & iInternalCode & "]: " & sInternalMsg & vbCrLf
   Print #iFileNum, sPrintMsg
   Close #iFileNum

   ' Output detailed string to the debug window
   Debug.Print sPrintMsg

    ' Display user string in a message box
    MsgBox sUserMsg, vbOKOnly + vbExclamation, Me.Title

End Sub
