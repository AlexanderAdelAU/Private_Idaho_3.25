VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objFileSystem As New FileSystemObject
Private objFilesFound As New Collection

Public Function FindFile(ByVal strFileName As String, ByVal strStartPath As String) As Collection

Dim objFolder As Folder
Dim objSubFolder As Folder
Dim objFile As File
Dim objItem As ListItem

' Check for the existence of the starting path
If Not objFileSystem.FolderExists(strStartPath) Then
    Err.Raise vbObjectError + 6004, "FindFile", "Start path does not exist"
    Set FindFile = Nothing
    Exit Function
End If

' Get a reference to the starting folder
Set objFolder = objFileSystem.GetFolder(strStartPath)

' Search the folders recursively
SearchFolder objFolder, strFileName

' Return the collection of file objects
Set FindFile = objFilesFound

End Function
Private Sub SearchFolder(ByVal objFolder As Folder, _
    ByVal strFileName As String)

Dim objSubFolder As Folder
Dim objFile As File

On Error GoTo SearchFolder

' Traverse the subfolder in the given folder
For Each objSubFolder In objFolder.SubFolders
    ' Call this procedure recursively
    SearchFolder objSubFolder, strFileName
Next

' Traverse the files in the folder and look for matches
For Each objFile In objFolder.Files
    If UCase(Mid(objFile.Name, 1, Len(strFileName))) = UCase(strFileName) Then
        ' Match! Add it to the collection
        objFilesFound.Add objFile
    End If
Next
Exit Sub
SearchFolder:
    ' Raise an error back to the calling program
    Err.Raise vbObjectError + 6005, "SearchFolder", Err.Description
    
End Sub


