VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIconImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************************************************************
'*                                                                                 *
'*  Author :        Chuck Spohr and Cosmo Gong                                     *
'*  Company:        Peabody Coal(Cosmo) and Furgurson Consulting(Chuck Spohr)      *
'*  Application:    VBPJ(April, 1998) Sample App                                   *
'*  Date:           2/98                                                           *
'*                                                                                 *
'***********************************************************************************

'Constants
Private Const MAX_PATH = 255

Private Const LVSIL_NORMAL = 0
Private Const LVSIL_SMALL = 1
Private Const LVSIL_STATE = 2
Private Const LVM_FIRST = &H1000
Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)

Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_ICONLOCATION = &H1000
Private Const SHGFI_LINKOVERLAY = &H8000
Private Const SHGFI_OPENICON = &H2
Private Const SHGFI_SELECTED = &H10000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_USEFILEATTRIBUTES = &H10

Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

'Types
Private Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type

'APIs
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" ( _
    ByVal hIml As Long, _
    ByVal i As Integer, _
    ByVal hIcon As Long) As Integer

Private Declare Function SHGetFileInfo Lib "shell32.dll" ( _
    ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As Long) As Long

'member variables
Private mFileTypes() As String
Private mImglstLarge As ImageList
Private mImglstSmall As ImageList
Private mLastImageNum As Integer

'Property Gets and Lets
Public Property Get Icons() As Variant
    Set Icons = mImglstLarge
End Property

Public Property Get SmallIcons() As Variant
    Set SmallIcons = mImglstSmall
End Property

'Functions and Subs
Public Function GetFileIconNum(psPathName As String) As Integer
    Dim i As Integer
    Dim lsExt As String
    Dim lsKey As String
    Dim shFileInfoStruct As SHFILEINFO
    
    If Trim(psPathName & "") <> "" Then
        lsExt = "." & GetExt(psPathName)
        lsKey = "a" & lsExt
        For i = 1 To mImglstLarge.ListImages.Count
            If lsKey = mImglstLarge.ListImages.Item(i).Key Then
                GetFileIconNum = i
                Exit Function
            End If
        Next
    Else
        GetFileIconNum = 1
        Exit Function
    End If
    
    If GetFileIconNum = 0 And lsExt <> "" Then
        'add an icon for it
        Call SHGetFileInfo(lsExt, 0, shFileInfoStruct, LenB(shFileInfoStruct), SHGFI_SYSICONINDEX Or SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES)
        If shFileInfoStruct.hIcon > 0 Then
            mLastImageNum = mLastImageNum + 1
            'imglstLarge.ListImages.Add , lsKey, picGenericIcon.Image
            i = ImageList_ReplaceIcon(mImglstLarge.hImageList, mLastImageNum - 1, shFileInfoStruct.hIcon)
            Call SHGetFileInfo(lsExt, 0, shFileInfoStruct, LenB(shFileInfoStruct), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES)
            'imglstSmall.ListImages.Add , lsKey, picGenericIcon.Image
            Call ImageList_ReplaceIcon(mImglstSmall.hImageList, mLastImageNum - 1, shFileInfoStruct.hIcon)
            GetFileIconNum = mLastImageNum
        End If
    Else
        GetFileIconNum = 1
    End If
End Function

Private Sub picGenericIcon_Click()

End Sub

Private Sub UserControl_Initialize()
    mLastImageNum = 1
End Sub

Private Function GetExt(psFileName As String) As String
    Dim llPos As Long
    Dim lbFound As Boolean
    
    If Len(psFileName) <= 1 Then
        GetExt = ""
        Exit Function
    End If
    
    lbFound = False
    For llPos = Len(psFileName) - 1 To 1 Step -1
        If Mid$(psFileName, llPos, 1) = "." Then
            lbFound = True
            GetExt = Mid$(psFileName, llPos + 1)
            Exit For
        End If
    Next llPos
    
    If Not lbFound Then
        GetExt = ""
    End If
End Function


