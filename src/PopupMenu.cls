VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************
'*                                                      *
'*  Author :        Cosmo Gong                          *
'*  Company:        Peabody Coal                        *
'*  Application:    Security System                     *
'*  Module:         Security Manager                    *
'*  Written For:    Peabody Coal                        *
'*  Date:           1/97                                *
'*                                                      *
'********************************************************

' This class requires several items to function:
'    Declare
'       Global giPopUpIndex as Integer in a .BAS module

'    Create a menu structure
'       Generic - mnuGeneric
'           MenuLine - mnuMenuLine(1)
'
'    *NOTE - the menu structure should not be on the active form (causes problems)
'     Suggestion - create an invisible frmMenu with the generic menus
'
'    Place the following code in the mnuMenuLine_Click(Index) event
'       giPopupIndex = Index
'

Private Const mcModuleName = "ClsPopupMenu"

Private miMenuLineCount As Integer
Private mvMenuItems() As Variant
Private mMenuForm As Form
Private miBoldCommand As Integer
Private miSelectedIndex As Integer

Public Property Let SelectedIndex(i As Integer)
    miSelectedIndex = i
End Property
Public Property Let BoldCommand(piNewValue As Integer)
'The Index of the currently bolded menu item
    On Error Resume Next
    miBoldCommand = piNewValue
End Property

Public Property Get BoldCommand() As Integer
'The Index of the currently bolded menu item
    On Error Resume Next
    BoldCommand = miBoldCommand
End Property

Public Property Set MenuForm(pObjNewValue As Object)
'The form that contains mnuGeneric and mnuMenuLine
    On Error Resume Next
    Set mMenuForm = pObjNewValue
End Property

Public Property Get MenuForm() As Object
'The form that contains mnuGeneric and mnuMenuLine
    On Error Resume Next
    Set MenuForm = mMenuForm
End Property

Public Property Let MenuLineCount(piNewValue As Integer)
'The number of items in the current menu
    On Error Resume Next
    miMenuLineCount = piNewValue
End Property

Public Property Get MenuLineCount() As Integer
    On Error Resume Next
    MenuLineCount = miMenuLineCount
End Property

Public Sub AddMenuLine(psCaption As String, Optional pbEnabled, Optional pbChecked, Optional pbVisible)
'Adds a menu item for display in the next available slot (top to bottom), if psCaption is a dash then
'this item is a menu seperator
    On Error GoTo ErrorHandler
    Dim llErr As Long
    Dim lsErrDesc As String
    
    If IsMissing(pbEnabled) Then pbEnabled = True
    If IsMissing(pbChecked) Then pbChecked = False
    If IsMissing(pbVisible) Then pbVisible = True
    miMenuLineCount = miMenuLineCount + 1
    ReDim Preserve mvMenuItems(4, miMenuLineCount)
    
    mvMenuItems(1, miMenuLineCount) = psCaption
    mvMenuItems(2, miMenuLineCount) = pbEnabled
    mvMenuItems(3, miMenuLineCount) = pbChecked
    mvMenuItems(4, miMenuLineCount) = pbVisible
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Sub
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Sub

Public Sub BoldMenuLine(psCaption As String)
' sets psCaption to be the bolded menu line
    Dim llErr As Long
    Dim liTemp As Integer
    Dim lsErrDesc As String
    
    On Error GoTo ErrorHandler
    
    For liTemp = 1 To miMenuLineCount
        If mvMenuItems(1, liTemp) = psCaption Then
           miBoldCommand = liTemp
           Exit Sub
        End If
    Next
    
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Sub
    
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Sub

Public Sub Initialize(frm As Form)
    Set MenuForm = frm
    Set MenuForm.CurrentPopUpMenu = Me
End Sub

Public Function Show(Optional piBoldCommand As Variant) As String
'Displays the current menu
    On Error GoTo ErrorHandler
    Show = -1
    Dim i As Integer
    Dim liTemp As Integer
    Dim lsErrDesc As String
    Dim llErr As Long
    
    If Not IsMissing(piBoldCommand) Then BoldCommand = piBoldCommand
    
    For i = 2 To miMenuLineCount
        Load MenuForm.mnuMenuLine(i)
    Next
    For i = 1 To miMenuLineCount
        With MenuForm.mnuMenuLine(i)
            .Enabled = True
            .Caption = mvMenuItems(1, i)
            .Checked = mvMenuItems(3, i)
            .Visible = mvMenuItems(4, i)
            .Enabled = mvMenuItems(2, i)
        End With
    Next

    miSelectedIndex = -1
    If BoldCommand <> 0 Then
        If mMenuForm.mnuMenuLine(BoldCommand).Enabled And mMenuForm.mnuMenuLine(BoldCommand).Visible Then
            mMenuForm.PopupMenu mMenuForm.mnuGeneric, , , , mMenuForm.mnuMenuLine(BoldCommand)
        Else
            mMenuForm.PopupMenu mMenuForm.mnuGeneric
        End If
    Else
        mMenuForm.PopupMenu mMenuForm.mnuGeneric
    End If
    If miSelectedIndex <> -1 Then Show = mMenuForm.mnuMenuLine(miSelectedIndex).Caption
    MenuForm.mnuMenuLine(1).Visible = True
    For i = 2 To miMenuLineCount
        Unload MenuForm.mnuMenuLine(i)
    Next
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Function
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Function

Public Sub CheckMenuLine(psCaption As String, Optional pbCheck As Variant)
'Sets the checked status of a menu item to the value of pbCheck,
'if pbCheck is omitted the default value is true (checked)
    On Error GoTo ErrorHandler
    Dim llErr As Long
    Dim liTemp As Integer
    Dim lsErrDesc As String
    
    If IsMissing(pbCheck) Then pbCheck = True
    For liTemp = 1 To miMenuLineCount
        If mvMenuItems(1, liTemp) = psCaption Then
           mvMenuItems(3, liTemp) = CBool(pbCheck)
        End If
    Next
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Sub
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Sub

Public Sub EnableMenuLine(psCaption As String, Optional pbEnable As Variant)
'Sets the enabled property of the menu item with caption = psCaption to pbEnable,
'if omitted then pbEnable = True (enabled)
    On Error GoTo ErrorHandler
    Dim llErr As Long
    Dim lsErrDesc As String
    Dim liTemp As Integer
    
    If IsMissing(pbEnable) Then pbEnable = True
    For liTemp = 1 To miMenuLineCount
        If mvMenuItems(1, liTemp) = psCaption Then
           mvMenuItems(2, liTemp) = CBool(pbEnable)
           Exit Sub
        End If
    Next
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Sub
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Sub

Public Sub ShowMenuLine(psCaption As String, Optional pbShow As Variant)
'Sets the visible property of the menu item with caption = psCaption to pbShow,
'if omitted then pbShow = True (visible)
    On Error GoTo ErrorHandler
    Dim llErr As Long
    Dim lsErrDesc As String
    Dim liTemp As Integer
    
    If IsMissing(pbShow) Then pbShow = True
    For liTemp = 1 To miMenuLineCount
        If mvMenuItems(1, liTemp) = psCaption Then
            mvMenuItems(4, liTemp) = CBool(pbShow)
            Exit Sub
        End If
    Next
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Sub
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Sub

Public Sub ChangeMenuCaption(psOldCaption As String, psNewCaption As String)
'Changes a menu caption from psOldCaption to psNewCaption
    On Error GoTo ErrorHandler
    Dim llErr As Long
    Dim lsErrDesc As String
    
    Dim liTemp As Integer
    
    If Trim(psOldCaption & "") <> "" And Trim(psNewCaption & "") <> "" Then
        For liTemp = 1 To miMenuLineCount
            If mvMenuItems(1, liTemp) = psOldCaption Then
                mvMenuItems(1, liTemp) = psNewCaption
                Exit Sub
            End If
        Next
    End If
CleanUp:
    On Error Resume Next
    If llErr <> 0 Then
        On Error GoTo 0
        Err.Raise llErr, , lsErrDesc
    End If
    Exit Sub
ErrorHandler:
    llErr = Err.Number
    lsErrDesc = Err.Description
    Resume CleanUp
End Sub

Public Sub Clear()
'Removes all menu items
    On Error Resume Next
    Dim liTemp As Integer
    miMenuLineCount = 0
    miBoldCommand = 0
End Sub
