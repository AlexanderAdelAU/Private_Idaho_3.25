VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Name:   CWindow
' Author:       Deborah Kurata, InStep Technologies
' Date:         June 1, 1995
' Description:  Generic window class to handle basic form
'               and control processing.
'
' Revisions:
'
Option Explicit

' Constants for error handling
Const iCB_CLEAR_ERR = 3 + vbObjectError + 512
Const iCB_SET_ERR = 4 + vbObjectError + 512
Const iCB_PASTE_ERR = 5 + vbObjectError + 512

' PUBLIC: *****************************************
' Public data members

' When this property is set, the window is made top most
Public Property Let OnTop(frmTop As Form, bSetOnTop As Boolean)
    If bSetOnTop = True Then
      ' Set the window to topmost window
      Call SetWindowPos(frmTop.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
      
    Else
      ' Set the window to not topmost window
      Call SetWindowPos(frmTop.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    End If
End Property

' PRIVATE: *****************************************
' Private data members

' Centers a form.
' If it is an mdi form, it centers it on the MDI parent.
' Otherwise, it centers it on the screen.
' Parameters:
'  frmCenter      Form to center
Public Sub Center(frm As Form, frmParent As Variant)
Dim lHeight, lWidth, lOffset, tOffset
   If frm.MDIChild = True Then
      lHeight = frmParent.ScaleHeight
      lWidth = frmParent.ScaleWidth
   Else
      lHeight = Screen.Height
      lWidth = Screen.Width
   End If

   ' Get left offset
   lOffset = ((lWidth - frm.Width) / 2)
   
   ' Get top offset
   tOffset = ((lHeight - frm.Height) / 2)

   ' Center the form
   frm.Move lOffset, tOffset
End Sub

