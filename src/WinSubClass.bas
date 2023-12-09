Attribute VB_Name = "WinSubClass"
Public defWindowProc As Long

Public Sub SubClass(hWnd As Long)
 On Error Resume Next
 defWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
 End Sub

Public Sub UnSubClass(hWnd As Long)
 If defWindowProc Then
SetWindowLong hWnd, GWL_WNDPROC, defWindowProc
defWindowProc = 0
 End If
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
 Dim retVal As Long
retVal = CallWindowProc(defWindowProc, hWnd, uMsg, wParam, lParam)
If uMsg = WM_NCHITTEST Then
 If retVal = HTCLOSE Then retVal = HTNOWHERE
End If
WindowProc = retVal
End Function

