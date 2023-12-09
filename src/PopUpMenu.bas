Attribute VB_Name = "PopUpMenu"
Option Explicit

'Private lngChkMark As Long, xChk As Long, yChk As Long

Private nHeight As Integer
Private nWidth As Integer

'Public Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
'Public Const MF_BITMAP = &H4&
Private Const SM_CXMENUCHECK = 71   ' Use instead of GetMenuCheckMarkDimensions()
Private Const SM_CYMENUCHECK = 72

'Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Declare Function GetSystemMetrics Lib "user32" _
  (ByVal nIndex As Long) As Long
  
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetSubMenu Lib "user32" _
  (ByVal hmenu As Long, ByVal nPos As Long) As Long

Declare Function SetMenuItemBitmaps Lib "user32" _
  (ByVal hmenu As Long, ByVal nPosition As Long, _
  ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
  ByVal hBitmapChecked As Long) As Long

'Declare Function GetMenuItemID Lib "user32" _
  '(ByVal hmenu As Long, ByVal nPos As Long) As Long
  


'Function getLoHiWord(lparam As Long, loWord As Long, hiWord As Long)
      'loWord = lparam And &HFFFF&
      'hiWord = lparam \ &H10000 And &HFFFF&
      'getLoHiWord = 1
'End Function

'Function getHiLoWord(ByVal lngParam As Integer) As HILOWORD
  '    Dim hilo As HILOWORD
 '     hilo.loWord = lngParam And &HFF&
 '     hilo.hiWord = lngParam \ &H100 And &HFF&
 '     getHiLoWord = hilo
'End Function

 
 
Public Sub SetMenuBmp(frm As Form)
Dim hmenu As Long
Dim hsubmnuFile As Long
Dim hsubmnuEdit As Long
Dim hresult As Long
Dim imgX As ListImage



hmenu = GetMenu(frm.hwnd)
hsubmnuFile = GetSubMenu(hmenu, 3)
'hsubmnuEdit = GetSubMenu(hmenu, 1)
nWidth = GetSystemMetrics(SM_CXMENUCHECK)
nHeight = GetSystemMetrics(SM_CYMENUCHECK)

Dim colFiles As New Collection
'With colFiles
   ' .Add ("new.bmp")
   ' .Add ("open.bmp")
    '.Add ("save.bmp")
    ''.Add ("cut.bmp")
    '.Add ("copy.bmp")
    '.Add ("paste.bmp")
'End With
Dim img As ListImage
Dim x As Integer
'For x = 1 To colFiles.Count
    'With frmMain.PictureClip1
       ' Set .Picture = LoadPicture(colFiles.Item(x))
       ' .ClipHeight = .Height
        '.ClipWidth = .Width
        '.StretchX = nWidth
        '.StretchY = nHeight
        '.Picture = .Clip
       ' Set img = frmMain.ImageList1.ListImages.Add(, , .Picture)
   ' End With
    x = 1
    With frmMain.ImageList1
        If x <= 3 Then
        Call SetMenuItemBitmaps(hsubmnuFile, x - 1, MF_BYPOSITION, _
            .ListImages(22).Picture, 0)
        Else
        Call SetMenuItemBitmaps(hsubmnuEdit, x - 4, MF_BYPOSITION, _
            .ListImages(22).Picture, 0)
        End If
    End With
'Next x
End Sub
