VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUI"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'Class Name:  CUI  Common User Interface
'Author: Alex Cameron
'Date 6th December, 1996 in Singpore
Option Explicit
Private m_fChanged As Boolean

Public Sub onGotFocus()
    If TypeOf Screen.ActiveControl _
        Is TextBox Then
        With Screen.ActiveControl
            .SelStart = 0
            .SelLength = Len(.Text)
            .ForeColor = vbWindowText
        End With
    End If
End Sub

Public Sub onChange()
    
    If TypeOf Screen.ActiveControl Is TextBox Then
        With Screen.ActiveControl
            '.Font.Bold = True
            .backcolor = vbGreen
        End With
    End If
    If TypeOf Screen.ActiveControl Is DBCombo Then
        With Screen.ActiveControl
            '.Font.Bold = True
            .backcolor = vbGreen
        End With
    End If
     If TypeOf Screen.ActiveControl Is ComboBox Then
        With Screen.ActiveControl
            '.Font.Bold = True
            .backcolor = vbGreen
        End With
    End If
    'If TypeOf Screen.ActiveControl Is MaskEdBox Then
    '    With Screen.ActiveControl
            '.Font.Bold = True
      '      .backcolor = vbGreen
      '  End With
   ' End If
End Sub

Public Sub onRule(txt As Object)
    If TypeOf txt Is TextBox Then
        With txt
            .ForeColor = vbRed
        End With
    End If
End Sub

Public Sub onKeyPress(KeyAscii As Integer)
Dim Key As String

Key = Chr$(KeyAscii)
Select Case Key
    Case "0" To "9"
      
    Case vbCr
        SendKeys "{TAB}"
        KeyAscii = 0
     
    Case Chr(8)    'Backspace
    Case "-"
    Case "."
    Case ","
    Case "/"
    Case "$"
    'Case "a" To "z"
    'Case "A" To "Z"
    Case Else
        KeyAscii = 0
End Select
If KeyAscii = 0 Then
    m_fChanged = False
Else
    m_fChanged = True
End If
End Sub

Public Sub newFocus(obj As Object)
        With obj
            obj.Enabled = True
            obj.TabStop = True
            obj.TabIndex = 0
            
        End With
        If TypeOf obj Is CommandButton Then _
               obj.Default = True
        
End Sub






Public Sub onKeyPressAlpha(KeyAscii As Integer)
Dim Key As String

Key = Chr$(KeyAscii)
Select Case Key
    Case "0" To "9"
      
    Case vbCr
        SendKeys "{TAB}"
        KeyAscii = 0
     
    Case Chr(8)    'Backspace
    Case "-"
    Case "."
    Case " "
    Case "a" To "z"
    Case "A" To "Z"
    Case Else
        KeyAscii = 0
End Select
If KeyAscii = 0 Then
    m_fChanged = False
Else
    m_fChanged = True
End If
End Sub

Private Sub Class_Initialize()
m_fChanged = False
End Sub



Public Property Get FieldChanged()
FieldChanged = m_fChanged
End Property

Public Property Let FieldChanged(vNewValue)

End Property
