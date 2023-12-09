Attribute VB_Name = "Finger"
    'Workfile:      FINGER.BAS
    'Copyright:     Copyright © 1997 BC
    'Author:        BC
    'Description:   Module for managing finger ops.
    '---------------------------------------------
    'Change History:
    '---------------------------------------------
    'Developer:     BC
    'Change Date:   05/13/97
    'Description:   Created.
    '---------------------------------------------
    
    
    
    
    
    
    
Dim rUser$, rHost$

Public Sub SendFinger()
    'Workfile:      FINGER.BAS
    'Copyright:     Copyright © 1997 BC
    'Author:        BC
    'Description:   Manages basic finger ops for now.
    '---------------------------------------------
    'Change History:
    '---------------------------------------------
    'Developer:     BC
    'Change Date:   05/13/97
    'Description:   Created.
    '---------------------------------------------
    
     

frmPI.IPPort1.WinsockLoaded = True

rHost$ = "anon.lcs.mit.edu"
'rUser$ = "mail2news"
rUser$ = "groups"
'rUser$ = "groups+.*scientology"
'rUser$ = "groups+alt.*(security|privacy)"
'rUser$ = "complaints"
'rUser$ = "finger"
'rUser$ = "mix"
'rUser$ = "mix-info"
'rUser$ = "mix-admin"
'rUser$ = "mail2news"
'rUser$ = "groups"
'rUser$ = "gNym"
'rUser$ = "nymhelp"
'rUser$ = "nymhelp-html"
'rUser$ = "premail-info"

'If tHost = "" Then Beep: Exit Sub

'Raph's List
'rHost$ = "kiwi.cs.berkeley.edu"
'rUser$ = "remailer-list"

frmPI.MessageArea.SelFontName = "arial"
frmPI.MessageArea.SelBold = False
frmPI.MessageArea.SelItalic = False
frmPI.MessageArea.SelStrikeThru = False
frmPI.MessageArea.SelFontSize = 8.25

'Matt Ghio's List
'rUser$ = "remailer.help.all"
'rHost$ = "chaos.taylored.com"

'Matt's pinging service
'rUser$ = "remailer-list"
'rHost$ = "chaos.taylored.com"

'rHost$ = "204.95.228.28"


frmPI.ShowStatus ("")

frmPI.IPPort1.EOL = vbCrLf

'close old connections (if any)
If frmPI.IPPort1.Connected Then frmPI.IPPort1.Connected = False

'x% = InStr(tHost, "@")

'If x% <> 0 Then
'    rUser$ = Left$(tHost, x% - 1)
'    rHost$ = Mid$(tHost, x% + 1)
'Else
'    rUser$ = ""
'    rHost$ = tHost
'End If

frmPI.IPPort1.RemoteHost = rHost$
frmPI.IPPort1.RemotePort = 79 'finger service

'attempt connection
frmPI.ShowStatus ("Connecting to " & rHost$)
frmPI.IPPort1.Connected = True

'wait until the connection is achieved
'(timeout in 10 seconds)
After10Seconds = Now + 10# / (3600# * 24#)
Do Until Now > After10Seconds
    If frmPI.IPPort1.Connected Then
        frmPI.ShowStatus ("Connected to " & frmPI.IPPort1.RemoteHost & "...   ")
        Exit Do
    End If
    DoEvents
Loop
If Not frmPI.IPPort1.Connected Then
    MsgBox "Connection ended."
    frmPI.ShowStatus ("")
    Exit Sub
End If

'send the data
frmPI.IPPort1.DataToSend = rUser$ + vbCrLf

End Sub
