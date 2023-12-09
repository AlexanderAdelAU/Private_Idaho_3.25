Attribute VB_Name = "NETNEWS"
   
    
    
    
    
    
Global gNNTPLog As Integer
Global gNNTPFile As Integer
Global Const NNTPSTATE = 4

Sub SendNewsCommand()
    'Workfile:      NETNEWS.BAS
    'Copyright:     Copyright © 1997 BC
    'Author:        BC
    'Description:   NNTP ops.
    '---------------------------------------------
    'Change History:
    '---------------------------------------------
    'Developer:     BC
    'Change Date:   05/13/97
    'Description:   Created.
    '---------------------------------------------
    
    
    
    
    
    
    
    

    Dim FileNum As Integer

    On Error GoTo NetNewsError

    gServerState = NNTPSTATE

    gServerName = "netnews.worldnet.att.net"

    frmPI.NNTP1.WinsockLoaded = True
    
    frmPI.NNTP1.NewsServer = gServerName
    
    frmPI.NNTP1.CheckDate = "970509 000000"
    
    frmPI.NNTP1.Action = 1
    
    frmPI.NNTP1.Action = 4
    
    frmPI.NNTP1.CurrentGroup = "alt.online-service.att-worldnet"
    
    frmPI.NNTP1.OverviewRange = "-"
    
    frmPI.NNTP1.Action = 5
    
    frmPI.NNTP1.CurrentArticle = "<3373A3A7.19F@worldnet.att.net>"
    
    frmPI.NNTP1.Action = 6
    
    frmPI.NNTP1.Action = 2
    
    frmPI.NNTP1.WinsockLoaded = False
    '---------------------------------------------
    'frmPI.SMTP1.Action = 2
    '---------------------------------------------
    'HideStatus
    'gServerState = 0
    frmPI.ShowStatus ("Mail successfully sent.")
Exit Sub

NetNewsError:
    'HideStatus
    Select Case Err
        Case 25066
            '---------------------------------------------
            '[10065] No route to host.
            '---------------------------------------------
'            MsgBox "This is an ISP connection error.  Apparently, you have become disconnected from your Internet Service Provider.  Please connect to your ISP and retry."
        Case 26005
            '---------------------------------------------
            '[11004] Valid name, no data record (check DNS setup)
            '---------------------------------------------
'            #If win16 = true Then
'                MsgBox "FATAL INTERNET CONNECTION ERROR.  You tried to perform an internet operation without connecting to your Internet Service Provider first.  The only way to clear this error is to REBOOT your machine.  Please try again later after you reboot and connect to your ISP first."
'            #Else
'                Resume
'            #End If
        Case Else
'            MsgBox "This is an ISP connection error.  Please click on Help Send Feedback and send the attached message through your regular e-mail program.  This will enable you to obtain technical support for your problem."
    End Select
    '---------------------------------------------
    '@@?
    '---------------------------------------------
    Resume
    Exit Sub
End Sub
