VERSION 5.00
Begin VB.Form frmRemailerList 
   Caption         =   "Remailer List"
   ClientHeight    =   3945
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRemailer 
      Appearance      =   0  'Flat
      Caption         =   "Remailer Info URLs"
      Height          =   465
      Index           =   3
      Left            =   5880
      TabIndex        =   9
      Top             =   2820
      Width           =   1665
   End
   Begin VB.CommandButton btnRemailer 
      Appearance      =   0  'Flat
      Caption         =   "Add Remailer"
      Height          =   465
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   2820
      Width           =   1665
   End
   Begin VB.CommandButton btnRemailer 
      Appearance      =   0  'Flat
      Caption         =   "Get Remailer Keys"
      Height          =   465
      Index           =   1
      Left            =   2220
      TabIndex        =   7
      Top             =   2820
      Width           =   1725
   End
   Begin VB.CommandButton btnRemailer 
      Appearance      =   0  'Flat
      Caption         =   "Get Remailer Info."
      Height          =   465
      Index           =   0
      Left            =   270
      TabIndex        =   6
      Top             =   2820
      Width           =   1845
   End
   Begin VB.ListBox List3 
      Height          =   2205
      Left            =   270
      TabIndex        =   0
      Top             =   390
      Width           =   7305
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "mod"
      ForeColor       =   &H80000007&
      Height          =   225
      Left            =   300
      TabIndex        =   5
      Top             =   3570
      Width           =   7170
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "latency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3390
      TabIndex        =   4
      Top             =   90
      Width           =   615
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "up time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   4620
      TabIndex        =   3
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   330
      TabIndex        =   2
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "history"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   90
      Width           =   675
   End
   Begin VB.Menu mFIle 
      Caption         =   "File"
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mRemailers 
      Caption         =   "Remailers"
      Begin VB.Menu mRemailerType 
         Caption         =   "Remailer Type"
         Begin VB.Menu mRemailerTypeCypherPunks 
            Caption         =   "Update CypherPunks (Type I)"
         End
         Begin VB.Menu mRemailerTypeMixmaster 
            Caption         =   "Update Mixmaster (Type II)"
         End
      End
      Begin VB.Menu mRemailerUpdate 
         Caption         =   "Update Remailer info"
      End
      Begin VB.Menu mRemailerKeys 
         Caption         =   "Get Remailer keys"
      End
      Begin VB.Menu mEditRemailer 
         Caption         =   "Add Remailer"
      End
      Begin VB.Menu mRemailerURLs 
         Caption         =   "Edit/Update Remailer URLs"
      End
      Begin VB.Menu mBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mMixMasterPath 
         Caption         =   "Set Mixmaster Path"
      End
   End
   Begin VB.Menu mNothing 
      Caption         =   "Help"
      Begin VB.Menu mHelpRemailerCodes 
         Caption         =   "Remailer Codes"
      End
   End
End
Attribute VB_Name = "frmRemailerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnRemailer_Click(Index As Integer)
Dim i As Integer

On Error Resume Next
For i = 0 To btnRemailer.Count - 1
    btnRemailer(i).Enabled = False
Next
Select Case Index
    Case 0
        PIForm(gActivePIInstance).RemailerUpdate
    Case 1
        PIForm(gActivePIInstance).RemailerKeys
    Case 2
        frmAddRemailer.Show
    Case 3
        frmRemailerOptions.Show
        
End Select
For i = 0 To btnRemailer.Count - 1
    btnRemailer(i).Enabled = True
Next
End Sub

Private Sub Form_Activate()
'Dim win As New CWindow
Dim iNumRemailers As Integer
'win.OnTop(Me) = False
gTotalRemailers = 0
On Error Resume Next
'Kill App.Path & "\remailer.txt"
If gRemailerType = REMAILER_CYPHERPUNK Then
    InitializeRemailers (App.Path + "\remailer.htm")
    iNumRemailers = gTotalRemailers
    InitializeRemailers (App.Path + "\private.txt")
Else
    InitializeRemailers (App.Path + "\mixmaster.htm")
    iNumRemailers = gTotalRemailers
End If
SortRemailers
FillRemailerList
If iNumRemailers = 0 Then
    lblstatus = "No remailers found.  Click 'Get Remailer Info' or check for valid URL."
End If

End Sub



Private Sub Form_Load()

'Dim win As New CWindow
'win.OnTop(Me) = False
'Move it so we can see it.
Me.Top = frmMain.Top * 0.2
Me.Left = frmMain.Left * 0.2


gTotalRemailers = 0
If gRemailerType = REMAILER_MIX Then
    InitializeRemailers (App.Path & "\mixmaster.htm")
Else
    InitializeRemailers (App.Path & "\remailer.htm")
    InitializeRemailers (App.Path & "\private.txt")
End If

SortRemailers
FillRemailerList


End Sub

Private Sub Form_Paint()
'Me.Left = 2 * PIForm(gActivePIInstance).Left '+ PIForm(gActivePIInstance).Left
'Me.Top = 2 * PIForm(gActivePIInstance).Top '+ PIForm(gActivePIInstance).Top * 0.5
End Sub

Private Sub Form_Resize()
Dim lWidth As Long
Dim ButtonTop As Long
 If WindowState <> 1 Then
    On Error Resume Next
    lblstatus.Top = ScaleHeight - 1.5 * lblstatus.Height
    btnRemailer(0).Top = lblstatus.Top - btnRemailer(0).Height - 20
    btnRemailer(1).Top = btnRemailer(0).Top
    btnRemailer(2).Top = btnRemailer(0).Top
    btnRemailer(3).Top = btnRemailer(0).Top
    ButtonTop = btnRemailer(0).Top
    'ShowStatus.Top = ButtonTop + btnRemailer(0).Height + 100
    
    List3.Width = Width - List3.Left - Width * 0.1
    btnRemailer(0).Width = (List3.Width / 4.15)
    btnRemailer(1).Width = btnRemailer(0).Width
    btnRemailer(2).Width = btnRemailer(0).Width
    btnRemailer(3).Width = btnRemailer(0).Width
    btnRemailer(1).Left = btnRemailer(0).Left + btnRemailer(0).Width + 0.05 * btnRemailer(0).Width
    btnRemailer(2).Left = btnRemailer(1).Left + btnRemailer(0).Width + 0.05 * btnRemailer(0).Width
     btnRemailer(3).Left = btnRemailer(2).Left + btnRemailer(0).Width + 0.05 * btnRemailer(0).Width
    List3.Height = List3.Top + ButtonTop - btnRemailer(0).Height * 1.8
    lWidth = List3.Width
    lblstatus = ""
 End If
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmRemailerList = Nothing
Set frmRemailerCodes = Nothing
End Sub

Private Sub List3_Click()
Dim iSpace As Integer
Dim RemailerName As String
iSpace = InStr(List3.List(List3.ListIndex), " ")
If iSpace = 0 Then iSpace = Len(List3.List(List3.ListIndex))
If gRemailerType = STANDARD_EMAIL Then
    PIForm(gActivePIInstance).ShowRemailer ("None")
Else
    RemailerName = Mid(List3.List(List3.ListIndex), 1, iSpace - 1)
    PIForm(gActivePIInstance).ShowRemailer (RemailerName)
    If RemailerName = "chain" Then
        PIForm(gActivePIInstance).ShowRemailer ("chain")
    Else
        PIForm(gActivePIInstance).ShowRemailer (GetRemailer(RemailerName))
    End If
End If
End Sub

Private Sub mEditRemailer_Click()
frmAddRemailer.Show
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mHelpRemailerCodes_Click()
frmRemailerCodes.Show
End Sub

Private Sub mMixMasterPath_Click()
SetMixMasterPath
End Sub



Private Sub mRemailerKeys_Click()
PIForm(gActivePIInstance).RemailerKeys
End Sub

Private Sub mRemailerTypeCypherPunks_Click()
gRemailerType = REMAILER_CYPHERPUNK
InitializeRemailers (App.Path + "\remailer.htm")
InitializeRemailers (App.Path + "\private.txt")
SortRemailers
FillRemailerList
frmRemailerList.Caption = "Cypherpunks Remailer List"
End Sub

Private Sub mRemailerTypeMixmaster_Click()
gRemailerType = REMAILER_MIX
InitializeRemailers (App.Path + "\mixmaster.htm")
SortRemailers
FillRemailerList
frmRemailerList.Caption = "Mixmaster List"
End Sub

Private Sub mRemailerUpdate_Click()
PIForm(gActivePIInstance).RemailerUpdate
End Sub

Private Sub mRemailerURLs_Click()
frmRemailerOptions.Show
End Sub

Function CheckForStatus(TheLine As String) As String

    '---------------------------------------------
    'check if a line has status information
    '---------------------------------------------
    If Right(TheLine, 1) = "%" Then
        '---------------------------------------------
        'lets double check and see if there's an @
        '---------------------------------------------
        
        If gRemailerType = REMAILER_MIX Then
                CheckForStatus = TheLine
        End If
        
        If gRemailerType = REMAILER_CYPHERPUNK Then
            If InStr(1, TheLine, "@") Then
                CheckForStatus = TheLine
            Else
                CheckForStatus = ""
            End If
        End If
    
    
    Else
        '---------------------------------------------
        'if not, return null
        '---------------------------------------------
        CheckForStatus = ""
    End If
End Function

Function CheckForType(TheLine As String) As String
    '---------------------------------------------
    'see if the line contains remailer type information
    '---------------------------------------------
    If InStr(1, TheLine, "$remailer{") > 0 Then
        '---------------------------------------------
        'if so, return the string
        '---------------------------------------------
        CheckForType = TheLine
    Else
        '---------------------------------------------
        'if not return null
        '---------------------------------------------
        CheckForType = ""
    End If
End Function
Sub ClearRemailerSort()
   
    gSortRemailer.name = ""
    gSortRemailer.ShortName = ""
    gSortRemailer.Address = ""
    gSortRemailer.latency = ""
    gSortRemailer.uptime = ""
    gSortRemailer.history = ""
    gSortRemailer.cpunk = 0
    gSortRemailer.mix = 0
    gSortRemailer.penet = 0
    gSortRemailer.alpha = 0
    gSortRemailer.newnym = 0
    gSortRemailer.Encrypt = 0
    gSortRemailer.post = 0
    gSortRemailer.latent = 0
    gSortRemailer.hash = 0
    gSortRemailer.cut = 0
    gSortRemailer.Reserved1 = 0
    gSortRemailer.Reserved2 = 0
End Sub

Sub FillNymList()
    
    Dim maxName
    Dim maxUp
    Dim maxLatent As Integer
    Dim tempString As String
    Dim ListHandle As Long
    Dim i As Integer
    Dim tmpstr As String
    
    ReDim lTabPos(2) As Long
    On Error GoTo FillNymListErr
    lTabPos(0) = 10
    lTabPos(1) = 32
    lTabPos(2) = 45
    maxName = 0
    maxUp = 0
    maxLatent = 0
    ListHandle = frmNymServerStats.List1.hWnd
    Call SetTabStops(CLng(ListHandle), 3, lTabPos())
    frmNymServerStats.List1.Clear
    'get the max lengths for padding
    For i = 1 To gTotalMatchedRemailers
        If Len(gMatchedRemailers(i).name) > maxName Then
            maxName = Len(gMatchedRemailers(i).name)
        End If
    Next
    'now create the padded string
    For i = 1 To gTotalMatchedRemailers
        If Len(gMatchedRemailers(i).name) < maxName Then
            tmpstr = gMatchedRemailers(i).name + Space$((maxName - Len(gMatchedRemailers(i).name)) + 3)
            If gRemailerType = REMAILER_MIX Then
                tmpstr = tmpstr + Space$(10)
            End If
        Else
            tmpstr = gMatchedRemailers(i).name + Space$(3)
        End If
        tempString = tmpstr + Chr(9)
        tmpstr = gMatchedRemailers(i).latency
        Select Case Len(tmpstr)
        Case 3
            tmpstr = Space(9) + tmpstr
        Case 4
            tmpstr = Space(6) + tmpstr
        Case 5
            tmpstr = Space(5) + tmpstr
        Case 7
            tmpstr = Space(2) + tmpstr
        End Select
        tempString = tempString + tmpstr + Chr(9)
        tmpstr = gMatchedRemailers(i).uptime
        tmpstr = Space$(7 - Len(tmpstr)) + tmpstr
        tempString = tempString + tmpstr
        frmNymServerStats.List1.AddItem tempString
    Next
    On Error Resume Next
    frmNymServerStats.List1.Selected(0) = 1
    Exit Sub

FillNymListErr:
    MsgBox Err.Description & ": An error occurred filling the Nym Server list.  This is likely to be caused because when the remailer list is not opened. ", vbApplicationModal + vbCritical, App.Title
    Err.Clear
End Sub

Sub FillRemailerList()
   
    Dim maxName
    Dim maxUp
    Dim maxLatent As Integer
    Dim tmpstr1 As String
    Dim tmpstr2 As String
    Dim iListHandle As Long
    Dim i As Integer
    ReDim lTabPos(3) As Long
    
    On Error GoTo FillRemailerListErr
    'first sort remailers
    'SortCPRemailers
    maxName = 0
    maxUp = 0
    maxLatent = 0
    
    '---------------------------------------------
    'set remailer list tab stop locations
    '---------------------------------------------
    lTabPos(0) = 1
    lTabPos(1) = 45
    lTabPos(2) = 65
    lTabPos(3) = 85
    '---------------------------------------------
    'get hWnd for list window and set the tab stops
    '---------------------------------------------
    iListHandle = List3.hWnd
    '---------------------------------------------
    'initialize list
    '---------------------------------------------
    List3.Clear
   ' PIForm(gActivePIInstance).List3.AddItem "none"
    '---------------------------------------------
    'set tab stops for the list
    '---------------------------------------------
    Call SetTabStops(CLng(iListHandle), 4, lTabPos())
    If gSortRemailer.penet = 0 And gSortRemailer.eric = 0 Then
        frmRemailerList.List3.AddItem "chain"
    End If
    '---------------------------------------------
    'determine the length of the longest remailer name
    '---------------------------------------------
    For i = 1 To gTotalMatchedRemailers
        If Len(gMatchedRemailers(i).ShortName) > maxName Then
            maxName = Len(gMatchedRemailers(i).ShortName)
        End If
    Next
    '---------------------------------------------
    'pad remailer names so fields line up in list
    '---------------------------------------------
    For i = 1 To gTotalMatchedRemailers
        '---------------------------------------------
        'build up the remailer data string for each remailer
        '---------------------------------------------
        If Len(gMatchedRemailers(i).ShortName) < maxName Then
            '---------------------------------------------
            'build the name field for the shorter names
            '---------------------------------------------
            tmpstr2 = gMatchedRemailers(i).ShortName
            tmpstr2 = tmpstr2 & Space(maxName - Len(gMatchedRemailers(i).ShortName))
        Else
            '---------------------------------------------
            'build the name field for the longest name(s)
            '---------------------------------------------
            tmpstr2 = gMatchedRemailers(i).ShortName
        End If
        '---------------------------------------------
        'insert a tab
        '---------------------------------------------
        tmpstr1 = tmpstr2 & Space(12) & vbTab
        '--------------------------------------------
        'fetch histroy
        '--------------------------------------------
        'If gRemailerType = REMAILER_MIX Then
            tmpstr2 = gMatchedRemailers(i).history
            tmpstr1 = tmpstr1 & InsertWith(" ", 12 - Len(tmpstr2)) & tmpstr2 & Space(6) & vbTab
            'tmpstr1 = tmpstr1 & tmpstr2 & vbTab
        'End If
        '---------------------------------------------
        'fetch latency time
        '---------------------------------------------
        tmpstr2 = gMatchedRemailers(i).latency
        '---------------------------------------------
        'pad spaces for different latency time formats
        '---------------------------------------------
        tmpstr1 = tmpstr1 & Space(10 - Len(tmpstr2)) & tmpstr2 & Space(6) & vbTab
       'tmpstr1 = tmpstr1 & tmpstr2 & vbTab
        '---------------------------------------------
        'fetch uptime percentage
        '---------------------------------------------
        tmpstr2 = gMatchedRemailers(i).uptime
        '---------------------------------------------
        'prepend with appropriate number of spaces
        '---------------------------------------------
        '---------------------------------------------
        'build remailer data string to add to listbox
        '---------------------------------------------
        tmpstr1 = tmpstr1 & Space(10 - Len(tmpstr2)) & tmpstr2 '& Space(5)
        'tmpstr1 = tmpstr1 & tmpstr2 '& Space(5)
        '---------------------------------------------
        'add the string
        '---------------------------------------------
        List3.AddItem tmpstr1
    Next
    '---------------------------------------------
    'select the remailer list (equivalent to clicking on it)
    '---------------------------------------------
    List3.Refresh
    List3.Selected(0) = 1
    Exit Sub

FillRemailerListErr:
    MsgBox Err.Description & " You need to Fill the remailer list."
    Err.Clear
End Sub

Function GetRemailerIndex(theName As String) As Integer

    Dim Found As Integer
    Dim i As Integer
    
    Found = 0
    
    '---------------------------------------------
    'returns the remailer array index
    '---------------------------------------------
    If gTotalRemailers = 0 Then
        '---------------------------------------------
        'nothing in the array yet, so put the first remailer name in it
        '---------------------------------------------
        InitializeRemailerArray (1)
        gRemailerArray(1).name = theName
        gTotalRemailers = 1
        GetRemailerIndex = 1
    Else
        '---------------------------------------------
        'the remailers are there
        '---------------------------------------------
        For i = 1 To gTotalRemailers
            If gRemailerArray(i).name = theName Then
            GetRemailerIndex = i
            Found = 1
            Exit For
        End If
    Next
    '---------------------------------------------
    'couldnt find the name, so add it to the array
    '---------------------------------------------
    If Found = 0 Then
        If i < UBound(gRemailerArray) Then
            InitializeRemailerArray (i)
            gRemailerArray(i).name = theName
            gTotalRemailers = gTotalRemailers + 1
            GetRemailerIndex = i
        End If
    End If
End If
End Function
  Sub InitializeMixRemailers(theFile As String)
    
    Dim FileNum
    Dim y As Integer
    Dim J As Integer
    Dim i As Integer
    Dim TheLine As String
   'Dim tempstr As String
    Dim tmpstr1 As String
    
    On Error GoTo InitializeMixRemailersErr
    FileNum = FreeFile
    J = 0
    Open theFile For Input As FileNum
    Line Input #FileNum, TheLine
    Close FileNum
    tmpstr1 = TheLine
    If Len(tmpstr1) > 1024 Then
        MsgBox ("Converting mixmaster.htm file.  This may take a minute.")
        tmpstr1 = ConvertStr(tmpstr1)
        WriteStrToFile App.Path + "\mixmaster.htm", tmpstr1
        TheLine = ""
    End If
    Open theFile For Input As FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, TheLine
        'see if the line contains Mixmaster info
        If TheLine <> "" Then
            If Mid$(TheLine, Len(TheLine), 1) = "%" Then
                J = J + 1
                'get name
                'grab position of first non-space character after first space
                i = InStr(1, TheLine, " ")
                'everything between is the name
                gMatchedRemailers(J).name = Mid$(TheLine, 1, i - 1)
                
                               
                'get uptime
                'go to the end of the line and parse left until a space
                y = 0
                For y = Len(TheLine) To 1 Step -1
                    If Mid$(TheLine, y, 1) = " " Then
                        Exit For
                    End If
                Next
                gMatchedRemailers(J).uptime = Mid$(TheLine, y + 1, Len(TheLine) - y)
                
                'get latency
                'now parse to the left until we hit a non-space
                For i = y To 1 Step -1
                    If Mid$(TheLine, i, 1) <> " " Then
                        Exit For
                    End If
                Next
                'we have the left position, now we need the right position
                For y = i To 1 Step -1
                    If Mid$(TheLine, y, 1) = " " Then
                        Exit For
                    End If
                Next
                gMatchedRemailers(J).latency = Mid$(TheLine, y + 1, i - y)
                 
                'now get history
                For y = y To 1 Step -1
                    Select Case Mid$(TheLine, y, 1)
                        Case "+", "-", "*", "_", "."
                            Exit For
                    End Select
           
                    'If Not Mid$(TheLine, y, 1) = " " Then
                        'Exit For
                'End If
                Next
                i = y
                For y = y To 1 Step -1
                    If Mid$(TheLine, y, 1) = " " Then
                        Exit For
                    End If
                Next
                If y <= 8 Then
                    gMatchedRemailers(J).history = ""
                Else
                    gMatchedRemailers(J).history = Mid$(TheLine, y + 1, i - y)
                End If
                
                gTotalMatchedRemailers = J
            End If
        End If
    Wend
    Close FileNum
    lblstatus = "Remailer update as of " + Format(FileDateTime(theFile), "ddd, ddddd ttttt")
    Exit Sub

InitializeMixRemailersErr:
    Reset
    MsgBox Err.Description & ". Can't find the some mixmaster files.  Ensure that the mixmaster.htm has been updated. You need to Initialize Mix Remailers"
    Err.Clear
End Sub
  
    

Sub InitializeRemailerArray(theElement As Integer)
        
    gRemailerArray(theElement).name = ""
    gRemailerArray(theElement).ShortName = ""
    gRemailerArray(theElement).Address = ""
    gRemailerArray(theElement).latency = ""
    gRemailerArray(theElement).latency = ""
    gRemailerArray(theElement).history = ""
    gRemailerArray(theElement).cpunk = 0
    gRemailerArray(theElement).eric = 0
    gRemailerArray(theElement).mix = 0
    gRemailerArray(theElement).penet = 0
    gRemailerArray(theElement).alpha = 0
    gRemailerArray(theElement).newnym = 0
    gRemailerArray(theElement).Encrypt = 0
    gRemailerArray(theElement).post = 0
    gRemailerArray(theElement).latent = 0
    gRemailerArray(theElement).hash = 0
    gRemailerArray(theElement).cut = 0
    gRemailerArray(theElement).Reserved1 = 0
    gRemailerArray(theElement).Reserved2 = 0
End Sub

Sub InitializeRemailers(TheFileName As String)
    Dim FileNum As Integer
    Dim FileNum2 As Integer
    Dim FileNum3 As Integer
    Dim TheLine As String
    Dim tmpstr2 As String
    Dim tmpstr1 As String
    Dim Item As String
    Dim i As Integer
    Dim CreateTextFile As Boolean
    
    On Error Resume Next
      
    '----------------------------
    'Initialise the remailers to 0 so a refill is forced
    '-----------------------------
  
    'If Not iFileExists(TheFileName) Then
       ' MsgBox "Some files that are required to load the remailer list appear to be missing.  Please using the 'Get Remailer Info.' to create the files needed.", vbApplicationModal + vbCritical, "Initialise Remailers"
       ' Exit Sub
    'End If
    CreateTextFile = False
    
    On Error GoTo InitializeRemailerErr
    
   PIForm(gActivePIInstance).ShowStatus 1, "Loading remailer data..."
    
    If InStr(1, TheFileName, ".htm", vbTextCompare) Then
        If iFileExists(TheFileName) Then
            FileNum2 = FreeFile
            Open App.Path & "\remailer.txt" For Output As FileNum2
            CreateTextFile = True
        Else
            PIForm(gActivePIInstance).ShowStatus 1, "File: " & TheFileName & " was not found.  Data for this remailer not created."
            Exit Sub
        End If
    End If
    
    If InStr(1, TheFileName, "private.txt", vbTextCompare) Then
        If iFileExists(TheFileName) Then
            FileNum2 = FreeFile
            Open App.Path & "\remailer.txt" For Append As FileNum2
            CreateTextFile = True
        Else
            PIForm(gActivePIInstance).ShowStatus 1, "No Private Remailer data was found."
            Exit Sub
        End If
    End If
    
    '---------------------------------------------
    'usual case, process the converted html file
    '---------------------------------------------
    FileNum = FreeFile
    Open TheFileName For Input As #FileNum
    While Not EOF(FileNum) And gCancelAction = False
        Line Input #FileNum, TheLine
        'see if the line contains type info
        tmpstr1 = TheLine
        tmpstr2 = CheckForType(tmpstr1)
        'line contains type info
        If tmpstr2 <> "" Then
            '---------------------------------------------
            'valid remailer information, so parse the line
            'and update the array
            '---------------------------------------------
            If CreateTextFile Then Print #FileNum2, tmpstr1
            ParseRemailerLine (tmpstr1)
        Else
                
            'line does not contain type info,
            'so see if the line contains status info
            '---------------------------------------------
            tmpstr2 = CheckForStatus(tmpstr1)
            '---------------------------------------------
            'line contains status info
            '---------------------------------------------
            If tmpstr2 <> "" Then
                If CreateTextFile Then Print #FileNum2, tmpstr1
                ParseRemailerLine (tmpstr1)
            End If
        End If
    Wend
    Close FileNum
    If CreateTextFile Then
        Close FileNum2
       ' Close FileNum3
    End If
    
    '---------------------------------------------
    'Put the time and date of last update on form1
    '---------------------------------------------
    If gCancelAction Then
        gCancelAction = False
        lblstatus = "Remailer data not updated due to error."
    Else
        If InStr(1, TheFileName, "private") = 0 Then
            lblstatus = "Remailer update as of " + Format$(FileDateTime(TheFileName), "mm/dd/yy hh:mm")
        End If
    End If
    PIForm(gActivePIInstance).ShowStatus 1, ""
    Reset
    Exit Sub

InitializeRemailerErr:
    Reset
    PIForm(gActivePIInstance).ShowStatus 1, ""
    MsgBox Err.Description & " (in procedure InitializeRemailers)"
    Err.Clear
End Sub

Sub ParseRemailerLine(TheLine As String)
Dim tempName As String
Dim RemailerName As String
Dim remailerIndex As Integer
Dim x As Integer
Dim y As Integer
Dim tmpstr As String
'Exit Sub
    On Error GoTo ParseRemailerLineErr
    
    '---------------------------------------------
    'we have valid remailer information, so parse
    'the line and insert into the array
    '---------------------------------------------
    
    '---------------------------------------------
    'is this info or remailer
    '---------------------------------------------
    If Left$(TheLine, 10) = "$remailer{" Then
        '---------------------------------------------
        'its info - so grab the name
        '---------------------------------------------
        x = InStr(1, TheLine, "&lt;") ' This is the raw html and must be here
        y = InStr(1, TheLine, "&gt;") ' This is the raw html and must be here
        If x = 0 Or y = 0 Then
            x = InStr(1, TheLine, "<")
            y = InStr(x, TheLine, ">")
            x = x + 1 ' Jump over >
        Else
            x = x + 4 ' Jump over &lt;
        End If
        If x = 0 Or y = 0 Then Exit Sub
        tempName = Mid(TheLine, x, y - x)
      
      '----This is new to just get the remailer name not the address----'
        x = InStr(1, TheLine, "{&quot;")
        y = InStr(1, TheLine, "&quot;}")
        If x = 0 Or y = 0 Then
            x = InStr(1, TheLine, "{""") + 2 'jump over "{
            y = InStr(1, TheLine, """}")
        Else
            x = x + 7 ' Jump over &quot};
        End If
        RemailerName = Mid(TheLine, x, y - x)
        '--end new stuff
        
        '---------------------------------------------
        'see if it exists in the global array
        '---------------------------------------------
        remailerIndex = GetRemailerIndex(tempName)
        gRemailerArray(remailerIndex).name = tempName
        gRemailerArray(remailerIndex).ShortName = RemailerName
        '---------------------------------------------
        'check for cypherpunk type
        '---------------------------------------------
        If InStr(1, TheLine, " cpunk") Then
            gRemailerArray(remailerIndex).cpunk = 1
        End If
        '---------------------------------------------
        'check for eric type
        '---------------------------------------------
        If InStr(1, TheLine, " eric") Then
            gRemailerArray(remailerIndex).eric = 1
        End If
        '---------------------------------------------
        'check for penet type
        '---------------------------------------------
        If InStr(1, TheLine, " penet") Then
            gRemailerArray(remailerIndex).penet = 1
        End If
        '---------------------------------------------
        'check for mix type
        '---------------------------------------------
        If InStr(1, TheLine, " mix") Then
            gRemailerArray(remailerIndex).mix = 1
        End If
        '---------------------------------------------
        'check for alpha type
        '---------------------------------------------
        If InStr(1, TheLine, " alpha") Then
            gRemailerArray(remailerIndex).alpha = 1
        End If
        '---------------------------------------------
        'check for newnym type
        '---------------------------------------------
        If InStr(1, TheLine, " newnym") Then
            gRemailerArray(remailerIndex).newnym = 1
        End If
        '---------------------------------------------
        'check for pgp
        '---------------------------------------------
        If InStr(1, TheLine, " pgp") Then
            gRemailerArray(remailerIndex).Encrypt = 1
        End If
        '---------------------------------------------
        'check for latent
        '---------------------------------------------
        If InStr(1, TheLine, " latent") Then
            gRemailerArray(remailerIndex).latent = 1
        End If
        '---------------------------------------------
        'check for hash
        '---------------------------------------------
        If InStr(1, TheLine, " hash") Then
            gRemailerArray(remailerIndex).hash = 1
        End If
        '---------------------------------------------
        'check for cut
        '---------------------------------------------
        If InStr(1, TheLine, " cut") Then
            gRemailerArray(remailerIndex).cut = 1
        End If
        '---------------------------------------------
        'check for post
        '---------------------------------------------
        If InStr(1, TheLine, " post") Then
            gRemailerArray(remailerIndex).post = 1
        End If
    Else
        '---------------------------------------------
        'This is status information
        'get name
        'grab position of first non-space character after first space
        '---------------------------------------------
        If Right$(TheLine, 1) = "%" Then
          x = InStr(1, TheLine, " ")
          tempName = Mid$(TheLine, 1, x - 1)
        
          If gRemailerType = REMAILER_CYPHERPUNK Then
        
            '---------------------------------------------
            'grab position of first space after beginning position
            '---------------------------------------------
            For y = x To (Len(TheLine))
                If Mid$(TheLine, y, 1) <> " " Then
                    Exit For
                End If
            Next
            '---------------------------------------------
            'now we need to get the next space
            '---------------------------------------------
            tmpstr = Mid$(TheLine, y, Len(TheLine) - y)
            x = InStr(1, tmpstr, " ")
            tempName = Mid$(tmpstr, 1, x - 1)
            '---------------------------------------------
            'everything between is the name
            '---------------------------------------------
            remailerIndex = GetRemailerIndex(tempName) 'Mid$(tmpstr, 1, x - 1))
            gRemailerArray(remailerIndex).name = tempName
        Else
            tmpstr = TheLine
            remailerIndex = GetRemailerIndex(tempName) 'Mid$(tmpstr, 1, x - 1))
            gRemailerArray(remailerIndex).ShortName = tempName
        End If
        
        '--------------------------
        'Now get history
        '----------------------------
        Dim i, J As Integer
        'Dim ParsedString As String
        Dim ParsedLine As String
       'Get Uptime
        i = InStrRev(tmpstr, " ")
        ParsedLine = RTrim(tmpstr)
        i = InStrRev(ParsedLine, " ")
        gRemailerArray(remailerIndex).uptime = LTrim(Mid(ParsedLine, i, Len(ParsedLine) - i)) & "%"
        
        ParsedLine = RTrim(Mid(ParsedLine, 1, i))
        i = InStrRev(ParsedLine, " ")
        If Not i = 0 Then gRemailerArray(remailerIndex).latency = LTrim(Mid(ParsedLine, i, Len(ParsedLine) - i + 1))
        
        ParsedLine = RTrim(Mid(ParsedLine, 1, i))
        i = InStrRev(ParsedLine, " ")
        If gRemailerType = REMAILER_CYPHERPUNK Then
            J = InStr(1, ParsedLine, "@")
            J = InStr(J, ParsedLine, " ")
        Else
            J = InStr(1, ParsedLine, " ")
        End If
        If Not i = 0 Then gRemailerArray(remailerIndex).history = LTrim(Mid(ParsedLine, J, Len(ParsedLine) - J + 1))
    End If
   End If
    
    
    'End If
Exit Sub

ParseRemailerLineErr:
    ' "Error parsing remailer info." & Err.Description, vbCritical + vbApplicationModal
    gCancelAction = True
    Err.Clear
End Sub
'
'Sort the remailers and then transfer to the Matched remailer list
'
Sub SortRemailers()
      
    Dim isMix
    Dim isAlpha
    Dim isPGP
    Dim isLatent
    Dim isCut
    Dim isHeader
    Dim isCP
    Dim isPenet
    Dim isSoda
    Dim isPost As Integer
    Dim FileNum As Integer
    Dim y As Integer
    Dim i As Integer
    
    y = 0
    isMix = 0
    isAlpha = 0
    isPGP = 0
    isCut = 0
    isHeader = 0
    isLatent = 0
    isCP = 0
    isSoda = 0
    isPenet = 0
    isPost = 0
   'For i = 0 To 59
       ' gMatchedRemailers(i) = gRemailerArray(i)
       ' gTotalMatchedRemailers = i
    'Next
  ' Exit Sub
    'if we're dealing with mix, need to handle differently
    'gMatchedRemailers(y) = gRemailerArray(i)
    'gMatchedRemailers(34).history = gRemailerArray(5).history
    'If Not gRemailerType = REMAILER_MIX Then
        gTotalMatchedRemailers = 0
        For i = 1 To gTotalRemailers
          If Trim(gRemailerArray(i).history) <> "" Then     'Debug.Print gRemailerArray(i).history
            If gRemailerType = REMAILER_CYPHERPUNK Then
                Dim s As String
                
                If gRemailerArray(i).cpunk = 1 Then
                    isCP = True
                Else
                    isCP = False
                End If
            End If
            If gRemailerType = REMAILER_MIX Then
                'If gRemailerArray(i).mix = 1 Then
                    isMix = True
                'End If
            End If
        
            'now compare and filter
            If ((isCP And gRemailerType = REMAILER_CYPHERPUNK)) Then
                If gRemailerArray(i).alpha = 0 Then
                    y = y + 1
                    gMatchedRemailers(y) = gRemailerArray(i)
                    s = gMatchedRemailers(y).name
                    gTotalMatchedRemailers = y
                    isMix = 0
                    isAlpha = 0
                    isPGP = 0
                    isCut = 0
                    isLatent = 0
                    isHeader = 0
                    isCP = 0
                    isSoda = 0
                    isPenet = 0
                    isPost = 0
                    'EXIT FOR
                End If
          End If
          If ((isMix And gRemailerType = REMAILER_MIX)) Then
                    y = y + 1
                    gMatchedRemailers(y) = gRemailerArray(i)
                    gTotalMatchedRemailers = y
                    'gMatchedRemailers(y).Name = gMatchedRemailers(y).Name
         End If
       End If
    Next

End Sub

Sub SortRemailersByType()
    Dim x
    Dim y As Integer

    y = 0
    gTotalMatchedRemailers = 0
    '[debug]
    If gSortRemailer.cpunk Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).cpunk = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.mix Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).mix = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.penet Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).penet = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.eric Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).eric = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    ElseIf gSortRemailer.alpha Or gSortRemailer.newnym Then
        For x = 1 To gTotalRemailers
            If gRemailerArray(x).alpha = 1 Or gRemailerArray(x).newnym = 1 Then
                y = y + 1
                gMatchedRemailers(y) = gRemailerArray(x)
                gTotalMatchedRemailers = y
            End If
        Next
    End If
End Sub

Function InsertWith(FillChar As String, Length As Long) As String
Dim i As Integer
InsertWith = ""
For i = 1 To Length
    InsertWith = InsertWith & FillChar
Next
End Function

Public Function GetRemailer(theName As String) As String

    Dim i As Integer
    GetRemailer = ""
    For i = 1 To gTotalMatchedRemailers
            If gMatchedRemailers(i).ShortName = theName Then
                GetRemailer = gMatchedRemailers(i).name
            Exit For
        End If
    Next

End Function

