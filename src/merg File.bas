Attribute VB_Name = "MergeFile"
    'Type declaration for ChunkSize variable
    Type ChunkSize
        S12000 As String * 12000
        S6000 As String * 6000
        S3000 As String * 3000
        S1500 As String * 1500
        S500 As String * 500
        S100 As String * 100
        S25 As String * 25
        S5 As String * 5
        S1 As String * 1
    End Type
    
    'Declare the variable Bytes as of ChunkSize type
    Dim Bytes As ChunkSize

Function MergeFiles(SourceFile As String, PercentShow As Control, Optional NumOfSegments As Integer) As String

    Dim TotalBytes As Long
    Dim DestinationFile As String
    Dim SegmentFile As String
    Dim SegmentNumber As Integer
    Dim Segments As Integer
    Dim BytesDone As Long
    Dim FPath As String
    Dim FName As String
    Dim FNameNoExt As String
    Dim ErrorCode As String
    
    On Error GoTo ErrorHandler
    
    'Make sure the source file name is given and is valid (exists)
    If SourceFile = "" Or Dir(SourceFile) = "" Then
        ErrorCode = "File does not exist."
        GoTo ErrorHandler
    End If
    
    'Find the number of segments of the split file
    'Retrieve the path name where files exist
    Do
        i = i + 1
        'Find the first occurance of the "\" in the SourceFile string from the right
        J = InStr(Len(SourceFile) - i, SourceFile, "\", vbTextCompare)
    Loop Until J > 0
    
    'Extract the file name
    FName = Right$(SourceFile, Len(SourceFile) - J)
    
    'Extract the path name
    FPath = Left$(SourceFile, J)
    
    'Find the file name without extension
    'Find the first occurance of the '.' in the SourceFile string
    J = InStrRev(FName, ".", , vbTextCompare)
    If J = 0 Then   'File name does not contain a '.' character
        FNameNoExt = FName
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(FName, J - 1)
    End If
    
    'Now find the number of segments of the split file that reside in
    'the same directory where the source file is
    'Also count the total number of bytes in the segments (this will be
    'used for the calculation of the percent done value
    Do
        'Increase the number of segments counter by 1
       ' Segments = 0
        Segments = Segments + 1
        
        'Compose the segment file name and check
        Select Case Segments
            Case Is < 10
                SegmentFile = FPath & FNameNoExt & ".00" & CStr(Segments)
            Case 10 To 99
                SegmentFile = FPath & FNameNoExt & ".0" & CStr(Segments)
            Case 100 To 999
                SegmentFile = FPath & FNameNoExt & "." & CStr(Segments)
        End Select
        If Dir(SegmentFile) = "" Then Exit Do
        TotalBytes = TotalBytes + FileLen(SegmentFile)
    Loop
    
    Segments = Segments - 1 'This is the number of segments found
    
    'Check the detected number of segments. If is =0, then the given
    'file name is not a segment file
    If Segments = 0 Then
        ErrorCode = "No segments not valid"
        GoTo ErrorHandler
    End If
    
    'Check if the destination file to be created does exist in the same dir
    'If yes, return error in the function return value
    DestinationFile = FPath & FNameNoExt '& ".src"
    If Dir(DestinationFile) <> "" Then
        ErrorCode = "Destination file already exists"
        GoTo ErrorHandler
    End If

    'Open the destination file for binary write
    Open DestinationFile For Binary Access Write As #1 Len = 1
    Do
        'Increase the number of segments counter by 1
        SegmentNumber = SegmentNumber + 1
        
        'Compose the file name of the new segment file to be opened and read
        Select Case SegmentNumber
            Case Is < 10
                SourceFile = FPath & FNameNoExt & ".00" & CStr(SegmentNumber)
            Case 10 To 99
                SourceFile = FPath & FNameNoExt & ".0" & CStr(SegmentNumber)
            Case 100 To 999
                SourceFile = FPath & FNameNoExt & "." & CStr(SegmentNumber)
        End Select
        frmMergeFile.lblStatus = "Merging file: " & SourceFile
        'Open the source file segment for binary read
        Open SourceFile For Binary Access Read As #2 Len = 1
        DoEvents
        'Get the total number of bytes in the current segment file
        RemainingBytes = FileLen(SourceFile)
       'Read bytes from the source file (the current segment file) and write them to the destination file
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
       Do
            
            Select Case RemainingBytes
                Case Is >= 12000
                    'Read 12000 bytes of data from the source file
                    Get #2, , Bytes.S12000
                    'Write the bytes to the destination file
                    Put #1, , Bytes.S12000
                    'Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    'Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    'Yield to windows and other processes to do their jobs
                    'Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #2, , Bytes.S6000
                    Put #1, , Bytes.S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #2, , Bytes.S3000
                    Put #1, , Bytes.S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #2, , Bytes.S1500
                    Put #1, , Bytes.S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #2, , Bytes.S500
                    Put #1, , Bytes.S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #2, , Bytes.S100
                    Put #1, , Bytes.S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #2, , Bytes.S25
                    Put #1, , Bytes.S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #2, , Bytes.S5
                    Put #1, , Bytes.S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #2, , Bytes.S1
                    Put #1, , Bytes.S1
                    RemainingBytes = RemainingBytes - 1
                    BytesDone = BytesDone + 1
                    DoEvents
                Case Is = 0
                    'When the loop enters here, the segment bytes are completed.
                    'Close the segment file and exit the loop
                    Close 2
                    DoEvents
                    Exit Do
            End Select
            
            'Update the percent control on the form
            PercentShow = Int((BytesDone / TotalBytes) * 100)
            'Refresh the form and yield to windows
            frmMergeFile.lblFileSize = Format(BytesDone, "###,###,###") & " Bytes"
            DoEvents
        Loop
        frmMergeFile.lblFileSize = Format(TotalBytes, "###,###,###") & " Bytes"
    Loop Until SegmentNumber = Segments
    'Close the destination file
    Close 1
    
    'When the code reaches this point, everything went OK.
    'Acknowledge the number of segments, assign the value '0' to the function and exit
    NumOfSegments = Segments
    MergeFiles = Segments & " segments" & " successfully merged."
    Exit Function
    
ErrorHandler:
    
    'This is entered only when an error occures
    Select Case ErrorCode
        Case Is = "" 'Unknown error
            Reset   'Close any open files
            MergeFiles = "Unknown error."   'Assign error code 4 to the function
        Case Else 'Assign error code value to the function
            MergeFiles = ErrorCode
    End Select
    
    Exit Function

End Function
