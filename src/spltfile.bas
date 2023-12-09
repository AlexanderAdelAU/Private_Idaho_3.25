Attribute VB_Name = "SplitFileModule"
   Option Explicit
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

Public Function SplitFile(FileName As String, FileSize As Double, SegmentSize As Double, PercentShow As Control, Optional NumOfSegments As Integer) As Integer

    Dim SourceBytes As Double
    Dim SourceFile As String
    Dim DestinationFile As String
    Dim SegmentNumber As Integer
    Dim BytesDone As Double
    Dim FPath As String
    Dim FName As String
    Dim FNameNoExt As String
    Dim ErrorCode As Integer
    Dim i As Integer
    Dim J As Integer
    Dim RemainingBytes As Double
    
    On Error GoTo ErrorHandler
    
    'Make sure the file exists
    If FileName = "" Or Dir(FileName) = "" Then
        ErrorCode = 1
        GoTo ErrorHandler
    End If
    
    'Ensure that the segment size is valid
    If SegmentSize = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
    
    'Retrieve the path name where file exists
    Do
        i = i + 1
        'Find the first occurance of the "\" in the FileName string from the right
        J = InStr(Len(FileName) - i, FileName, "\", vbTextCompare)
    Loop Until J > 0
    
    'Extract the file name
    FName = Right$(FileName, Len(FileName) - J)
    
    'Extract the path name
    FPath = Left$(FileName, J)
    
    'Find the file name without extension
    'Find the first occurance of the '.' in the FileName string
    J = InStr(1, FName, ".", vbTextCompare)
    If J = 0 Then   'File name does not contain a '.' character
        FNameNoExt = FName
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(FName, J - 1)
    End If
    
    'Get total number or bytes in the source file
    SourceBytes = FileSize
    
    'Ensure that the resultant file segments will not exceed 999 segments
    'because otherwise we will have incorrect file extensions
    If SourceBytes / SegmentSize >= 1000 Then
        ErrorCode = 3
        GoTo ErrorHandler
    End If

    'Open the source file for binary read
    Open FileName For Binary Access Read As #1 Len = 1
    
    Do
        'Increase the number of segments counter by 1
        SegmentNumber = SegmentNumber + 1
        
        'Compose the file name of the new file to be created (file segment)
        Select Case SegmentNumber
            'changed Fnamenoext to Fname
            Case Is < 10
                DestinationFile = FPath & FName & ".00" & CStr(SegmentNumber)
            Case 10 To 99
                DestinationFile = FPath & FName & ".0" & CStr(SegmentNumber)
            Case 100 To 999
                DestinationFile = FPath & FName & "." & CStr(SegmentNumber)
        End Select
            
        'Create the new file segment and open it for binary write
        Open DestinationFile For Binary Access Write As #2 Len = 1
        
        'Check whether the remaining bytes to process in the source file are
        'less than Segment bytes
        If SourceBytes - BytesDone < SegmentSize Then
            RemainingBytes = SourceBytes - BytesDone
        Else
            RemainingBytes = SegmentSize
        End If
       
       'Read bytes from the source file and write them to the destination file (the current segment file)
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
       Do
            
            Select Case RemainingBytes
                Case Is >= 12000
                    'Read 12000 bytes of data from the source file
                    Get #1, , Bytes.S12000
                    'Write the bytes to the destination file
                    Put #2, , Bytes.S12000
                    'Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    'Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    'Yield to windows and other processes to do their jobs
                    'Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #1, , Bytes.S6000
                    Put #2, , Bytes.S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #1, , Bytes.S3000
                    Put #2, , Bytes.S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #1, , Bytes.S1500
                    Put #2, , Bytes.S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #1, , Bytes.S500
                    Put #2, , Bytes.S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #1, , Bytes.S100
                    Put #2, , Bytes.S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #1, , Bytes.S25
                    Put #2, , Bytes.S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #1, , Bytes.S5
                    Put #2, , Bytes.S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #1, , Bytes.S1
                    Put #2, , Bytes.S1
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
            PercentShow = Int((BytesDone / SourceBytes) * 100)
            'Refresh the form and yield to windows
            DoEvents
        Loop
        
    Loop Until BytesDone = SourceBytes
    'Close the source file
    Close 1
    
    'When the code reaches this point, everything went OK.
    'Acknowledge the number of segments, assign the value '0' to the function and exit
    NumOfSegments = SegmentNumber
    SplitFile = 0
    Exit Function
    
ErrorHandler:
    
    'This is entered only when an error occures
    Select Case ErrorCode
        Case Is = 0 'Unknown error
            Reset   'Close any open files
            SplitFile = 4   'Assign error code 4 to the function
        Case Else 'Assign error code value to the function (1 to 3)
            SplitFile = ErrorCode
    End Select
    
    Exit Function

End Function
