Attribute VB_Name = "modCopy"
Option Explicit

Private Const BUFFER_FLOPPY = 5120
Private Const BUFFER_HDD = 1024000
Private Const BUFFER_CD = 102400
Private Const BUFFER_REMOTE = BUFFER_CD
Private Const BUFFER_REMOVABLE = 102400
Private Const BUFFER_OTHER = BUFFER_CD

Global TotalFilesToCopy As Long
Global TotalBytesToCopy As Double

Global TotalFilesCopied As Long
Global TotalBytesCopied As Double

Dim FileReader As EasyCopy2DLL.HugeBinaryFile
Dim FileWriter As EasyCopy2DLL.HugeBinaryFile

Dim IgnoreAllErrors As Boolean
Dim AlwaysOverwrite As Boolean
Dim DelAfterCopy As Boolean
Dim DelAfterErr As Boolean
Dim ResetAttributes As Boolean

Dim SourceFile As String
Dim TargetFile As String
Dim FileSize As Double
Dim FilePosition As Double
Dim FileData() As Byte

Global WasError As Boolean

Public Sub CountAllFilesAndFolders()
  Dim i As Long
  frmMain.SetToPause True
  frmMain.lBytes.Caption = "Initializing - please wait..."
  With frmMain.pBytes
    .Min = 0
    .Value = .Min
    .Max = CurrentCopyJob.JobItems
  End With
  For i = 1 To CurrentCopyJob.JobItems
    frmMain.lFiles.Caption = CurrentCopyJob.JobItem(i).SourcePath
    frmMain.pBytes.Value = i
    TotalBytesToCopy = (TotalBytesToCopy + CountJobItemSize(CurrentCopyJob.JobItem(i)))
    If ExitNow Then
      Exit For
    End If
  Next
  frmMain.SetToPause False
End Sub

Public Sub StartCopying()
  Dim i As Long
  With frmMain.pBytes
    .Min = 0
    .Value = .Min
    .Max = TotalBytesToCopy
  End With
  With frmMain.pFiles
    .Min = 0
    .Value = 0
    .Max = TotalFilesToCopy
  End With
  Set FileReader = New HugeBinaryFile
  Set FileWriter = New HugeBinaryFile
  Do Until i = CurrentCopyJob.JobItems
    i = (i + 1)
    CopyThisJobItem CurrentCopyJob.JobItem(i)
    If ExitNow Then
      Exit Do
    End If
  Loop
End Sub

Private Function CountJobItemSize(ThisJobItem As JobItem) As Double
  Dim i As Long
  Dim fI As FolderItem
  Dim jS As Double
  If ThisJobItem.IsFolder Then
    Set fI = New FolderItem
    fI.Path = ThisJobItem.SourcePath
    jS = CountFolderItemSize(fI, ThisJobItem.IncludeSubFolders)
  Else
    jS = GetFileSize(ThisJobItem.SourcePath)
    TotalFilesToCopy = (TotalFilesToCopy + 1)
  End If
  CountJobItemSize = jS
End Function

Private Function CountFolderItemSize(ThisFolderItem As FolderItem, DoSubFolders As Boolean) As Double
  Dim i As Long
  Dim fS As Double
  ThisFolderItem.Refresh
  For i = 1 To ThisFolderItem.Files
    fS = (fS + ThisFolderItem.File(i).FileSize)
    TotalFilesToCopy = (TotalFilesToCopy + 1)
  Next
  If DoSubFolders Then
    For i = 1 To ThisFolderItem.SubFolders
      fS = (fS + CountFolderItemSize(ThisFolderItem.SubFolder(i), DoSubFolders))
      DoEvents
      If ExitNow Then
        Exit For
      End If
    Next
  End If
  CountFolderItemSize = fS
End Function

Private Sub CopyThisJobItem(ThisJobItem As JobItem)
  Dim tP As String
  Dim fI As FolderItem
  tP = FixPath(ThisJobItem.TargetPath) & "\"
  If Len(ThisJobItem.SourcePath) <= 3 Then
    tP = tP & Left(ThisJobItem.SourcePath, 1)
  Else
    tP = tP & GetFileName(ThisJobItem.SourcePath)
  End If
  With ThisJobItem
    IgnoreAllErrors = .IgnoreErrors
    AlwaysOverwrite = .Overwrite
    DelAfterCopy = .DeleteAfterCopy
    DelAfterErr = .DeleteAfterError
    If .IsFolder Then
      Set fI = New FolderItem
      fI.Path = .SourcePath
      CopyThisFolderItem fI, tP, .IncludeSubFolders
      Set fI = Nothing
    Else
      If MakeFolder(.TargetPath) Then
        If CopyThisFile(.SourcePath, tP) Then
          TotalFilesCopied = (TotalFilesCopied + 1)
        End If
      End If
    End If
  End With
End Sub

Private Function CopyThisFolderItem(ThisFolderItem As FolderItem, TargetPath As String, DoSubFolders As Boolean) As Boolean
  Dim i As Long
  Dim sP As String
  Dim sF As String
  Dim tF As String
  Dim r As Boolean
  Dim m As String
  If MakeFolder(TargetPath) Then
    r = True
    ThisFolderItem.Refresh
    With ThisFolderItem
      sP = FixPath(.Path)
      i = 0
      Do Until i = .Files
        i = (i + 1)
        sF = sP & "\" & .File(i).FileName
        tF = TargetPath & "\" & .File(i).FileName
        If Not CopyThisFile(sF, tF) Then
          r = False
        End If
        DoEvents
        If ExitNow Then
          Exit Do
        End If
      Loop
      If Not ExitNow Then
        If DoSubFolders Then
          i = 0
          Do Until i = .SubFolders
            i = (i + 1)
            tF = TargetPath & "\" & GetFileName(.SubFolder(i).Path)
            If Not CopyThisFolderItem(.SubFolder(i), tF, DoSubFolders) Then
              r = False
            End If
            If ExitNow Then
              Exit Do
            End If
          Loop
        End If
      End If
    End With
  Else
    WasError = True
  End If
  If Not ExitNow Then
    If Not CopyPathAttributes(ThisFolderItem.Path, TargetPath) Then
      r = False
    End If
    If Not RemoveSourceFolder(ThisFolderItem.Path, r) Then
      r = False
    End If
    CopyThisFolderItem = r
  End If
End Function

Private Function RemoveSourceFolder(SourceFolder As String, CopySuccess As Boolean) As Boolean
  Dim eM As String
  Dim DoRem As Boolean
  If DelAfterCopy Then
    If CopySuccess Then
      DoRem = True
    Else
      If DelAfterErr Then
        DoRem = True
      End If
    End If
  End If
  If DoRem Then
    On Error GoTo RemoveError
    RmDir SourceFolder
  End If
  RemoveSourceFolder = True
RemoveExit:
  On Error GoTo 0
  Exit Function
RemoveError:
  If IgnoreAllErrors Then
    Resume RemoveExit
  Else
    eM = "There was an error trying to remove the recently-copied folder: """ & SourceFolder & "" & vbCrLf & vbCrLf
    eM = eM & "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Could not remove folder!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume RemoveExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume RemoveExit
        Else
          GoTo RemoveError
        End If
    End Select
  End If
End Function

Private Function CopyThisFile(SourcePath As String, TargetPath As String) As Boolean
  Dim r As Boolean
  Dim tV As Double
  Dim buff As Long
  SourceFile = SourcePath
  TargetFile = TargetPath
  If CheckIfTargetExists Then
    FileSize = GetFileSize(SourceFile)
    FilePosition = 0
    With frmMain.pFile
      .Min = 0
      .Value = 0
      .Max = FileSize
    End With
    UpdateStatus
    If Not ExitNow Then
      If OpenFiles() Then
        FileWriter.AutoFlush = BufferSettings.AlwaysWrite
        If CopyFileContents Then
          r = True
        End If
        If Not CloseFiles Then
          r = False
        End If
      End If
    End If
  End If
  If Not ExitNow Then
    If r Then
      TotalFilesCopied = (TotalFilesCopied + 1)
      r = CopyPathAttributes(SourcePath, TargetPath)
    End If
    If Not DeleteSourceFile(r) Then
      r = False
    End If
    CopyThisFile = r
  End If
End Function

Private Function GetBufferSize() As Long
  Dim b1 As Long
  Dim b2 As Long
  Dim r As Long
  b1 = GetBufferSizeForFile(SourceFile)
  b2 = GetBufferSizeForFile(TargetFile)
  If b1 < b2 Then
    If BufferSettings.BufferToUse = buLowest Then
      r = b1
    ElseIf BufferSettings.BufferToUse = byHighest Then
      r = b2
    End If
  Else
    If BufferSettings.BufferToUse = buLowest Then
      r = b2
    ElseIf BufferSettings.BufferToUse = byHighest Then
      r = b1
    End If
  End If
  If BufferSettings.BufferToUse = buAverage Then
    r = ((b1 + b2) / 2)
  End If
  GetBufferSize = r
End Function

Private Function GetBufferSizeForFile(FileName As String) As Long
  Dim r As Long
  Select Case GetDriveType(Left(SourceFile, 1))
    Case EasyCopy2DLL.DRIVE_TYPE.dtRemovable
      r = BufferSettings.RemovableBuffer
    Case EasyCopy2DLL.DRIVE_TYPE.dtHDD
      r = BufferSettings.HardDiskBuffer
    Case EasyCopy2DLL.DRIVE_TYPE.dtCD
      r = BufferSettings.OpticalBuffer
    Case EasyCopy2DLL.DRIVE_TYPE.dtRemote
      r = BufferSettings.NetworkBuffer
    Case Else
      r = BufferSettings.RemovableBuffer
  End Select
  If r = 0 Then
    r = BufferSettings.RemovableBuffer
  End If
  GetBufferSizeForFile = r
End Function

Private Function CopyFileContents() As Boolean
  Dim r As Boolean
  Dim tV As Double
  Dim buff As Long
  r = True
  FilePosition = 0
  buff = GetBufferSize
  ReDim FileData(1 To buff) As Byte
  Do Until FilePosition = FileSize
    If (FileSize - FilePosition) < buff Then
      buff = (FileSize - FilePosition)
      ReDim FileData(1 To buff) As Byte
    End If
    If ReadFromSource Then
      If Not ExitNow Then
        If WriteToTarget Then
          FilePosition = (FilePosition + buff)
          TotalBytesCopied = (TotalBytesCopied + buff)
        Else
          r = False
        End If
      End If
    Else
      r = False
    End If
    If ExitNow Then
      Exit Do
    End If
    UpdateStatus
  Loop
  If Not ExitNow Then
    CopyFileContents = r
  End If
End Function

Private Function ReadFromSource() As Boolean
  Dim eM As String
  On Error GoTo ReadError
  Call FileReader.ReadBytes(FileData())
  ReadFromSource = True
ReadExit:
  On Error GoTo 0
  Exit Function
ReadError:
  If IgnoreAllErrors Then
    Resume ReadExit
  Else
    eM = "There was an error trying to read from the file """ & SourceFile & """." & vbCrLf & vbCrLf
    eM = "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Error reading from source file!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume ReadExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume ReadExit
        Else
          GoTo ReadError
        End If
    End Select
  End If
End Function

Private Function WriteToTarget() As Boolean
  Dim eM As String
  On Error GoTo WriteError
  Call FileWriter.WriteBytes(FileData())
  WriteToTarget = True
WriteExit:
  On Error GoTo 0
  Exit Function
WriteError:
  If IgnoreAllErrors Then
    Resume WriteExit
  Else
    eM = "There was an error trying to write to the file """ & TargetFile & """." & vbCrLf & vbCrLf
    eM = "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Error writing to destination file!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume WriteExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume WriteExit
        Else
          GoTo WriteError
        End If
    End Select
  End If
End Function

Private Function DeleteSourceFile(CopySuccess As Boolean) As Boolean
  Dim eM As String
  Dim DoDel As Boolean
  If DelAfterCopy Then
    If CopySuccess Then
      DoDel = True
    Else
      If DelAfterErr Then
        DoDel = True
      End If
    End If
  End If
  If DoDel Then
    On Error GoTo DelSourceErr
    SetAttr SourceFile, vbNormal
    Kill SourceFile
  End If
  DeleteSourceFile = True
DelSourceExit:
  On Error GoTo 0
  Exit Function
DelSourceErr:
  If IgnoreAllErrors Then
    Resume DelSourceExit
  Else
    eM = "Could not delete the recently copied source-file: """ & SourceFile & "" & vbCrLf & vbCrLf
    eM = eM & "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Could not delete file!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume DelSourceExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume DelSourceExit
        Else
          GoTo DelSourceErr
        End If
    End Select
  End If
End Function

Private Function CheckIfTargetExists() As Boolean
  Dim r As Boolean
  Dim eM As String
  If FileExists(TargetFile) Then
    If AlwaysOverwrite Then
      If DeleteTargetFile Then
        r = True
      End If
    Else
      eM = "The file you are trying to copy already exists at the destination!" & vbCrLf & """" & TargetFile & """" & vbCrLf & vbCrLf
      eM = eM & "Do you want to delete this file before copying?" & vbCrLf & "If you select 'No', the file will not be copied!"
TargetQuestion:
      Select Case MsgBox(eM, vbYesNoCancel Or vbQuestion, "Destination-file already exists!")
        Case vbYes
          If AlwaysOverwrite Then
            r = True
          End If
        Case vbNo
          'Do nothing...
        Case vbCancel
          If CanAbort Then
            ExitNow = True
          Else
            GoTo TargetQuestion
          End If
      End Select
    End If
  Else
    r = True
  End If
  CheckIfTargetExists = r
End Function

Private Function DeleteTargetFile() As Boolean
  Dim eM As String
  On Error GoTo DelTargetError
  SetAttr TargetFile, vbNormal
  Kill TargetFile
DelTargetExit:
  On Error GoTo 0
  DeleteTargetFile = Not FileExists(TargetFile)
  Exit Function
DelTargetError:
  If IgnoreAllErrors Then
    Resume DelTargetExit
  Else
    eM = "Could not delete the pre-existing destination file: """ & TargetFile & """." & vbCrLf & vbCrLf
    eM = eM & "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Could not delete file!")
      Case vbIgnore
        Resume DelTargetExit
      Case vbRetry
        Resume
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume DelTargetExit
        Else
          GoTo DelTargetError
        End If
    End Select
  End If
End Function

Private Function OpenFiles() As Boolean
  If OpenSourceFile() Then
    If OpenTargetFile() Then
      OpenFiles = True
    Else
      CloseSourceFile
    End If
  End If
End Function

Private Function CloseFiles() As Boolean
  Dim r As Boolean
  r = True
  If Not CloseSourceFile Then
    r = False
  End If
  If Not CloseTargetFile Then
    r = False
  End If
  CloseFiles = r
End Function

Private Function OpenSourceFile() As Boolean
  Dim eM As String
  On Error GoTo SourceError
  FileReader.OpenForRead SourceFile
  OpenSourceFile = True
SourceExit:
  On Error GoTo 0
  Exit Function
SourceError:
  If IgnoreAllErrors Then
    Resume SourceExit
  Else
    eM = "There was an error trying to open the file """ & SourceFile & """ for reading." & vbCrLf
    eM = eM & "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Error opening source file!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume SourceExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume SourceExit
        Else
          GoTo SourceError
        End If
    End Select
  End If
End Function

Private Function OpenTargetFile() As Boolean
  Dim eM As String
  On Error GoTo TargetError
  FileWriter.OpenForWrite TargetFile
  FileWriter.AutoFlush = True
  OpenTargetFile = True
TargetExit:
  On Error GoTo 0
  Exit Function
TargetError:
  If IgnoreAllErrors Then
    Resume TargetExit
  Else
    eM = "There was an error trying to open the file """ & TargetFile & """ for writing." & vbCrLf
    eM = eM & "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Error opening destination file!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume TargetExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume TargetExit
        Else
          GoTo TargetError
        End If
    End Select
  End If
End Function

Private Function CloseSourceFile() As Boolean
  FileReader.CloseFile
  CloseSourceFile = True
End Function

Private Function CloseTargetFile() As Boolean
  FileWriter.CloseFile
  CloseTargetFile = True
End Function

Private Function GetPathAttributes(FromPath As String) As VbFileAttribute
  Dim fA As VbFileAttribute
  If FolderExists(FromPath) Then
    fA = vbDirectory
  Else
    fA = vbNormal
  End If
  On Error GoTo GetAttrErr
  fA = GetAttr(FromPath)
GetAttrExit:
  On Error GoTo 0
  GetPathAttributes = fA
  Exit Function
GetAttrErr:
  Resume GetAttrExit
End Function

Private Function CopyPathAttributes(FromPath As String, ToPath As String) As Boolean
  Dim eM As String
  Dim pA As VbFileAttribute
  If Not ResetAttributes Then
    pA = GetPathAttributes(FromPath)
    On Error GoTo AttribError
    SetAttr ToPath, pA
  End If
  CopyPathAttributes = True
AttribExit:
  On Error GoTo 0
  Exit Function
AttribError:
  CopyPathAttributes = True
  Resume AttribExit
  'This fucker makes too much trouble!
  
  If IgnoreAllErrors Then
    Resume AttribExit
  Else
    eM = "There was an error when trying to copy attributes from """ & FromPath & """." & vbCrLf & vbCrLf
    eM = eM & "Error Message: " & Error
    Select Case MsgBox(eM, vbAbortRetryIgnore Or vbCritical, "Error Copying Attributes!")
      Case vbRetry
        Resume
      Case vbIgnore
        Resume AttribExit
      Case vbAbort
        If CanAbort Then
          ExitNow = True
          Resume AttribExit
        Else
          GoTo AttribError
        End If
    End Select
  End If
End Function

Private Sub UpdateStatus()
  With frmMain
    .lBytes = Replace(Replace(LABEL_BYTES, "%b%", GetByteSizeString(TotalBytesCopied)), "%t%", GetByteSizeString(TotalBytesToCopy))
    .pBytes.Value = TotalBytesCopied
    .lFiles.Caption = Replace(Replace(LABEL_FILES, "%f%", Trim(Str(TotalFilesCopied))), "%t%", Trim(Str(TotalFilesToCopy)))
    .pFiles.Value = TotalFilesCopied
    .lFile.Caption = Replace(Replace(LABEL_FILE, "%b%", GetByteSizeString(FilePosition)), "%t%", GetByteSizeString(FileSize))
    .lFileName.Caption = SourceFile
    .pFile.Value = FilePosition
  End With
  DoEvents
End Sub

Public Function CanAbort() As Boolean
  Select Case MsgBox("Are you sure you want to abort the current operation?", vbYesNo Or vbQuestion, "Confirm early exit")
    Case vbYes
      CanAbort = True
  End Select
End Function
