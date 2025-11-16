Attribute VB_Name = "modLoadSave"
Option Explicit

Public Const FILE_EXTENSION = ".ecj"

Private Const LEADING_BYTES = "EasyCopy2Project"

Private Const SAVE_NAME = "Name"
Private Const SAVE_ITEMSTART = "Item"
Private Const SAVE_SOURCE = "Source"
Private Const SAVE_TARGET = "Target"
Private Const SAVE_SUBS = "Subs"
Private Const SAVE_ATTRIB = "Attrib"
Private Const SAVE_OVERWRITE = "Overwrite"
Private Const SAVE_IGNORE = "Ignore"
Private Const SAVE_DELETE = "DelSource"
Private Const SAVE_DELERR = "DelError"
Private Const SAVE_ITEMEND = "ItemEnd"
Private Const SAVE_SEP1 = ":"
Private Const SAVE_SEP2 = vbNullChar

Public Function SaveCopyJob(ThisCopyJob As CopyJob) As Boolean
  If CreateFolderA(GetSavePath) Then
    SaveCopyJob = SaveCopyJobData(ThisCopyJob, ThisCopyJob.FileName)
  End If
End Function

Public Function LoadCopyJob(ThisCopyJob As CopyJob) As Boolean
  Dim fData As String
  fData = RemoveLeadingBytes(ReadDataFromFile(ThisCopyJob.FileName))
  If Not fData = "" Then
    LoadCopyJob = FillCopyJobWithData(ThisCopyJob, DecryptDataA(fData))
    ThisCopyJob.JobNotChanged
  End If
End Function

Public Function GetNewCopyJobFileName() As String
  Dim fP As String
  Dim fN As String
  Dim nF As String
  fP = GetSavePath
  Do
    fN = GetRandomStringA(GetRandomNumberA(5, 26)) & FILE_EXTENSION
    nF = fP & "\" & fN
    If Not FileExistsA(nF) Then
      GetNewCopyJobFileName = nF
      Exit Do
    End If
  Loop
End Function

Public Function SaveCopyJobToFile(ThisCopyJob As CopyJob, FileName As String) As Boolean
  If Not FileName = "" Then
    If CreateFolderA(GetParentFolderA(FileName)) Then
      SaveCopyJobToFile = SaveCopyJobData(ThisCopyJob, FileName)
    End If
  End If
End Function

Public Function LoadCopyJobFromFileA(FileName As String) As CopyJob
  Dim nJ As CopyJob
  Set nJ = New CopyJob
  nJ.FileName = FileName
  If LoadCopyJob(nJ) Then
    Set LoadCopyJobFromFileA = nJ
  End If
End Function

Private Function SaveCopyJobData(ThisCopyJob As CopyJob, FileName As String) As Boolean
  SaveCopyJobData = WriteDataToFile(FileName, CollectCopyJobData(ThisCopyJob))
End Function

Private Function FillCopyJobWithData(ThisCopyJob As CopyJob, JobData As String) As Boolean
  Dim i As Long
  Dim jD As String
  Dim jC As String
  Dim jV As String
  Dim cItem As JobItem
  Dim WasError As Boolean
  jD = JobData
  On Error GoTo FillJobError
  Do Until jD = ""
    i = InStr(jD, SAVE_SEP2)
    If i = 0 Then
      jC = jD
      jD = ""
    Else
      jC = Left(jD, (i - 1))
      jD = Right(jD, (Len(jD) - i))
    End If
    i = InStr(jC, SAVE_SEP1)
    If i = 0 Then
      jV = ""
    Else
      jV = Right(jC, (Len(jC) - i))
      jC = Left(jC, (i - 1))
    End If
    Select Case jC
      Case SAVE_NAME
        ThisCopyJob.Name = jV
      Case SAVE_ITEMSTART
        Set cItem = ThisCopyJob.AddJobItem("")
      Case SAVE_SOURCE
        cItem.SourcePath = jV
      Case SAVE_TARGET
        cItem.TargetPath = jV
      Case SAVE_SUBS
        cItem.IncludeSubFolders = GetBooleanFromString(jV)
      Case SAVE_ATTRIB
        cItem.ResetAttributes = GetBooleanFromString(jV)
      Case SAVE_OVERWRITE
        cItem.Overwrite = GetBooleanFromString(jV)
      Case SAVE_IGNORE
        cItem.IgnoreErrors = GetBooleanFromString(jV)
      Case SAVE_DELETE
        cItem.DeleteAfterCopy = GetBooleanFromString(jV)
      Case SAVE_DELERR
        cItem.DeleteAfterError = GetBooleanFromString(jV)
      Case SAVE_ITEMEND
        Set cItem = Nothing
    End Select
  Loop
  FillCopyJobWithData = Not WasError
FillJobExit:
  On Error GoTo 0
  Exit Function
FillJobError:
  WasError = True
  Resume Next
End Function

Private Function CollectCopyJobData(ThisCopyJob As CopyJob) As String
  Dim i As Long
  Dim cData As String
  With ThisCopyJob
    cData = SAVE_NAME & SAVE_SEP1 & .Name & SAVE_SEP2
    For i = 1 To .JobItems
      cData = cData & SAVE_ITEMSTART & SAVE_SEP2
      With .JobItem(i)
        If Not .SourcePath = "" Then
          cData = cData & SAVE_SOURCE & SAVE_SEP1 & .SourcePath & SAVE_SEP2
        End If
        If Not .TargetPath = "" Then
          cData = cData & SAVE_TARGET & SAVE_SEP1 & .TargetPath & SAVE_SEP2
        End If
        If .IncludeSubFolders Then
          cData = cData & SAVE_SUBS & SAVE_SEP1 & GetStringFromBoolean(.IncludeSubFolders) & SAVE_SEP2
        End If
        If .ResetAttributes Then
          cData = cData & SAVE_ATTRIB & SAVE_SEP1 & GetBooleanFromString(.ResetAttributes) & SAVE_SEP2
        End If
        If .Overwrite Then
          cData = cData & SAVE_OVERWRITE & SAVE_SEP1 & GetBooleanFromString(.Overwrite) & SAVE_SEP2
        End If
        If .IgnoreErrors Then
          cData = cData & SAVE_IGNORE & SAVE_SEP1 & GetBooleanFromString(.IgnoreErrors) & SAVE_SEP2
        End If
        If .DeleteAfterCopy Then
          cData = cData & SAVE_DELETE & SAVE_SEP1 & GetBooleanFromString(.DeleteAfterCopy) & SAVE_SEP2
        End If
        If .DeleteAfterError Then
          cData = cData & SAVE_DELERR & SAVE_SEP1 & GetBooleanFromString(.DeleteAfterError) & SAVE_SEP2
        End If
      End With
      cData = cData & SAVE_ITEMEND & SAVE_SEP2
    Next
  End With
  CollectCopyJobData = GetLeadingBytes & EncryptDataA(cData)
End Function

Private Function GetLeadingBytes() As String
  GetLeadingBytes = EncryptDataA(LEADING_BYTES)
End Function

Private Function CheckLeadingBytes(FileData As String) As Boolean
  CheckLeadingBytes = (DecryptDataA(Left(FileData, (Len(LEADING_BYTES) + 1))) = LEADING_BYTES)
End Function

Private Function RemoveLeadingBytes(FileData As String) As String
  If CheckLeadingBytes(FileData) Then
    RemoveLeadingBytes = Right(FileData, (Len(FileData) - (Len(LEADING_BYTES) + 1)))
  End If
End Function

Private Function WriteDataToFile(FileName As String, FileData As String) As Boolean
  Dim fI As Integer
  If DeleteFileA(FileName) Then
    On Error GoTo WriteError
    fI = FreeFile
    Open FileName For Binary Access Write As fI
    Put #fI, , FileData
    WriteDataToFile = True
  End If
WriteExit:
  Close #fI
  On Error GoTo 0
  Exit Function
WriteError:
  Resume WriteExit
End Function

Private Function ReadDataFromFile(FileName As String) As String
  Dim fI As Integer
  Dim fD As String
  On Error GoTo ReadError
  fI = FreeFile
  fD = String(GetFileSizeA(FileName), " ")
  Open FileName For Binary Access Read As fI
  Get #fI, , fD
  ReadDataFromFile = fD
ReadExit:
  Close #fI
  On Error GoTo 0
  Exit Function
ReadError:
  Resume ReadExit
End Function

Public Function GetSavePath() As String
  GetSavePath = FixPath(GetSpecialFolderPathA(CSIDL_APPDATA)) & "\Ubehage's EasyCopy2"
End Function
