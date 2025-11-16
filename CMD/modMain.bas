Attribute VB_Name = "modMain"
Option Explicit

Global Const LABEL_BYTES = "Total progress - bytes copied: %b% of %t%"
Global Const LABEL_FILES = "Files copied: %f% of %t%"
Global Const LABEL_FILE = "Current file: %b% of %t%"

Private Const CMD_LAUNCH = "gocopythis"
Private Const CMD_FILE = "file"
Private Const CMD_SEP1 = ":"
Private Const CMD_SEP2 = ";"

Dim DoLaunch As Boolean
Dim LaunchFile As String

Global CurrentCopyJob As CopyJob

Global ExitNow As Boolean
Global UnloadedByCode As Boolean

Global BufferSettings As BUFFER_Settings

Sub Main()
  Dim DoShutdown As Boolean
  InitCommonControls
  SplitCommandLine Command
  
  
  'DoLaunch = True
  'LaunchFile = "C:\Users\Ubehage\AppData\Roaming\Ubehage's EasyCopy2\2ok7xxpeuz74.ecj"
  
  
  If Not DoLaunch Then
    MsgBox "This application can be launched from the main GUI!", vbOKOnly Or vbInformation, "No support"
  Else
    Set CurrentCopyJob = LoadCopyJobFromFile(LaunchFile)
    If CurrentCopyJob Is Nothing Then
      MsgBox "The selected file could not be loaded!" & vbCrLf & vbCrLf & "The file might be corrupted, or it is not an EasyCopy2 Job-file!", vbOKOnly Or vbInformation, "Cannot continue"
    Else
      LoadForm
      CountAllFilesAndFolders
      If Not ExitNow Then
        ReadBufferSettings BufferSettings
        StartCopying
      End If
      If frmMain.chkShutdown.Value = vbChecked Then
        DoShutdown = True
      End If
      UnloadForm
    End If
  End If
  If Not ExitNow Then
    If DoShutdown Then
      Call ExitWindows(ewShutDown)
    End If
  End If
End Sub

Private Sub LoadForm()
  Load frmMain
  frmMain.SetForm
End Sub

Private Sub UnloadForm()
  UnloadedByCode = True
  Unload frmMain
  UnloadedByCode = False
  Set frmMain = Nothing
End Sub

Private Sub SplitCommandLine(CommandLine As String)
  Dim i As Long
  Dim cL As String
  Dim cC As String
  Dim cV As String
  cL = CommandLine
  Do Until cL = ""
    i = InStr(cL, CMD_SEP2)
    If i = 0 Then
      cC = cL
      cL = ""
    Else
      cC = Left(cL, (i - 1))
      cL = Right(cL, (Len(cL) - i))
    End If
    i = InStr(cC, CMD_SEP1)
    If i = 0 Then
      cV = ""
    Else
      cV = Right(cC, (Len(cC) - i))
      cC = Left(cC, (i - 1))
    End If
    Select Case cC
      Case CMD_LAUNCH
        DoLaunch = True
      Case CMD_FILE
        LaunchFile = cV
    End Select
  Loop
End Sub

Public Function MakeFolder(Path As String) As Boolean
  Dim r As Boolean
  Dim p As String
  Dim m As String
  If Not Path = "" Then
    r = True
    p = FixPath(Path)
CreatingFolder:
    If Not FolderExists(p) Then
      If MakeFolder(GetParentFolder(p)) Then
        On Error GoTo MakeFolderError
        MkDir p
      Else
DiskNotReady:
        If Len(Path) <= 3 Then
          m = "The selected target path is not accessible!" & vbCrLf & vbCrLf
          m = m & "Please insert a disk into drive " & UCase(Left(p, 2)) & " and try again."
          Select Case MsgBox(m, vbAbortRetryIgnore Or vbCritical, "Error!")
            Case vbAbort
              If CanAbort Then
                ExitNow = True
                r = False
              Else
                GoTo DiskNotReady
              End If
            Case vbRetry
              GoTo CreatingFolder
            Case vbIgnore
              GoTo MakeFolderExit
            Case Else
              GoTo DiskNotReady
          End Select
        End If
      End If
    End If
  End If
MakeFolderExit:
  On Error GoTo 0
  MakeFolder = r
  Exit Function
MakeFolderError:
  m = "Could not create the folder """ & p & """!"
  m = m & vbCrLf & vbCrLf & "Error Message: " & Error
  m = m & vbCrLf & "Error Code: " & Trim(Str(Err))
  Select Case MsgBox(m, vbAbortRetryIgnore Or vbCritical, "Error!")
    Case vbAbort
      If CanAbort Then
        ExitNow = True
        r = False
        Resume MakeFolderExit
      Else
        GoTo MakeFolderError
      End If
    Case vbRetry
      Resume
    Case vbIgnore
      Resume MakeFolderExit
    Case Else
      GoTo MakeFolderError
  End Select
End Function
