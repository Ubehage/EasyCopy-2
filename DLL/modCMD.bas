Attribute VB_Name = "modCMD"
Option Explicit

Private Const CMD_LAUNCH = "gocopythis"
Private Const CMD_FILE = "file"
Private Const CMD_SEP1 = ":"
Private Const CMD_SEP2 = ";"

Public Function LaunchCopyJobA(CMDPath As String, ThisCopyJob As CopyJob) As Boolean
  Dim fN As String
  fN = GetRandomFileNameA
  If SaveCopyJobToFile(ThisCopyJob, fN) Then
    LaunchCopyJobA = LaunchCopyJobFileA(CMDPath, fN)
  End If
End Function

Public Function LaunchCopyJobFileA(CMDPath As String, FileName As String) As Boolean
  If Not FileName = "" Then
    If FileExistsA(FileName) Then
      LaunchCopyJobFileA = ExecuteCommandLine(GetNewCommandLine(CMDPath, FileName))
    End If
  End If
End Function

Private Function GetNewCommandLine(CMDPath As String, FileName As String) As String
  GetNewCommandLine = """" & CMDPath & """ " & CMD_LAUNCH & CMD_SEP2 & CMD_FILE & CMD_SEP1 & FileName
End Function

Private Function ExecuteCommandLine(NewCommandLine As String) As Boolean
  On Error GoTo ShellError
  Debug.Print NewCommandLine
  Shell NewCommandLine, vbNormalFocus
  ExecuteCommandLine = True
ShellExit:
  On Error GoTo 0
  Exit Function
ShellError:
  Resume ShellExit
End Function
