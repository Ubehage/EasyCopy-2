Attribute VB_Name = "modFileSystem"
Option Explicit

Private Const LB_SETTABSTOPS As Long = &H192

Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5  'can be a CD or a DVD
Public Const DRIVE_RAMDISK As Long = 6
Public Const DRIVE_UNKNOWN = 7
Public Const DRIVE_NOTASSIGNED = 8
Public Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Public Const SHGFP_TYPE_DEFAULT = &H1

Private Const MAX_LENGTH = 260
Private Const S_OK = 0
Private Const S_FALSE = 1

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long
   
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetDriveTypeB Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Function GetSpecialFolderPathA(SpecialFolder As Long) As String
  Dim r As String
  Dim dwFlags As Long
  r = Space(MAX_LENGTH)
  If SHGetFolderPath(0&, SpecialFolder, -1, SHGFP_TYPE_CURRENT, r) = S_OK Then
    GetSpecialFolderPathA = TrimNull(r)
  End If
End Function

Public Function CreateFolderA(Path As String) As Boolean
  If Not FolderExistsA(Path) Then
    If CreateFolderA(GetParentFolderA(Path)) Then
      On Error GoTo FolderError
      MkDir Path
      CreateFolderA = True
    End If
  Else
    CreateFolderA = True
  End If
FolderExit:
  On Error GoTo 0
  Exit Function
FolderError:
  Resume FolderExit
End Function

Public Function GetRandomFileNameA() As String
  Dim fP As String
  Dim fN As String
  Dim nF As String
  fP = FixPath(Environ("temp"))
  Do
    fN = GetRandomStringA(GetRandomNumberA(5, 39)) & "." & GetRandomStringA(GetRandomNumberA(3, 7))
    nF = fP & "\" & fN
  Loop Until Not FileExistsA(nF)
  GetRandomFileNameA = nF
End Function

Public Function GetAllDrivesA() As String
  Dim i As Long
  Dim aD As String
  Dim tD As String
  Dim r As String
  aD = Space(((26 * 4) + 1))
  Call GetLogicalDriveStrings(Len(aD), aD)
  Do Until i = Len(aD)
    i = (i + 1)
    tD = Mid(aD, i, 1)
    Select Case tD
      Case ":"
        'do nothing...
      Case "\"
        'do nothing...
      Case " "
        'do nothing...
      Case vbNullChar
        'do nothing...
      Case Else
        r = r & tD
    End Select
  Loop
  GetAllDrivesA = r
End Function

Public Function GetDriveTypeA(DriveLetter As String) As Long
  Dim dT As Long
  dT = GetDriveTypeB(DriveLetter & ":\")
  If dT = 0 Then
    GetDriveTypeA = DRIVE_UNKNOWN
  ElseIf dT = 1 Then
    GetDriveTypeA = DRIVE_NOTASSIGNED
  Else
    GetDriveTypeA = dT
  End If
End Function

Public Function GetDiskVolumeA(DriveLetter As String) As String
  Dim i As Long
  Dim dV As String
  dV = String(255, vbNullChar)
  Call GetVolumeInformation(DriveLetter & ":\", dV, 255, 0, 0, 0, 0, 255)
  i = InStr(dV, vbNullChar)
  If Not i = 0 Then
    dV = Left(dV, (i - 1))
  End If
  GetDiskVolumeA = dV
End Function
