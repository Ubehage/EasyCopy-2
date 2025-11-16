VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BUFFER_DEFAULT_HD = 1024000
Private Const BUFFER_DEFAULT_OPT = 102400
Private Const BUFFER_DEFAULT_NET = BUFFER_DEFAULT_OPT
Private Const BUFFER_DEFAULT_REM = BUFFER_DEFAULT_OPT

Private Const SETTINGS_APPNAME = "Ubehage's EasyCopy v2"
Private Const SETTINGS_SETTINGS = "Settings"
Private Const SETTINGS_BUFFER_HD = "Buffer Harddisk"
Private Const SETTINGS_BUFFER_OPT = "Buffer Optical"
Private Const SETTINGS_BUFFER_NET = "Buffer Network"
Private Const SETTINGS_BUFFER_REM = "Buffer Removable"
Private Const SETTINGS_BUFFER_TOUSE = "Buffer Usage"
Private Const SETTINGS_ALWAYSWRITE = "Always Write"

Private Const SETTINGS_BUFFERUSE_LOWEST = "Lowest"
Private Const SETTINGS_BUFFERUSE_HIGHEST = "Highest"
Private Const SETTINGS_BUFFERUSE_AVERAGE = "Average"

Private Const BCM_SETSHIELD = &H160C&

Public Enum CSIDL_VALUES
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

Public Enum DRIVE_TYPE
  dtRemovable = DRIVE_REMOVABLE
  dtHDD = DRIVE_FIXED
  dtRemote = DRIVE_REMOTE
  dtCD = DRIVE_CDROM
  dtRamDisk = DRIVE_RAMDISK
  dtUnknown = DRIVE_UNKNOWN
  dtNotAssigned = DRIVE_NOTASSIGNED
End Enum

Public Enum EXITWINDOWS_Mode
  ewLogOff = &H0
  ewShutDown = &H1
  ewReboot = &H2
  ewPowerOff = &H4
End Enum

Public Enum BUFFER_USERS
  buLowest = &H1
  byHighest = &H2
  buAverage = &H4
End Enum

Public Type FILE_DATA
  Path As String
  FileName As String
  FileSize As Double
  FileIcon As Long
  FileAttributes As VbFileAttribute
End Type

Public Type DRIVE_Information
  DriveLetter As String
  DriveType As DRIVE_TYPE
  VolumeLabel As String
End Type

Public Type BUFFER_Settings
  HardDiskBuffer As Double
  OpticalBuffer As Double
  NetworkBuffer As Double
  RemovableBuffer As Double
  BufferToUse As BUFFER_USERS
  AlwaysWrite As Boolean
End Type

Public Function GetFolderContent(Path As String) As FolderItem
  Set GetFolderContent = ScanFolderContent(Path)
End Function

Public Function GetAssociatedIcon(FileName As String, Optional LargeIcon As Boolean = True) As IPictureDisp
  Set GetAssociatedIcon = GetIconFromFile(FileName, LargeIcon)
End Function

Public Function ExtractIconFromFile(FileName As String, IconIndex As Long) As IPictureDisp
  Set ExtractIconFromFile = ExtractIconFromFileA(FileName, IconIndex)
End Function

Public Function GetIconFromHandle(IconHandle As Long, Optional LargeIcon As Boolean = True) As IPictureDisp
  Set GetIconFromHandle = GetIconFromHandleA(IconHandle, LargeIcon)
End Function

Public Sub AddUACShieldToButton(ButtonHwnd As Long)
  Call SendMessage(ButtonHwnd, BCM_SETSHIELD, 0&, 1&)
End Sub

Public Function GetRandomNumber(Min As Long, Max As Long) As Long
  GetRandomNumber = GetRandomNumberA(Min, Max)
End Function

Public Function GetRandomCharacter() As String
  GetRandomCharacter = GetRandomCharacterA
End Function

Public Function GetRandomString(StringLength As Long) As String
  GetRandomString = GetRandomStringA(StringLength)
End Function

Public Function BrowseForFolder(Title As String, hwndOwner As Long) As String
  BrowseForFolder = BrowseForFolderA(Title, hwndOwner)
End Function

Public Function FileExists(FileName As String) As Boolean
  FileExists = FileExistsA(FileName)
End Function

Public Function FolderExists(Path As String) As Boolean
  FolderExists = FolderExistsA(Path)
End Function

Public Function BrowseForFile(Title As String, hwndOwner As Long) As String
  BrowseForFile = BrowseForFileA(Title, hwndOwner)
End Function

Public Function GetDiskName(DiskLetter As String) As String
  GetDiskName = GetDiskNameA(DiskLetter)
End Function

Public Function DeleteFile(FileName As String) As Boolean
  DeleteFile = DeleteFileA(FileName)
End Function

Public Function GetByteSizeString(ByteValue As Double) As String
  GetByteSizeString = GetByteSizeStringA(ByteValue)
End Function

Public Function GetFileName(Path As String) As String
  GetFileName = GetFileNameA(Path)
End Function

Public Function GetFileExtension(FileName As String) As String
  GetFileExtension = GetFileExtensionA(FileName)
End Function

Public Function GetParentFolder(Path As String) As String
  GetParentFolder = GetParentFolderA(Path)
End Function

Public Function FixPath(Path As String) As String
  FixPath = FixPathA(Path)
End Function

Public Function InitCommonControls() As Boolean
  InitCommonControls = InitCommonControlsA
End Function

Public Function EncryptData(DataString As String) As String
  EncryptData = EncryptDataA(DataString)
End Function

Public Function DecryptData(EncryptedString As String) As String
  DecryptData = DecryptDataA(EncryptedString)
End Function

Public Function GetSpecialFolderPath(SpecialFolder As CSIDL_VALUES) As String
  GetSpecialFolderPath = GetSpecialFolderPathA(SpecialFolder)
End Function

Public Function LaunchCopyJob(CMDPath As String, ThisCopyJob As CopyJob) As Boolean
  LaunchCopyJob = LaunchCopyJobA(CMDPath, ThisCopyJob)
End Function

Public Function LaunchCopyJobFile(CMDPath As String, FileName As String) As Boolean
  LaunchCopyJobFile = LaunchCopyJobFileA(CMDPath, FileName)
End Function

Public Function GetRandomFileName() As String
  GetRandomFileName = GetRandomFileNameA
End Function

Public Function GetFileSize(FileName As String) As Double
  GetFileSize = GetFileSizeA(FileName)
End Function

Public Function LoadCopyJobFromFile(FileName As String) As CopyJob
  Set LoadCopyJobFromFile = LoadCopyJobFromFileA(FileName)
End Function

Public Function CreateFolder(Path As String) As Boolean
  CreateFolder = CreateFolderA(Path)
End Function

Public Function ExitWindows(ExitMode As EXITWINDOWS_Mode, Optional ForceImmediate As Boolean) As Boolean
  ExitWindows = ExitWindowsA(ExitMode, ForceImmediate)
End Function

Public Function GetAllDrives() As String
  GetAllDrives = GetAllDrivesA
End Function

Public Function GetDriveType(DriveLetter As String) As DRIVE_TYPE
  GetDriveType = GetDriveTypeA(Left(DriveLetter, 1))
End Function

Public Function GetDiskVolume(DriveLetter As String) As String
  GetDiskVolume = GetDiskVolumeA(DriveLetter)
End Function

Public Function GetAssociatedFileType(FileName As String) As String
  GetAssociatedFileType = GetAssociatedFileTypeA(FileName)
End Function

Public Function GetCurrentUserName() As String
  GetCurrentUserName = GetCurrentUserNameA
End Function

Public Sub ReadBufferSettings(BufferSettings As BUFFER_Settings)
  With BufferSettings
    .HardDiskBuffer = ReadSingleBufferSetting(1)
    .OpticalBuffer = ReadSingleBufferSetting(2)
    .NetworkBuffer = ReadSingleBufferSetting(3)
    .RemovableBuffer = ReadSingleBufferSetting(4)
    .BufferToUse = ReadBufferToUse()
    .AlwaysWrite = ReadAlwaysWrite()
  End With
End Sub

Private Function ReadSingleBufferSetting(BufferIndex As Single) As Double
  Dim k As String
  Dim v As String
  Dim d As Long
  Select Case BufferIndex
    Case 1 'HardDisk
      k = SETTINGS_BUFFER_HD
      d = BUFFER_DEFAULT_HD
    Case 2 'Optical
      k = SETTINGS_BUFFER_OPT
      d = BUFFER_DEFAULT_OPT
    Case 3 'Network
      k = SETTINGS_BUFFER_NET
      d = BUFFER_DEFAULT_NET
    Case 4 'Removable
      k = SETTINGS_BUFFER_REM
      d = BUFFER_DEFAULT_REM
  End Select
  On Error GoTo BufferError
  v = GetSetting(SETTINGS_APPNAME, SETTINGS_SETTINGS, k)
BufferExit:
  On Error GoTo 0
  If v = "" Then
    v = Trim(Str(d))
  End If
  ReadSingleBufferSetting = CDbl(Val(v))
  Exit Function
BufferError:
  Resume BufferExit
End Function

Private Function ReadBufferToUse() As BUFFER_USERS
  Dim v As String
  On Error GoTo ReadBufferUseError
  v = GetSetting(SETTINGS_APPNAME, SETTINGS_SETTINGS, SETTINGS_BUFFER_TOUSE)
ReadBufferUseExit:
  On Error GoTo 0
  Select Case v
    Case SETTINGS_BUFFERUSE_LOWEST
      ReadBufferToUse = buLowest
    Case SETTINGS_BUFFERUSE_HIGHEST
      ReadBufferToUse = byHighest
    Case SETTINGS_BUFFERUSE_AVERAGE
      ReadBufferToUse = buAverage
    Case Else
      ReadBufferToUse = buLowest
  End Select
  Exit Function
ReadBufferUseError:
  Resume ReadBufferUseExit
End Function

Private Function ReadAlwaysWrite() As Boolean
  Dim v As String
  On Error GoTo ReadWriteError
  v = GetSetting(SETTINGS_APPNAME, SETTINGS_SETTINGS, SETTINGS_ALWAYSWRITE)
ReadWriteExit:
  On Error GoTo 0
  If v = "" Then
    v = BOOL_TRUE
  End If
  ReadAlwaysWrite = GetBooleanFromString(v)
  Exit Function
ReadWriteError:
  Resume ReadWriteExit
End Function

Public Sub SaveBufferSettings(BufferSettings As BUFFER_Settings)
  With BufferSettings
    SaveSingleBufferSetting 1, .HardDiskBuffer
    SaveSingleBufferSetting 2, .OpticalBuffer
    SaveSingleBufferSetting 3, .NetworkBuffer
    SaveSingleBufferSetting 4, .RemovableBuffer
    SaveBufferToUse .BufferToUse
    SaveAlwaysWrite .AlwaysWrite
  End With
End Sub

Private Sub SaveSingleBufferSetting(BufferIndex As Single, NewValue As Double)
  Dim k As String
  Dim v As String
  Dim d As String
  Select Case BufferIndex
    Case 1 'HardDisk
      k = SETTINGS_BUFFER_HD
      d = BUFFER_DEFAULT_HD
    Case 2 'Optical
      k = SETTINGS_BUFFER_OPT
      d = BUFFER_DEFAULT_OPT
    Case 3 'Network
      k = SETTINGS_BUFFER_NET
      d = BUFFER_DEFAULT_NET
    Case 4 'Removable
      k = SETTINGS_BUFFER_REM
      d = BUFFER_DEFAULT_REM
  End Select
  If NewValue = 0 Then
    v = Trim(Str(d))
  Else
    v = Trim(Str(NewValue))
  End If
  On Error GoTo BufferError2
  SaveSetting SETTINGS_APPNAME, SETTINGS_SETTINGS, k, v
BufferExit2:
  On Error GoTo 0
  Exit Sub
BufferError2:
  Resume BufferExit2
End Sub

Private Sub SaveBufferToUse(NewSetting As BUFFER_USERS)
  Dim v As String
  Select Case NewSetting
    Case BUFFER_USERS.buLowest
      v = SETTINGS_BUFFERUSE_LOWEST
    Case BUFFER_USERS.byHighest
      v = SETTINGS_BUFFERUSE_HIGHEST
    Case BUFFER_USERS.buAverage
      v = SETTINGS_BUFFERUSE_AVERAGE
  End Select
  On Error GoTo SaveBufferUseError
  SaveSetting SETTINGS_APPNAME, SETTINGS_SETTINGS, SETTINGS_BUFFER_TOUSE, v
SaveBufferUseExit:
  On Error GoTo 0
  Exit Sub
SaveBufferUseError:
  Resume SaveBufferUseExit
End Sub

Private Sub SaveAlwaysWrite(NewValue As Boolean)
  On Error GoTo SaveWriteError
  SaveSetting SETTINGS_APPNAME, SETTINGS_SETTINGS, SETTINGS_ALWAYSWRITE, GetStringFromBoolean(NewValue)
SaveWriteExit:
  On Error GoTo 0
  Exit Sub
SaveWriteError:
  Resume SaveWriteExit
End Sub
