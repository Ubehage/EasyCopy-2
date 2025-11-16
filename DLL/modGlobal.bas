Attribute VB_Name = "modGlobal"
Option Explicit

Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Public Const EWX_LOGOFF As Long = &H0
Public Const EWX_SHUTDOWN As Long = &H1
Public Const EWX_REBOOT As Long = &H2
Public Const EWX_FORCE As Long = &H4
Public Const EWX_POWEROFF As Long = &H8

Private Const VER_PLATFORM_WIN32_NT As Long = 2

Public Const ICC_LISTVIEW_CLASSES As Long = &H1 'listview, header
Public Const ICC_TREEVIEW_CLASSES As Long = &H2 'treeview, tooltips
Public Const ICC_BAR_CLASSES As Long = &H4      'toolbar, statusbar, trackbar, tooltips
Public Const ICC_TAB_CLASSES As Long = &H8      'tab, tooltips
Public Const ICC_UPDOWN_CLASS As Long = &H10    'updown
Public Const ICC_PROGRESS_CLASS As Long = &H20  'progress
Public Const ICC_HOTKEY_CLASS As Long = &H40    'hotkey
Public Const ICC_ANIMATE_CLASS As Long = &H80   'animate
Public Const ICC_WIN95_CLASSES As Long = &HFF   'everything else
Public Const ICC_DATE_CLASSES As Long = &H100   'month picker, date picker, time picker, updown
Public Const ICC_USEREX_CLASSES As Long = &H200 'comboex
Public Const ICC_COOL_CLASSES As Long = &H400   'rebar (coolbar) control

'WIN32_IE >= 0x0400
Public Const ICC_INTERNET_CLASSES As Long = &H800
Public Const ICC_PAGESCROLLER_CLASS As Long = 1000 'page scroller
Public Const ICC_NATIVEFNTCTL_CLASS As Long = 2000 'native font control

'WIN32_WINNT >= 0x501
Public Const ICC_STANDARD_CLASSES As Long = 4000
Public Const ICC_LINK_CLASS As Long = 8000

Private Const FILESIZE_FIX = 4294967296#

Public Const BOOL_TRUE = "True"
Public Const BOOL_FALSE = "False"

Global Const BIF_RETURNONLYFSDIRS = &H1
Global Const BIF_DONTGOBELOWDOMAIN = &H2
Global Const BIF_STATUSTEXT = &H4
Global Const BIF_RETURNFSANCESTORS = &H8
Global Const BIF_BROWSEFORCOMPUTER = &H1000
Global Const BIF_BROWSEFORPRINTER = &H2000
Global Const MAX_PATH As Long = 260

Global Const MAXDWORD As Long = &HFFFFFFFF
Global Const INVALID_HANDLE_VALUE = -1
Global Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'system icon index
Public Const SHGFI_LARGEICON = &H0 'large icon
Public Const SHGFI_SMALLICON = &H1 'small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const ILD_TRANSPARENT = &H1 'display transparent
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
                                 SHGFI_SHELLICONSIZE Or _
                                 SHGFI_SYSICONINDEX Or _
                                 SHGFI_DISPLAYNAME Or _
                                 SHGFI_EXETYPE

Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000& 'see comments
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_READONLY As Long = &H1
Public Const OFN_SHAREAWARE As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH As Long = 2
Public Const OFN_SHAREWARN As Long = 0
Public Const OFN_SHARENOWARN As Long = 1
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFS_MAXPATHNAME As Long = 260

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS _
             Or OFN_FILEMUSTEXIST

Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Type LUID
   dwLowPart As Long
   dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   udtLUID As LUID
   dwAttributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   laa As LUID_AND_ATTRIBUTES
End Type

Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Public Type OPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Public Type tagINITCOMMONCONTROLSEX   ' icc
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be initialized.
End Type

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Sub InitOldCC Lib "comctl32" Alias "InitCommonControls" ()

Public Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function GetUserNameB Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function ScanFolderContent(Path As String, Optional TargetFolderItem As FolderItem) As FolderItem
  Dim wfd As WIN32_FIND_DATA
  Dim sP As String
  Dim hFile As Long
  Dim tF As String
  Dim fD As FILE_DATA
  Dim ScanResult As FolderItem
  If TargetFolderItem Is Nothing Then
    Set ScanResult = New FolderItem
    ScanResult.Path = Path
  Else
    Set ScanResult = TargetFolderItem
  End If
  fD.Path = FixPath(Path)
  sP = fD.Path & "\"
  hFile = FindFirstFile(sP & "*.*", wfd)
  If hFile <> INVALID_HANDLE_VALUE Then
    Do
      tF = TrimNull(wfd.cFileName)
      If (wfd.dwFileAttributes And vbDirectory) Then
        If (tF = "." Or tF = "..") Then
          'do nothing...
        Else
          ScanResult.AddSubFolder tF
        End If
      Else
        With fD
          .FileName = tF
          .FileSize = GetFileSizeFromValues(wfd.nFileSizeHigh, wfd.nFileSizeLow)
          .FileIcon = 0
          .FileAttributes = wfd.dwFileAttributes
        End With
        ScanResult.AddFile fD
      End If
    Loop While FindNextFile(hFile, wfd)
  End If
  Call FindClose(hFile)
  Set ScanFolderContent = ScanResult
End Function

Public Function FixPath(Path As String) As String
  If Right(Path, 1) = "\" Then
    FixPath = Left(Path, (Len(Path) - 1))
  Else
    FixPath = Path
  End If
End Function

Public Function TrimNull(TrimString As String) As String
  TrimNull = Left(TrimString, lstrlen(StrPtr(TrimString)))
End Function

Public Function GetFileSizeFromValues(HighSize As Long, LowSize As Long) As Double
  Dim fS As Double
  fS = LowSize
  If LowSize < 0 Then
    fS = (fS + 4294967296@)
  End If
  If HighSize > 0 Then
    fS = (fS + (HighSize * FILESIZE_FIX))
  End If
  GetFileSizeFromValues = fS
End Function

Public Function GetFileSizeA(FileName As String) As Double
  Dim wfd As WIN32_FIND_DATA
  Dim hFile As Long
  Dim fS As Currency
  hFile = FindFirstFile(FileName, wfd)
  If Not hFile = INVALID_HANDLE_VALUE Then
    fS = wfd.nFileSizeLow
    If wfd.nFileSizeLow < 0 Then
      fS = (fS + FILESIZE_FIX)
    End If
    If wfd.nFileSizeHigh > 0 Then
      fS = (fS + (wfd.nFileSizeHigh * FILESIZE_FIX))
    End If
  End If
  Call FindClose(hFile)
  GetFileSizeA = CDbl(fS)
End Function

Public Function GetIconFromFile(FileName As String, Optional LargeIcon As Boolean = True) As IPictureDisp
  Dim hImg As Long
  Dim shinfo As SHFILEINFO
  Dim IconFlags As Long
  Dim r As Long
  IconFlags = BASIC_SHGFI_FLAGS
  If LargeIcon Then
    IconFlags = IconFlags Or SHGFI_LARGEICON
  Else
    IconFlags = IconFlags Or SHGFI_SMALLICON
  End If
  hImg = SHGetFileInfo(FileName, 0&, shinfo, Len(shinfo), IconFlags)
  LoadObjectForm
  If hImg = 0 Then
    Set GetIconFromFile = frmObj.pUnknown.Picture
  Else
    frmObj.ResizePictureBox LargeIcon
    Call ImageList_Draw(hImg, shinfo.iIcon, frmObj.p.hDC, 0, 0, ILD_TRANSPARENT)
    frmObj.p.Picture = frmObj.p.Image
    Set GetIconFromFile = frmObj.p.Picture
  End If
  UnloadObjectForm
End Function

Public Function ExtractIconFromFileA(FileName As String, IconIndex As Long) As IPictureDisp
  Dim hIcon As Long
  hIcon = ExtractIcon(App.hInstance, FileName, IconIndex)
  If hIcon = 0 Then
    LoadObjectForm
    frmObj.ResizePictureBox True
    Set ExtractIconFromFileA = frmObj.pUnknown.Picture
    UnloadObjectForm
  Else
    Set ExtractIconFromFileA = GetIconFromHandleA(hIcon, True)
    Call DestroyIcon(hIcon)
  End If
End Function

Public Function GetIconFromHandleA(IconHandle As Long, Optional LargeIcon As Boolean = True) As IPictureDisp
  LoadObjectForm
  frmObj.ResizePictureBox LargeIcon
  If IconHandle = 0 Then
    Set GetIconFromHandleA = frmObj.pUnknown.Picture
  Else
    DrawIcon frmObj.p.hDC, 0, 0, IconHandle
    frmObj.p.Picture = frmObj.p.Image
    Set GetIconFromHandleA = frmObj.p.Picture
  End If
  UnloadObjectForm
End Function

Public Function GetRandomNumberA(Min As Long, Max As Long) As Long
  Dim rV As Long
  rV = ((Rnd * Max) + Min)
  If rV < Min Then
    rV = GetRandomNumberA(Min, Max)
  ElseIf rV > Max Then
    rV = GetRandomNumberA(Min, Max)
  End If
  GetRandomNumberA = rV
End Function

Public Function GetRandomCharacterA() As String
  Dim rC As String
  Select Case GetRandomNumberA(1, 36)
    Case 1
      rC = "a"
    Case 2
      rC = "b"
    Case 3
      rC = "c"
    Case 4
      rC = "d"
    Case 5
      rC = "e"
    Case 6
      rC = "f"
    Case 7
      rC = "g"
    Case 8
      rC = "h"
    Case 9
      rC = "i"
    Case 10
      rC = "j"
    Case 11
      rC = "k"
    Case 12
      rC = "l"
    Case 13
      rC = "m"
    Case 14
      rC = "n"
    Case 15
      rC = "o"
    Case 16
      rC = "p"
    Case 17
      rC = "q"
    Case 18
      rC = "r"
    Case 19
      rC = "s"
    Case 20
      rC = "t"
    Case 21
      rC = "u"
    Case 22
      rC = "v"
    Case 23
      rC = "w"
    Case 24
      rC = "x"
    Case 25
      rC = "y"
    Case 26
      rC = "z"
    Case 27
      rC = "0"
    Case 28
      rC = "1"
    Case 29
      rC = "2"
    Case 30
      rC = "3"
    Case 31
      rC = "4"
    Case 32
      rC = "5"
    Case 33
      rC = "6"
    Case 34
      rC = "7"
    Case 35
      rC = "8"
    Case 36
      rC = "9"
  End Select
  GetRandomCharacterA = rC
End Function

Public Function GetRandomStringA(StringLength As Long) As String
  Dim i As Long
  Dim rS As String
  For i = 1 To StringLength
    rS = rS & GetRandomCharacterA()
  Next
  GetRandomStringA = rS
End Function

Public Function BrowseForFolderA(Title As String, hwndOwner As Long) As String
  Dim bI As BROWSEINFO
  Dim pidl As Long
  Dim nP As String
  Dim pos As Long
  With bI
    .hOwner = hwndOwner
    .pidlRoot = 0&
    .lpszTitle = Title
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  pidl = SHBrowseForFolder(bI)
  nP = Space(MAX_PATH)
  If SHGetPathFromIDList(ByVal pidl, ByVal nP) Then
    pos = InStr(nP, Chr(0))
    nP = Left(nP, (pos - 1))
    BrowseForFolderA = nP
  End If
  Call CoTaskMemFree(pidl)
End Function

Public Function FileExistsA(FileName As String) As Boolean
  Dim wfd As WIN32_FIND_DATA
  Dim hFile As Long
  hFile = FindFirstFile(FileName, wfd)
  FileExistsA = Not hFile = INVALID_HANDLE_VALUE
  Call FindClose(hFile)
End Function

Public Function FolderExistsA(Path As String) As Boolean
  Dim wfd As WIN32_FIND_DATA
  Dim hFile As Long
  Dim sP As String
  sP = FixPath(Path)
  If Len(sP) <= 3 Then
    If InStr(UCase(GetAllDrivesA), UCase(Left(sP, 1))) Then
      FolderExistsA = True
    End If
  Else
    hFile = FindFirstFile(sP, wfd)
    If Not hFile = INVALID_HANDLE_VALUE Then
      FolderExistsA = (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
    End If
    Call FindClose(hFile)
  End If
End Function

Public Function BrowseForFileA(Title As String, hwndOwner As Long) As String
  Dim ofn As OPENFILENAME
  Dim sFilter As String
  Dim pos As Long
  Dim buff As String
  Dim sLongName As String
  Dim sShortName As String
  sFilter = "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
  With ofn
    .nStructSize = Len(ofn)
    .hwndOwner = hwndOwner
    .sFilter = sFilter
    .nFilterIndex = 0
    .sFile = Space(1024) & vbNullChar & vbNullChar
    .nMaxFile = Len(.sFile)
    .sDefFileExt = "" & vbNullChar & vbNullChar
    .sFileTitle = vbNullChar & Space(512) & vbNullChar & vbNullChar
    .nMaxTitle = Len(.sFileTitle)
    .sInitialDir = vbNullChar & vbNullChar
    .sDialogTitle = Title
    .Flags = OFS_FILE_OPEN_FLAGS Or OFN_ALLOWMULTISELECT
  End With
  If GetOpenFileName(ofn) Then
    BrowseForFileA = Trim(Left(ofn.sFile, (Len(ofn.sFile) - 2)))
  End If
End Function

Public Function GetDiskNameA(DiskLetter As String) As String
  Dim dL As String
  Dim drvN As String
  Dim pos As Integer
  Dim uV1 As Long
  Dim uV2 As Long
  Dim uV3 As Long
  Dim uV4 As String
  dL = Left(DiskLetter, 1) & ":\"
  drvN = Space(14)
  uV4 = Space(32)
  Call GetVolumeInformation(dL, drvN, Len(drvN), uV1, uV2, uV3, uV4, Len(uV4))
  pos = InStr(drvN, vbNullChar)
  If Not pos = 0 Then
    drvN = Left(drvN, (pos - 1))
  End If
  If Len(Trim(drvN)) = 0 Then
    drvN = "[No label]"
  End If
  GetDiskNameA = drvN & " (" & UCase(Left(DiskLetter, 1)) & ":" & ")"
End Function

Public Function GetStringFromBoolean(BoolValue As Boolean) As String
  If BoolValue Then
    GetStringFromBoolean = BOOL_TRUE
  Else
    GetStringFromBoolean = BOOL_FALSE
  End If
End Function

Public Function GetBooleanFromString(BoolString As String) As Boolean
  GetBooleanFromString = (LCase(BoolString) = LCase(BOOL_TRUE))
End Function

Public Function DeleteFileA(FileName As String) As Boolean
  On Error GoTo DeleteError
  SetAttr FileName, vbNormal
  Kill FileName
DeleteExit:
  On Error GoTo 0
  DeleteFileA = Not FileExistsA(FileName)
  Exit Function
DeleteError:
  Resume DeleteExit
End Function

Public Function GetByteSizeStringA(ByteValue As Double) As String
  Dim bV As Double
  Dim bN As String
  bV = ByteValue
  If bV >= 1024 Then
    bV = (bV / 1024)
    If bV >= 1024 Then
      bV = (bV / 1024)
      If bV >= 1024 Then
        bV = (bV / 1024)
        If bV >= 1024 Then
          bV = (bV / 1024)
          bN = "TBytes"
        Else
          bN = "GBytes"
        End If
      Else
        bN = "MBytes"
      End If
    Else
      bN = "KBytes"
    End If
  Else
    bN = "Bytes"
  End If
  If bV = 1 Then
    bN = Left(bN, (Len(bN) - 1))
  End If
  GetByteSizeStringA = RoundByteSizeToString(bV) & " " & bN
End Function

Private Function RoundByteSizeToString(ByteValue As Double) As String
  Dim i As Long
  Dim bV As String
  bV = Trim(Str(ByteValue))
  i = InStr(bV, ".")
  If i = 0 Then
    bV = bV & ".00"
  Else
    If i = 2 Then
      bV = Left(bV, (i + 2))
    ElseIf i = 3 Then
      bV = Left(bV, (i + 2))
    ElseIf i = 4 Then
      bV = Left(bV, (i + 1))
    ElseIf i = 5 Then
      bV = Left(bV, (i - 1))
    End If
  End If
  RoundByteSizeToString = bV
End Function

Public Function GetFileNameA(Path As String) As String
  Dim i As Long
  Dim p As String
  p = Path
  i = Len(p)
  Do Until i = 0
    If Mid(p, i, 1) = "\" Then
      p = Right(p, (Len(p) - i))
      Exit Do
    End If
    i = (i - 1)
  Loop
  GetFileNameA = p
End Function

Public Function GetFileExtensionA(Path As String)
  Dim i As Long
  i = Len(Path)
  Do Until i = 0
    If Mid(Path, i, 1) = "." Then
      GetFileExtensionA = Right(Path, (Len(Path) - i))
      Exit Do
    End If
    i = (i - 1)
  Loop
End Function

Public Function GetParentFolderA(Path As String) As String
  Dim f As String
  f = GetFileNameA(Path)
  If f = Path Then
    GetParentFolderA = ""
  Else
    GetParentFolderA = Left(Path, (Len(Path) - (Len(f) + 1)))
  End If
End Function

Public Function FixPathA(Path As String) As String
  If Right(Path, 1) = "\" Then
    FixPathA = Left(Path, (Len(Path) - 1))
  Else
    FixPathA = Path
  End If
End Function

Public Function InitCommonControlsA() As Boolean
  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo InitCCError
  With icc
    .dwSize = Len(icc)
    .dwICC = ICC_LISTVIEW_CLASSES
  End With
  InitCommonControlsA = InitCommonControlsEx(icc)
InitCCExit:
  On Error GoTo 0
  Exit Function
InitCCError:
  InitOldCC
  Resume InitCCExit
End Function

Public Function EncryptDataA(DataString As String) As String
  Dim i As Long
  Dim cC As Integer
  Dim dC As Integer
  Dim nC As Integer
  Dim eS As String
  If Not DataString = "" Then
    cC = CInt(GetRandomNumberA(35, 231))
    eS = Chr(cC)
    For i = 1 To Len(DataString)
      dC = Asc(Mid(DataString, i, 1))
      nC = (dC + cC)
      If nC > 255 Then
        nC = (nC - 256)
      End If
      eS = eS & Chr(nC)
      cC = dC
    Next
    EncryptDataA = eS
  End If
End Function

Public Function DecryptDataA(EncryptedString As String) As String
  Dim i As Long
  Dim cC As Integer
  Dim eC As Integer
  Dim nC As Integer
  Dim dS As String
  If Not EncryptedString = "" Then
    cC = Asc(Left(EncryptedString, 1))
    For i = 2 To Len(EncryptedString)
      eC = Asc(Mid(EncryptedString, i, 1))
      nC = (eC - cC)
      If nC < 0 Then
        nC = (nC + 256)
      End If
      dS = dS & Chr(nC)
      cC = nC
    Next
    DecryptDataA = dS
  End If
End Function

Public Function ExitWindowsA(ExitMode As EXITWINDOWS_Mode, ForceImmediate As Boolean) As Boolean
  Dim ewFlags As Long
  Dim r As Boolean
  Select Case ExitMode
    Case EXITWINDOWS_Mode.ewLogOff
      ewFlags = EWX_LOGOFF
    Case EXITWINDOWS_Mode.ewReboot
      ewFlags = EWX_REBOOT
    Case EXITWINDOWS_Mode.ewShutDown
      ewFlags = EWX_SHUTDOWN
    Case EXITWINDOWS_Mode.ewPowerOff
      ewFlags = EWX_POWEROFF
    Case Else
      GoTo ExitWindowsExit
  End Select
  If ForceImmediate Then
    ewFlags = ewFlags Or EWX_FORCE
  End If
  If EnableShutdownPrivileges Then
    Call ExitWindowsEx(ewFlags, 0&)
    r = True
  End If
ExitWindowsExit:
  ExitWindowsA = r
End Function

Public Function GetAssociatedFileTypeA(FileName As String) As String
  Dim hK As RegKey
  Dim rK As RegKey
  Dim sV As String
  Dim rV As String
  If FolderExistsA(FileName) Then
    rV = "Folder"
  Else
    sV = GetFileExtensionA(FileName)
    If sV = "" Then
      rV = "File"
    Else
      sV = "." & LCase(sV)
      On Error GoTo TypeError
      Set hK = RegKeyFromHKey(HKEY_CLASSES_ROOT)
      Set rK = hK.ParseKeyName(sV)
      If rK Is Nothing Then
        rV = sV & "-file"
      Else
        Set rK = hK.ParseKeyName(rK.Value)
        If rK Is Nothing Then
          rV = sV & "-file"
        Else
          rV = rK.Value
        End If
      End If
    End If
  End If
TypeExit:
  On Error GoTo 0
  Set hK = Nothing
  Set rK = Nothing
  GetAssociatedFileTypeA = rV
  Exit Function
TypeError:
  Resume Next
End Function

Public Function GetCurrentUserNameA() As String
  Dim uBuff As String * 25
  Dim ret As Long
  ret = GetUserNameB(uBuff, 25)
  GetCurrentUserNameA = Left(uBuff, (InStr(uBuff, Chr(0)) - 1))
End Function

Private Function EnableShutdownPrivileges() As Boolean
  Dim hProcessHandle As Long
  Dim hTokenHandle As Long
  Dim lpv_la As LUID
  Dim token As TOKEN_PRIVILEGES
  hProcessHandle = GetCurrentProcess()
  If Not hProcessHandle = 0 Then
    If Not OpenProcessToken(hProcessHandle, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hTokenHandle) = 0 Then
      If Not LookupPrivilegeValue(vbNullString, "SeShutdownPrivilege", lpv_la) = 0 Then
        With token
          .PrivilegeCount = 1
          .laa.udtLUID = lpv_la
          .laa.dwAttributes = SE_PRIVILEGE_ENABLED
        End With
        If Not AdjustTokenPrivileges(hTokenHandle, False, token, ByVal 0&, ByVal 0&, ByVal 0&) = 0 Then
          EnableShutdownPrivileges = True
        End If
      End If
    End If
  End If
End Function
