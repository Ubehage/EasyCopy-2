VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JobItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Parent As CopyJob

Dim m_SourcePath As String
Dim m_TargetPath As String
Dim m_IsFolder As Boolean
Dim m_Size As Double
Dim m_IncludeSubFolders As Boolean
Dim m_ResetAttributes As Boolean
Dim m_Overwrite As Boolean
Dim m_IgnoreErrors As Boolean
Dim m_DeleteAfterCopy As Boolean
Dim m_DeleteAfterError As Boolean
Dim m_Key As String
Dim m_ItemIndex As Long
Dim m_ListItem As ComctlLib.ListItem
Dim m_ListViewParent As ComctlLib.ListView

Dim FoldersToScan As Long
Dim FolderToScan() As String

Dim HasListView As Boolean

Dim WithEvents ScanTimer As EasyTimer
Attribute ScanTimer.VB_VarHelpID = -1

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Friend Property Get Parent() As CopyJob
  Set Parent = m_Parent
End Property
Public Property Get SourcePath() As String
  SourcePath = m_SourcePath
End Property
Public Property Get TargetPath() As String
  TargetPath = m_TargetPath
End Property
Public Property Get IsFolder() As Boolean
  IsFolder = m_IsFolder
End Property
Public Property Get Size() As Double
  Size = m_Size
End Property
Public Property Get IncludeSubFolders() As Boolean
  IncludeSubFolders = m_IncludeSubFolders
End Property
Public Property Get ResetAttributes() As Boolean
  ResetAttributes = m_ResetAttributes
End Property
Public Property Get Overwrite() As Boolean
  Overwrite = m_Overwrite
End Property
Public Property Get IgnoreErrors() As Boolean
  IgnoreErrors = m_IgnoreErrors
End Property
Public Property Get DeleteAfterCopy() As Boolean
  DeleteAfterCopy = m_DeleteAfterCopy
End Property
Public Property Get DeleteAfterError() As Boolean
  DeleteAfterError = m_DeleteAfterError
End Property
Public Property Get Key() As String
  Key = m_Key
End Property
Public Property Get ItemIndex() As Long
  ItemIndex = m_ItemIndex
End Property
Public Property Get ListItem() As ComctlLib.ListItem
  Set ListItem = m_ListItem
End Property
Public Property Get ListViewParent() As Object
  Set ListViewParent = m_ListViewParent
End Property
Public Property Get Flags() As String
  Flags = GetFlags
End Property
Friend Property Let Parent(New_Parent As CopyJob)
  Set m_Parent = New_Parent
End Property
Friend Property Let SourcePath(New_SourcePath As String)
  If Not m_SourcePath = New_SourcePath Then
    m_SourcePath = New_SourcePath
    m_IsFolder = FolderExistsA(m_SourcePath)
  End If
End Property
Public Property Let TargetPath(New_TargetPath As String)
  If Not m_TargetPath = New_TargetPath Then
    m_TargetPath = New_TargetPath
    JobHasChanged
  End If
End Property
Public Property Let IncludeSubFolders(New_IncludeSubFolders As Boolean)
  If Not m_IncludeSubFolders = New_IncludeSubFolders Then
    m_IncludeSubFolders = New_IncludeSubFolders
    JobHasChanged
  End If
End Property
Public Property Let ResetAttributes(New_ResetAttributes As Boolean)
  If Not m_ResetAttributes = New_ResetAttributes Then
    m_ResetAttributes = New_ResetAttributes
    JobHasChanged
  End If
End Property
Public Property Let Overwrite(New_Overwrite As Boolean)
  If Not m_Overwrite = New_Overwrite Then
    m_Overwrite = New_Overwrite
    JobHasChanged
  End If
End Property
Public Property Let IgnoreErrors(New_IgnoreErrors As Boolean)
  If Not m_IgnoreErrors = New_IgnoreErrors Then
    m_IgnoreErrors = New_IgnoreErrors
    JobHasChanged
  End If
End Property
Public Property Let DeleteAfterCopy(New_DeleteAfterCopy As Boolean)
  If Not m_DeleteAfterCopy = New_DeleteAfterCopy Then
    m_DeleteAfterCopy = New_DeleteAfterCopy
    JobHasChanged
  End If
End Property
Public Property Let DeleteAfterError(New_DeleteAfterError As Boolean)
  If Not m_DeleteAfterError = New_DeleteAfterError Then
    m_DeleteAfterError = New_DeleteAfterError
    JobHasChanged
  End If
End Property
Friend Property Let Key(New_Key As String)
  If Not m_Key = New_Key Then
    m_Key = New_Key
  End If
End Property
Friend Property Let ItemIndex(New_ItemIndex As Long)
  m_ItemIndex = New_ItemIndex
End Property
Public Property Let ListItem(New_ListItem As ComctlLib.ListItem)
  If Not m_ListItem Is New_ListItem Then
    Set m_ListItem = New_ListItem
    CheckForItemSize
  End If
End Property
Public Property Let ListViewParent(New_ListViewParent As Object)
  If Not m_ListViewParent Is New_ListViewParent Then
    Set m_ListViewParent = New_ListViewParent
    CheckForItemSize
  End If
End Property

Public Sub CountItemSizeForGUI()
  If Not m_IsFolder Then
    m_Size = GetFileSizeA(m_SourcePath)
    UpdateListItem
  Else
    m_Size = 0
    FoldersToScan = 0
    Erase FolderToScan
    AddFolderToScan m_SourcePath
    StartScanTimer
  End If
End Sub

Friend Sub RefreshItemSize()
  CountItemSizeForGUI
End Sub

Private Sub CheckForItemSize()
  HasListView = False
  If Not m_ListItem Is Nothing Then
    If Not m_ListViewParent Is Nothing Then
      HasListView = True
      If m_Size = 0 Then
        CountItemSizeForGUI
      Else
        UpdateListItem
      End If
    End If
  End If
End Sub

Private Sub AddFolderToScan(NewPath As String)
  If (FoldersToScan Mod 10) = 0 Then
    ReDim Preserve FolderToScan(1 To (FoldersToScan + 10)) As String
  End If
  FoldersToScan = (FoldersToScan + 1)
  FolderToScan(FoldersToScan) = NewPath
End Sub

Private Sub ScanNextFolder()
  Dim i As Long
  Dim fI As FolderItem
  If Not FoldersToScan = 0 Then
    Set fI = New FolderItem
    fI.Path = FolderToScan(1)
    For i = 1 To (FoldersToScan - 1)
      FolderToScan(i) = FolderToScan((i + 1))
    Next
    FolderToScan(FoldersToScan) = Empty
    FoldersToScan = (FoldersToScan - 1)
    If (FoldersToScan Mod 10) = 0 Then
      If FoldersToScan = 0 Then
        Erase FolderToScan
      Else
        ReDim Preserve FolderToScan(1 To FoldersToScan) As String
      End If
    End If
    ScanThisFolder fI
    Set fI = Nothing
  End If
End Sub

Private Sub ScanThisFolder(ThisFolder As FolderItem)
  Dim i As Long
  ThisFolder.Refresh
  m_Size = (m_Size + ThisFolder.CountFolderSize(False))
  For i = 1 To ThisFolder.SubFolders
    AddFolderToScan ThisFolder.SubFolder(i).Path
  Next
End Sub

Private Sub StartScanTimer()
  Set ScanTimer = New EasyTimer
  With ScanTimer
    .Interval = 1
    .Enabled = True
  End With
End Sub

Private Sub KillScanTimer()
  If Not ScanTimer Is Nothing Then
    ScanTimer.Enabled = False
    Set ScanTimer = Nothing
  End If
End Sub

Private Function GetFlags() As String
  Dim cF As String
  If m_IncludeSubFolders Then
    cF = cF & "S"
  End If
  If m_ResetAttributes Then
    cF = cF & "A"
  End If
  If m_Overwrite Then
    cF = cF & "O"
  End If
  If m_IgnoreErrors Then
    cF = cF & "I"
  End If
  If m_DeleteAfterCopy Then
    cF = cF & "D"
  End If
  If m_DeleteAfterError Then
    cF = cF & "E"
  End If
  GetFlags = cF
End Function

Private Sub JobHasChanged()
  If Not m_Parent Is Nothing Then
    m_Parent.JobHasChanged
  End If
  UpdateListItem
End Sub

Private Sub UpdateListItem()
  On Error GoTo UpdateListError
  If HasListView Then
    LockWindowUpdate m_ListViewParent.hWnd
    m_ListItem.SubItems(2) = GetByteSizeStringA(m_Size)
    m_ListItem.SubItems(3) = GetFlags
    m_ListItem.SubItems(4) = m_TargetPath
    LockWindowUpdate 0&
  End If
ExitNow:
  On Error GoTo 0
  Exit Sub
UpdateListError:
  Resume ExitNow
End Sub

Private Sub Clear()
  KillScanTimer
  Set m_Parent = Nothing
  m_SourcePath = Empty
  m_TargetPath = Empty
  m_IncludeSubFolders = False
  m_ResetAttributes = False
  m_Overwrite = False
  m_IgnoreErrors = False
  m_DeleteAfterCopy = False
  m_DeleteAfterError = False
End Sub

Private Sub Class_Terminate()
  Clear
End Sub

Private Sub ScanTimer_Timer()
  ScanTimer.Enabled = False
  ScanNextFolder
  UpdateListItem
  If Not FoldersToScan = 0 Then
    ScanTimer.Enabled = True
  Else
    KillScanTimer
  End If
End Sub
