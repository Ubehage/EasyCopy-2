VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Path As String
Dim m_SubFolders As Long
Dim m_SubFolder() As String
Dim m_Files As Long
Dim m_File() As FILE_DATA

Public Property Get Path() As String
  Path = m_Path
End Property
Public Property Get SubFolders() As Long
  SubFolders = m_SubFolders
End Property
Public Property Get SubFolder(Index As Long) As FolderItem
  Set SubFolder = New FolderItem
  SubFolder.Path = m_Path & "\" & m_SubFolder(Index)
End Property
Public Property Get Files() As Long
  Files = m_Files
End Property
Public Property Get File(Index As Long) As FILE_DATA
  CopyMemory File, m_File(Index), LenB(m_File(Index))
  'With File
  '  .Path = m_Path
  '  .FileName = m_File(Index).FileName
  '  .FileSize = m_File(Index).FileSize
  '  .FileIcon = m_File(Index).FileIcon
  '  .FileAttributes = m_File(Index).FileAttributes
  'End With
End Property
Public Property Let Path(New_Path As String)
  If Not m_Path = New_Path Then
    m_Path = FixPath(New_Path)
    Clear
  End If
End Property

Friend Sub AddSubFolder(NewPath As String)
  If (m_SubFolders Mod 10) = 0 Then
    ReDim Preserve m_SubFolder(1 To (m_SubFolders + 10)) As String
  End If
  m_SubFolders = (m_SubFolders + 1)
  m_SubFolder(m_SubFolders) = NewPath
End Sub

Friend Sub AddFile(NewFileData As FILE_DATA)
  If (m_Files Mod 10) = 0 Then
    ReDim Preserve m_File(1 To (m_Files + 10)) As FILE_DATA
  End If
  m_Files = (m_Files + 1)
  CopyMemory m_File(m_Files), NewFileData, LenB(NewFileData)
  'With m_File(m_Files)
  '  .Path = NewFileData.Path
  '  .FileName = NewFileData.FileName
  '  .FileSize = NewFileData.FileSize
  '  .FileIcon = NewFileData.FileIcon
  '  .FileAttributes = NewFileData.FileAttributes
  'End With
End Sub

Public Sub Refresh()
  Clear
  If Not m_Path = "" Then
    PopulateClassInfo
  End If
End Sub

Public Sub Clear()
  If Not m_SubFolders = 0 Then
    m_SubFolders = 0
    Erase m_SubFolder
  End If
  If Not m_Files = 0 Then
    m_Files = 0
    Erase m_File
  End If
End Sub

Public Function CountFolderSize(DoSubFolders As Boolean) As Double
  Dim i As Long
  Dim fS As Double
  For i = 1 To m_Files
    fS = (fS + m_File(i).FileSize)
  Next
  If DoSubFolders Then
    For i = 1 To m_SubFolders
      With SubFolder(i)
        .Refresh
        fS = (fS + .CountFolderSize(DoSubFolders))
      End With
    Next
  End If
  CountFolderSize = fS
End Function

Private Sub PopulateClassInfo()
  Call ScanFolderContent(m_Path, Me)
End Sub

Private Sub Class_Terminate()
  m_Path = Empty
  Clear
End Sub
