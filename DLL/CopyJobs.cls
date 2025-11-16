VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CopyJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_CopyJobs As Long
Dim m_CopyJob() As CopyJob

Public Property Get CopyJobs() As Long
  CopyJobs = m_CopyJobs
End Property
Public Property Get CopyJob(Index As Long) As CopyJob
  If (Index > 0 And Index <= m_CopyJobs) Then
    Set CopyJob = m_CopyJob(Index)
  End If
End Property

Public Function AddCopyJob(Optional NewJobName As String) As CopyJob
  If (m_CopyJobs Mod 10) = 0 Then
    ReDim Preserve m_CopyJob(1 To (m_CopyJobs + 10)) As CopyJob
  End If
  m_CopyJobs = (m_CopyJobs + 1)
  Set m_CopyJob(m_CopyJobs) = New CopyJob
  If Not NewJobName = "" Then
    m_CopyJob(m_CopyJobs).Name = NewJobName
  End If
  'm_CopyJob(m_CopyJobs).JobNotChanged
  m_CopyJob(m_CopyJobs).JobHasChanged
  Set AddCopyJob = m_CopyJob(m_CopyJobs)
End Function

Public Function RemoveCopyJob(Index As Long) As Boolean
  Dim i As Long
  If (Index > 0 And Index <= m_CopyJobs) Then
    RemoveCopyJob = DeleteFileA(m_CopyJob(Index).FileName)
    For i = Index To (m_CopyJobs - 1)
      Set m_CopyJob(i) = m_CopyJob((i + 1))
    Next
    Set m_CopyJob(m_CopyJobs) = Nothing
    m_CopyJobs = (m_CopyJobs - 1)
    If (m_CopyJobs Mod 10) = 0 Then
      If m_CopyJobs = 0 Then
        Erase m_CopyJob
      Else
        ReDim Preserve m_CopyJob(1 To m_CopyJobs) As CopyJob
      End If
    End If
  End If
End Function

Public Sub LoadAllCopyJobs()
  Dim i As Long
  Clear
  With ScanFolderContent(GetSavePath)
    For i = 1 To .Files
      If LCase(Right(.File(i).FileName, Len(FILE_EXTENSION))) = LCase(FILE_EXTENSION) Then
        AddCopyJob().FileName = .Path & "\" & .File(i).FileName
        m_CopyJob(m_CopyJobs).LoadJob
      End If
    Next
  End With
End Sub

Public Function SaveAllCopyJobs() As Boolean
  Dim i As Long
  Dim r As Boolean
  r = True
  For i = 1 To m_CopyJobs
    If Not m_CopyJob(i).SaveJob Then
      r = False
    End If
  Next
  SaveAllCopyJobs = r
End Function

Public Sub Clear()
  m_CopyJobs = 0
  Erase m_CopyJob
End Sub

Private Sub Class_Terminate()
  Clear
End Sub
