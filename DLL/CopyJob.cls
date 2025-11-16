VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CopyJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MIN_KEY_LENGTH = 5
Private Const MAX_KEY_LENGTH = 31

Dim m_Name As String
Dim m_JobItems As Long
Dim m_JobItem() As JobItem
Dim m_JobSaved As Boolean
Dim m_FileName As String

Public Property Get Name() As String
  Name = m_Name
End Property
Public Property Get JobItems() As Long
  JobItems = m_JobItems
End Property
Public Property Get JobItem(Index As Long) As JobItem
  If (Index > 0 And Index <= m_JobItems) Then
    Set JobItem = m_JobItem(Index)
  End If
End Property
Public Property Get JobSaved() As Boolean
  JobSaved = m_JobSaved
End Property
Friend Property Get FileName() As String
  FileName = m_FileName
End Property
Public Property Let Name(New_Name As String)
  If Not m_Name = New_Name Then
    m_Name = New_Name
    JobHasChanged
  End If
End Property
Friend Property Let FileName(New_FileName As String)
  If Not m_FileName = New_FileName Then
    m_FileName = New_FileName
    m_JobSaved = False
  End If
End Property

Public Function AddJobItem(Optional SourcePath As String) As JobItem
  If (m_JobItems Mod 10) = 0 Then
    ReDim Preserve m_JobItem(1 To (m_JobItems + 10)) As JobItem
  End If
  m_JobItems = (m_JobItems + 1)
  Set m_JobItem(m_JobItems) = New JobItem
  With m_JobItem(m_JobItems)
    If Not SourcePath = "" Then
      .SourcePath = SourcePath
    End If
    .Key = GetNewUniqueKey
    .ItemIndex = m_JobItems
    .Parent = Me
  End With
  JobHasChanged
  Set AddJobItem = m_JobItem(m_JobItems)
End Function

Public Sub RemoveJobItem(Index As Long)
  Dim i As Long
  If (Index > 0 And Index <= m_JobItems) Then
    For i = Index To (m_JobItems - 1)
      Set m_JobItem(i) = m_JobItem((i + 1))
      m_JobItem(i).ItemIndex = i
    Next
    Set m_JobItem(m_JobItems) = Nothing
    m_JobItems = (m_JobItems - 1)
    If (m_JobItems Mod 10) = 0 Then
      If m_JobItems = 0 Then
        Erase m_JobItem
      Else
        ReDim Preserve m_JobItem(1 To m_JobItems) As JobItem
      End If
    End If
    JobHasChanged
  End If
End Sub

Public Sub RefreshJobSize()
  Dim i As Long
  For i = 1 To m_JobItems
    m_JobItem(i).RefreshItemSize
  Next
End Sub

Public Sub Clear()
  m_Name = Empty
  m_JobItems = 0
  Erase m_JobItem
  m_JobSaved = True
End Sub

Public Function GetJobItemFromKey(ThisKey As String) As JobItem
  Dim i As Long
  i = GetJobIndexFromKey(ThisKey)
  If Not i = 0 Then
    Set GetJobItemFromKey = m_JobItem(i)
  End If
End Function

Public Function GetJobIndexFromKey(ThisKey As String) As Long
  Dim i As Long
  For i = 1 To m_JobItems
    If m_JobItem(i).Key = ThisKey Then
      GetJobIndexFromKey = i
      Exit For
    End If
  Next
End Function

Public Function SaveJob() As Boolean
  If m_FileName = "" Then
    m_FileName = GetNewCopyJobFileName
  End If
  If SaveCopyJob(Me) Then
    m_JobSaved = True
    SaveJob = True
  End If
End Function

Public Function LoadJob() As Boolean
  If Not m_FileName = "" Then
    LoadJob = LoadCopyJob(Me)
  End If
End Function

Public Function CanRun() As Boolean
  Dim i As Long
  Dim r As Boolean
  r = True
  For i = 1 To m_JobItems
    With m_JobItem(i)
      If (.SourcePath = "" Or .TargetPath = "") Then
        r = False
        Exit For
      End If
    End With
  Next
  CanRun = r
End Function

Friend Sub JobHasChanged()
  m_JobSaved = False
End Sub

Friend Sub JobNotChanged()
  m_JobSaved = True
End Sub

Private Function GetNewUniqueKey() As String
  Dim nK As String
  Do
    nK = GetRandomStringA(GetRandomNumberA(MIN_KEY_LENGTH, MAX_KEY_LENGTH))
  Loop Until DoesKeyExist(nK) = False
  GetNewUniqueKey = nK
End Function

Private Function DoesKeyExist(ThisKey As String) As Boolean
  DoesKeyExist = Not GetJobItemFromKey(ThisKey) Is Nothing
End Function
