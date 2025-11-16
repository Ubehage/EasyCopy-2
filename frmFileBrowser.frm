VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFileBrowser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ubehage's EasyCopy v2 - File Browser"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lvwFiles 
      Height          =   2115
      Left            =   3360
      TabIndex        =   1
      Top             =   765
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   3731
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView tvwFolders 
      Height          =   2535
      Left            =   480
      TabIndex        =   0
      Top             =   675
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   4471
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.ImageList imgFiles 
      Left            =   2460
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imgFolders 
      Left            =   750
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmFileBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLUMN_NAME = "Name"
Private Const COLUMN_SIZE = "Size"
Private Const COLUMN_TYPE = "Type"

Private Const KEY_DESKTOP = "%[desk]%"
Private Const KEY_USER = "%[user]%"
Private Const KEY_PICTURES = "%[pic]%"
Private Const KEY_DOCUMENTS = "%[doc]%"
Private Const KEY_VIDEOS = "%[vid]%"
Private Const KEY_DOWNLOADS = "%[down]%"
Private Const KEY_COMPUTER = "%[comp]%"

Dim StopExpanding As Boolean

Dim BrowserPath As String

Dim WithEvents NodeTimer As EasyCopy2DLL.EasyTimer
Attribute NodeTimer.VB_VarHelpID = -1

Friend Sub SetForm()
  Me.Show , frmMain
  StopExpanding = False
  BrowserPath = ""
  CreateTimer
  PrepareFolderView
  PrepareFileView
  tvwFolders.Nodes.Item(1).Selected = True
  tvwFolders_NodeClick tvwFolders.SelectedItem
  tvwFolders.SelectedItem.Expanded = True
End Sub

Private Sub CreateTimer()
  KillTimer
  Set NodeTimer = New EasyCopy2DLL.EasyTimer
  NodeTimer.Interval = 300
End Sub

Private Sub KillTimer()
  If Not NodeTimer Is Nothing Then
    NodeTimer.Enabled = False
    Set NodeTimer = Nothing
  End If
End Sub

Private Sub PrepareFolderView()
  ClearFolderView
  PrepareFolderIcons
  FillFolderView
End Sub

Private Sub PrepareFileView()
  ClearFileView
  PrepareFileIcons
  With lvwFiles
    .View = lvwReport
    With .ColumnHeaders
      Call .Add(, COLUMN_NAME, "Name")
      Call .Add(, COLUMN_SIZE, "Size")
      Call .Add(, COLUMN_TYPE, "Type")
    End With
  End With
End Sub

Private Sub ClearFolderView()
  With tvwFolders
    .Nodes.Clear
    Set .ImageList = Nothing
  End With
End Sub

Private Sub PrepareFolderIcons()
  With imgFolders
    .ListImages.Clear
    .ImageWidth = 16
    .ImageHeight = 16
    Call .ListImages.Add(, , frmMain.Icon)
  End With
  Set tvwFolders.ImageList = imgFolders
End Sub

Private Sub ClearFileView()
  With lvwFiles
    .ListItems.Clear
    .ColumnHeaders.Clear
    Set .SmallIcons = Nothing
  End With
End Sub

Private Sub PrepareFileIcons()
  With imgFiles
    .ListImages.Clear
    .ImageWidth = 16
    .ImageHeight = 16
    Call .ListImages.Add(, , frmMain.Icon)
  End With
  Set lvwFiles.SmallIcons = imgFiles
End Sub

Private Sub FillFolderView()
  AddNodeExpander tvwFolders.Nodes.Add(, , KEY_DESKTOP, "Desktop", imgFolders.ListImages.Add(, , GetAssociatedIcon(GetSpecialFolderPath(CSIDL_DESKTOP), False)).Index)
End Sub

Private Sub FillFileView()
  If Not tvwFolders.SelectedItem Is Nothing Then
    AddFileItems GetNodePath(tvwFolders.SelectedItem)
  End If
End Sub

Private Sub AddFileItems(SourcePath As String)
  Dim i As Long
  Dim d As String
  Dim t As String
  Dim fScan As EasyCopy2DLL.FolderItem
  PrepareFileView
  Select Case SourcePath
    Case KEY_DESKTOP & "\" & KEY_COMPUTER, KEY_COMPUTER
      d = GetAllDrives
      Do Until d = ""
        t = Left(d, 1)
        d = Right(d, (Len(d) - 1))
        AddFileItem t & ":", -1
      Loop
    Case KEY_DESKTOP & "\" & KEY_USER, KEY_USER
      AddFileItem KEY_PICTURES, -1
      AddFileItem KEY_DOCUMENTS, -1
      AddFileItem KEY_VIDEOS, -1
    Case Else
      If BrowserPath = KEY_DESKTOP Then
        AddFileItem KEY_USER, -1
        AddFileItem KEY_COMPUTER, -1
      End If
      Set fScan = New EasyCopy2DLL.FolderItem
      fScan.Path = GetRealPath(SourcePath)
      fScan.Refresh
      For i = 1 To fScan.SubFolders
        AddFileItem fScan.SubFolder(i).Path, -1
      Next
      For i = 1 To fScan.Files
        With fScan.File(i)
          AddFileItem fScan.Path & "\" & .FileName, .FileSize
        End With
      Next
  End Select
End Sub

Private Sub AddFileItem(SourcePath As String, SourceSize As Double)
  Dim iIcon As IPictureDisp
  With lvwFiles.ListItems.Add()
    .Key = SourcePath
    If Len(.Key) = 2 Then
      .Text = GetDiskVolume(Left(.Key, 1)) & " (" & .Key & ")"
      Set iIcon = GetAssociatedIcon(SourcePath, False)
    Else
      If .Key = KEY_USER Then
        .Text = GetCurrentUserName
        Set iIcon = GetUserIcon
      ElseIf .Key = KEY_COMPUTER Then
        .Text = "Computer"
        Set iIcon = GetComputerIcon
      ElseIf .Key = KEY_PICTURES Then
        .Text = GetPicturesName
        Set iIcon = GetPicturesIcon
      ElseIf .Key = KEY_DOCUMENTS Then
        .Text = GetDocumentsName
        Set iIcon = GetDocumentsIcon
      ElseIf .Key = KEY_VIDEOS Then
        .Text = GetVideosName
        Set iIcon = GetVideosIcon
      Else
        .Text = GetFileName(SourcePath)
        Set iIcon = GetAssociatedIcon(SourcePath, False)
      End If
    End If
    .SmallIcon = imgFiles.ListImages.Add(, , iIcon).Index
    If Not SourceSize = -1 Then
      .SubItems(1) = GetByteSizeString(SourceSize)
    End If
    If Len(.Key) = 2 Then
      .SubItems(2) = GetDiskTypeName(Left(SourcePath, 1))
    Else
      If .Key = KEY_USER Then
        .SubItems(2) = "User's Personal Folder"
      ElseIf .Key = KEY_COMPUTER Then
        .SubItems(2) = "This Computer"
      Else
        .SubItems(2) = GetAssociatedFileType(SourcePath)
      End If
    End If
  End With
End Sub

Private Sub AddNodeExpander(ThisNode As ComctlLib.Node)
  Call tvwFolders.Nodes.Add(ThisNode.Key, tvwChild, ThisNode.Key & "\exp")
End Sub

Private Sub AddSubNodes(ThisNode As ComctlLib.Node)
  Do Until ThisNode.Children = 0
    tvwFolders.Nodes.Remove ThisNode.Child.Index
  Loop
  Select Case ThisNode.Key
    Case KEY_DESKTOP
      ExpandDesktopNode ThisNode
    Case KEY_DESKTOP & "\" & KEY_COMPUTER
      ExpandComputerNode ThisNode
    Case KEY_DESKTOP & "\" & KEY_USER
      ExpandUserNode ThisNode
    Case Else
      AddSubFolderNodes ThisNode
  End Select
End Sub

Private Sub ExpandDesktopNode(ThisNode As ComctlLib.Node)
  With tvwFolders.Nodes
    AddNodeExpander .Add(ThisNode.Key, tvwChild, KEY_DESKTOP & "\" & KEY_USER, GetCurrentUserName, imgFolders.ListImages.Add(, , GetUserIcon).Index)
    AddNodeExpander .Add(ThisNode.Key, tvwChild, KEY_DESKTOP & "\" & KEY_COMPUTER, "Computer", imgFolders.ListImages.Add(, , GetComputerIcon).Index)
  End With
  AddSubFolderNodes ThisNode
End Sub

Private Sub ExpandUserNode(ThisNode As ComctlLib.Node)
  AddNodeExpander tvwFolders.Nodes.Add(ThisNode.Key, tvwChild, ThisNode.Key & "\" & KEY_PICTURES, GetPicturesName, imgFolders.ListImages.Add(, , GetPicturesIcon).Index)
  AddNodeExpander tvwFolders.Nodes.Add(ThisNode.Key, tvwChild, ThisNode.Key & "\" & KEY_DOCUMENTS, GetDocumentsName, imgFolders.ListImages.Add(, , GetDocumentsIcon).Index)
  AddNodeExpander tvwFolders.Nodes.Add(ThisNode.Key, tvwChild, ThisNode.Key & "\" & KEY_VIDEOS, GetVideosName, imgFolders.ListImages.Add(, , GetVideosIcon).Index)
End Sub

Private Sub ExpandComputerNode(ThisNode As ComctlLib.Node)
  Dim d As String
  Dim cD As String
  Dim nNode As ComctlLib.Node
  d = GetAllDrives
  Do Until d = ""
    cD = Left(d, 1)
    d = Right(d, (Len(d) - 1))
    Set nNode = tvwFolders.Nodes.Add(ThisNode.Key, tvwChild)
    With nNode
      .Key = ThisNode.Key & "\" & UCase(cD) & ":"
      .Text = GetDiskVolume(cD) & " (" & UCase(cD) & ":)"
      .Image = imgFolders.ListImages.Add(, , GetAssociatedIcon(UCase(cD) & ":", False)).Index
    End With
    AddNodeExpander nNode
  Loop
End Sub

Private Sub AddSubFolderNodes(ThisNode As ComctlLib.Node)
  Dim i As Long
  Dim nP As String
  Dim fScan As EasyCopy2DLL.FolderItem
  Dim nNode As ComctlLib.Node
  nP = GetNodePath(ThisNode)
  Set fScan = New EasyCopy2DLL.FolderItem
  fScan.Path = nP
  fScan.Refresh
  For i = 1 To fScan.SubFolders
    Set nNode = tvwFolders.Nodes.Add(ThisNode.Key, tvwChild)
    With nNode
      .Text = GetFileName(fScan.SubFolder(i).Path)
      .Key = ThisNode.Key & "\" & .Text
      .Image = imgFolders.ListImages.Add(, , GetAssociatedIcon(fScan.SubFolder(i).Path)).Index
    End With
    AddNodeExpander nNode
  Next
End Sub

Private Function GetNodePath(ThisNode As ComctlLib.Node) As String
  GetNodePath = GetRealPath(ThisNode.Key)
End Function

Private Function TrackNode(NewPath As String) As Boolean
  Dim i As Long
  Dim r As Boolean
  Dim nPath As String
  Dim nName As String
  Dim cNodes As ComctlLib.Nodes
  Dim tNodes As ComctlLib.Nodes
  Dim cNode As ComctlLib.Node
  Dim tNode As ComctlLib.Node
  nPath = NewPath
  ChangedByCode = True
  r = True
  Do Until nPath = ""
    If Not tNode Is Nothing Then
      If Not tNode.Expanded Then
        tNode.Expanded = True
      End If
      Set tNode = Nothing
    End If
    i = InStr(nPath, "\")
    If i = 0 Then
      nName = nName & "\" & nPath
      nPath = ""
    Else
      nName = nName & "\" & Left(nPath, (i - 1))
      nPath = Right(nPath, (Len(nPath) - i))
    End If
    If Left(nName, 1) = "\" Then
      nName = Right(nName, (Len(nName) - 1))
    End If
    For Each cNode In tvwFolders.Nodes
      If LCase(cNode.Key) = LCase(nName) Then
        Set tNode = cNode
        Exit For
      End If
    Next
    If tNode Is Nothing Then
      Exit Do
      r = False
    Else
      tNode.Selected = True
    End If
  Loop
  ChangedByCode = False
  TrackNode = r
End Function

Private Sub UpdateNewPath(NewPath As String)
  TrackNode NewPath
  AddFileItems NewPath
End Sub

Private Function GetDiskTypeName(DriveLetter As String)
  Select Case GetDriveType(DriveLetter)
    Case DRIVE_TYPE.dtHDD
      GetDiskTypeName = "Harddisk"
    Case DRIVE_TYPE.dtCD
      GetDiskTypeName = "Optical Disk"
    Case DRIVE_TYPE.dtRemovable
      GetDiskTypeName = "Removable Disk"
    Case DRIVE_TYPE.dtRemote
      GetDiskTypeName = "Remote Disk"
    Case DRIVE_TYPE.dtRamDisk
      GetDiskTypeName = "RAM-Disk"
    Case DRIVE_TYPE.dtNotAssigned
      GetDiskTypeName = "Not Assigned"
    Case DRIVE_TYPE.dtUnknown
      GetDiskTypeName = "Unknown"
  End Select
End Function

Private Function GetRealPath(NodePath As String) As String
  Dim rP As String
  Dim tT As String
  rP = NodePath
  If InStr(rP, KEY_USER) Then
    If Left(rP, Len(KEY_DESKTOP)) = KEY_DESKTOP Then
      rP = Right(rP, (Len(rP) - (Len(KEY_DESKTOP) + 1)))
    End If
    If Not rP = KEY_USER Then
      rP = Right(rP, (Len(rP) - (Len(KEY_USER) + 1)))
    End If
  ElseIf InStr(rP, KEY_COMPUTER) Then
    If Left(rP, Len(KEY_DESKTOP)) = KEY_DESKTOP Then
      rP = Right(rP, (Len(rP) - (Len(KEY_DESKTOP) + 1)))
    End If
    If Not rP = KEY_COMPUTER Then
      rP = Right(rP, (Len(rP) - (Len(KEY_COMPUTER) + 1)))
    End If
  End If
  rP = Replace(rP, KEY_DESKTOP, GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY))
  rP = Replace(rP, KEY_PICTURES, GetSpecialFolderPath(CSIDL_MYPICTURES))
  rP = Replace(rP, KEY_DOCUMENTS, GetSpecialFolderPath(CSIDL_PERSONAL))
  rP = Replace(rP, KEY_VIDEOS, GetSpecialFolderPath(CSIDL_MYVIDEO))
  GetRealPath = rP
End Function

Private Function GetUserIcon() As IPictureDisp
  Set GetUserIcon = GetAssociatedIcon(GetSpecialFolderPath(CSIDL_PROFILE), False)
End Function

Private Function GetPicturesIcon() As IPictureDisp
  Set GetPicturesIcon = GetAssociatedIcon(GetSpecialFolderPath(CSIDL_MYPICTURES), False)
End Function

Private Function GetDocumentsIcon() As IPictureDisp
  Set GetDocumentsIcon = GetAssociatedIcon(GetSpecialFolderPath(CSIDL_PERSONAL), False)
End Function

Private Function GetVideosIcon() As IPictureDisp
  Set GetVideosIcon = GetAssociatedIcon(GetSpecialFolderPath(CSIDL_MYVIDEO), False)
End Function

Private Function GetComputerIcon() As IPictureDisp
  Set GetComputerIcon = ExtractIconFromFile("c:\windows\system32\shell32.dll", 17)
End Function

Private Function GetPicturesName() As String
  GetPicturesName = GetFileName(GetSpecialFolderPath(CSIDL_MYPICTURES))
End Function

Private Function GetDocumentsName() As String
  GetDocumentsName = GetFileName(GetSpecialFolderPath(CSIDL_PERSONAL))
End Function

Private Function GetVideosName() As String
  GetVideosName = GetFileName(GetSpecialFolderPath(CSIDL_MYVIDEO))
End Function

Private Sub ResizeObjects()
  tvwFolders.Move 60, 60, (Me.ScaleWidth * 0.4), (Me.ScaleHeight - 120)
  lvwFiles.Move ((tvwFolders.Left + tvwFolders.Width) + 30), tvwFolders.Top, 200, tvwFolders.Height
  lvwFiles.Width = (Me.ScaleWidth - (lvwFiles.Left + tvwFolders.Left))
End Sub

Private Sub SetSortKey(NewSortKey As Integer)
  If lvwFiles.SortKey = NewSortKey Then
    If lvwFiles.SortOrder = lvwAscending Then
      lvwFiles.SortOrder = lvwDescending
    Else
      lvwFiles.SortOrder = lvwAscending
    End If
  Else
    lvwFiles.SortKey = NewSortKey
    lvwFiles.SortOrder = lvwAscending
  End If
End Sub

Private Sub Form_Load()
  BrowserFormLoaded = True
End Sub

Private Sub Form_Resize()
  ResizeObjects
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillTimer
  BrowserFormLoaded = False
  StopExpanding = True
  frmMain.FileBrowserUnloaded
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  If Not lvwFiles.Sorted Then
    lvwFiles.Sorted = True
  End If
  Select Case ColumnHeader.Key
    Case COLUMN_NAME
      SetSortKey 0
    Case COLUMN_SIZE
      SetSortKey 1
    Case COLUMN_TYPE
      SetSortKey 2
  End Select
End Sub

Private Sub lvwFiles_DblClick()
  Dim DoGo As Boolean
  If Not ChangedByCode Then
    If Not lvwFiles.SelectedItem Is Nothing Then
      Select Case lvwFiles.SelectedItem.Key
        Case KEY_USER
          DoGo = True
        Case KEY_COMPUTER
          DoGo = True
        Case KEY_PICTURES
          DoGo = True
        Case KEY_DOCUMENTS
          DoGo = True
        Case KEY_VIDEOS
          DoGo = True
        Case Else
          If FolderExists(lvwFiles.SelectedItem.Key) Then
            DoGo = True
          End If
      End Select
      If DoGo Then
        BrowserPath = BrowserPath & "\" & GetFileName(lvwFiles.SelectedItem.Key)
        UpdateNewPath BrowserPath
      End If
    End If
  End If
End Sub

Private Sub lvwFiles_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyReturn
      lvwFiles_DblClick
    Case vbKeyBack
      If Not BrowserPath = KEY_DESKTOP Then
        BrowserPath = GetParentFolder(BrowserPath)
        UpdateNewPath BrowserPath
      End If
  End Select
End Sub

Private Sub lvwFiles_OLESetData(Data As ComctlLib.DataObject, DataFormat As Integer)
  Dim i As Long
  Dim cItem As ComctlLib.ListItem
  For Each cItem In lvwFiles.ListItems
    If cItem.Selected Then
      Select Case cItem.Key
        Case KEY_USER
          'do nothing...
        Case KEY_PICTURES
          Data.Files.Add GetSpecialFolderPath(CSIDL_MYPICTURES)
          i = (i + 1)
        Case KEY_DOCUMENTS
          Data.Files.Add GetSpecialFolderPath(CSIDL_PERSONAL)
          i = (i + 1)
        Case KEY_VIDEOS
          Data.Files.Add GetSpecialFolderPath(CSIDL_MYVIDEO)
          i = (i + 1)
        Case KEY_COMPUTER
          'do nothing...
        Case Else
          Data.Files.Add cItem.Key
          i = (i + 1)
      End Select
    End If
  Next
  If i = 0 Then
    DataFormat = vbCFText
  Else
    DataFormat = vbCFFiles
  End If
End Sub

Private Sub lvwFiles_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
  Dim i As Long
  Dim cItem As ComctlLib.ListItem
  For Each cItem In lvwFiles.ListItems
    If cItem.Selected Then
      Select Case cItem.Key
        Case KEY_USER
          'do nothing...
        Case KEY_PICTURES
          Data.Files.Add GetSpecialFolderPath(CSIDL_MYPICTURES)
          i = (i + 1)
        Case KEY_DOCUMENTS
          Data.Files.Add GetSpecialFolderPath(CSIDL_PERSONAL)
          i = (i + 1)
        Case KEY_VIDEOS
          Data.Files.Add GetSpecialFolderPath(CSIDL_MYVIDEO)
          i = (i + 1)
        Case KEY_COMPUTER
          'do nothing...
        Case Else
          Data.Files.Add cItem.Key
          i = (i + 1)
      End Select
    End If
  Next
  If i = 0 Then
    AllowedEffects = vbDropEffectNone
  Else
    Data.SetData , vbCFFiles
    AllowedEffects = vbDropEffectCopy
  End If
End Sub

Private Sub NodeTimer_Timer()
  NodeTimer.Enabled = False
  FillFileView
End Sub

Private Sub tvwFolders_Collapse(ByVal Node As ComctlLib.Node)
  If Node.Selected Then
    If Not Node.Key = BrowserPath Then
      BrowserPath = Node.Key
      Node.Expanded = False
      UpdateNewPath BrowserPath
    End If
  End If
End Sub

Private Sub tvwFolders_Expand(ByVal Node As ComctlLib.Node)
  If Not StopExpanding Then
    AddSubNodes Node
  End If
End Sub

Private Sub tvwFolders_NodeClick(ByVal Node As ComctlLib.Node)
  If Not ChangedByCode Then
    If Not Node.Key = BrowserPath Then
      NodeTimer.Enabled = False
      BrowserPath = Node.Key
      NodeTimer.Enabled = True
    End If
  End If
End Sub
