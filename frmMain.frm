VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Ubehage's EasyCopy v2"
   ClientHeight    =   7470
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "Advanced Options..."
      Height          =   420
      Left            =   4200
      TabIndex        =   27
      Top             =   6600
      Width           =   1665
   End
   Begin VB.Frame fJob 
      Caption         =   "Current Copy Job"
      Height          =   7215
      Left            =   3900
      TabIndex        =   5
      Top             =   150
      Width           =   6495
      Begin VB.PictureBox pJob 
         Height          =   6765
         Left            =   150
         ScaleHeight     =   6705
         ScaleWidth      =   6075
         TabIndex        =   6
         Top             =   225
         Width           =   6135
         Begin VB.CommandButton cmdGo 
            Caption         =   "Start Copying!"
            Height          =   420
            Left            =   4380
            TabIndex        =   22
            Top             =   6195
            Width           =   1590
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Current Job"
            Height          =   420
            Left            =   2940
            TabIndex        =   21
            Top             =   6195
            Width           =   1440
         End
         Begin VB.Frame fItems 
            Height          =   4950
            Left            =   120
            TabIndex        =   10
            Top             =   1170
            Width           =   5850
            Begin VB.PictureBox pItems 
               Height          =   4620
               Left            =   120
               ScaleHeight     =   4560
               ScaleWidth      =   5460
               TabIndex        =   11
               Top             =   255
               Width           =   5520
               Begin VB.Frame fOptions 
                  Caption         =   "Options"
                  Height          =   2355
                  Left            =   105
                  TabIndex        =   13
                  Top             =   2160
                  Width           =   5265
                  Begin VB.PictureBox pOptions 
                     Height          =   1980
                     Left            =   120
                     ScaleHeight     =   1920
                     ScaleWidth      =   4905
                     TabIndex        =   14
                     Top             =   270
                     Width           =   4965
                     Begin VB.Frame fTarget 
                        Caption         =   "Target Path"
                        Height          =   915
                        Left            =   135
                        TabIndex        =   23
                        Top             =   60
                        Width           =   4680
                        Begin VB.PictureBox pTarget 
                           Height          =   540
                           Left            =   180
                           ScaleHeight     =   480
                           ScaleWidth      =   4365
                           TabIndex        =   24
                           Top             =   285
                           Width           =   4425
                           Begin VB.CommandButton cmdTarget 
                              Caption         =   "..."
                              Height          =   330
                              Left            =   3990
                              TabIndex        =   26
                              Top             =   45
                              Width           =   330
                           End
                           Begin VB.TextBox txtTarget 
                              Height          =   330
                              Left            =   75
                              TabIndex        =   25
                              Top             =   75
                              Width           =   3825
                           End
                        End
                     End
                     Begin VB.CheckBox chkDelError 
                        Caption         =   "Even after errors"
                        Height          =   225
                        Left            =   3225
                        TabIndex        =   20
                        Top             =   1635
                        Width           =   1500
                     End
                     Begin VB.CheckBox chkDelete 
                        Caption         =   "Delete after successful copy"
                        Height          =   225
                        Left            =   2460
                        TabIndex        =   19
                        Top             =   1365
                        Width           =   2340
                     End
                     Begin VB.CheckBox chkError 
                        Caption         =   "Ignore All Errors"
                        Height          =   225
                        Left            =   3300
                        TabIndex        =   18
                        Top             =   1110
                        Width           =   1440
                     End
                     Begin VB.CheckBox chkOverwrite 
                        Caption         =   "Overwrite existing files without asking"
                        Height          =   225
                        Left            =   120
                        TabIndex        =   17
                        Top             =   1635
                        Width           =   2955
                     End
                     Begin VB.CheckBox chkAttributes 
                        Caption         =   "Reset Attributes"
                        Height          =   225
                        Left            =   150
                        TabIndex        =   16
                        Top             =   1350
                        Width           =   1455
                     End
                     Begin VB.CheckBox chkSubFolders 
                        Caption         =   "Include Sub-Folders"
                        Height          =   225
                        Left            =   135
                        TabIndex        =   15
                        Top             =   1065
                        Width           =   1740
                     End
                  End
               End
               Begin ComctlLib.ListView lvwItems 
                  Height          =   1995
                  Left            =   105
                  TabIndex        =   12
                  Top             =   105
                  Width           =   4875
                  _ExtentX        =   8599
                  _ExtentY        =   3519
                  MultiSelect     =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  OLEDropMode     =   1
                  _Version        =   327682
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  OLEDropMode     =   1
                  NumItems        =   0
               End
            End
         End
         Begin VB.Frame fName 
            Caption         =   "Job Name"
            Height          =   1020
            Left            =   135
            TabIndex        =   7
            Top             =   90
            Width           =   5145
            Begin VB.PictureBox pName 
               Height          =   675
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   4515
               TabIndex        =   8
               Top             =   240
               Width           =   4575
               Begin VB.TextBox txtName 
                  Height          =   330
                  Left            =   105
                  TabIndex        =   9
                  Top             =   195
                  Width           =   2175
               End
            End
         End
      End
   End
   Begin VB.Frame fJobs 
      Caption         =   "Copy Jobs"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3660
      Begin VB.PictureBox pJobs 
         Height          =   5205
         Left            =   90
         ScaleHeight     =   5145
         ScaleWidth      =   3360
         TabIndex        =   1
         Top             =   300
         Width           =   3420
         Begin VB.CommandButton cmdDelJob 
            Caption         =   "Delete Selected Job"
            Height          =   420
            Left            =   1575
            TabIndex        =   4
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CommandButton cmdNewJob 
            Caption         =   "New Copy Job..."
            Height          =   420
            Left            =   120
            TabIndex        =   3
            Top             =   2820
            Width           =   1365
         End
         Begin VB.ListBox lstJobs 
            Height          =   2595
            Left            =   180
            TabIndex        =   2
            Top             =   150
            Width           =   2595
         End
      End
   End
   Begin ComctlLib.ImageList imlItems 
      Left            =   405
      Top             =   6315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuMainMenu 
      Caption         =   "&Main Menu"
      Begin VB.Menu mnuNewCopyJob 
         Caption         =   "&Create new Copy Job..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDeleteCopyJob 
         Caption         =   "&Delete selected Copy Job"
      End
      Begin VB.Menu mnuMainSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBrowser 
         Caption         =   "&Show File Browser..."
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuMainSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuJobs 
      Caption         =   "JobsList"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "New Copy Job..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Selected Job"
      End
   End
   Begin VB.Menu mnuItems 
      Caption         =   "JobItems"
      Visible         =   0   'False
      Begin VB.Menu mnuAddFolder 
         Caption         =   "Add Folder..."
      End
      Begin VB.Menu mnuAddFiles 
         Caption         =   "Add File(s)..."
      End
      Begin VB.Menu mnuItemsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Selected Items"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_CAPTION = "Ubehage's EasyCopy v2 (%ver%)"

Private Const CMD_EXENAME = "EasyCopy2CMD.exe"

Private Const MY_COMPUTER = "[Computer]"

Private Const FORECOLOR_ENABLED = vbButtonText
Private Const FORECOLOR_DISABLED = 8355711

Private Const MULTIPLE_VALUES_SELECTED = "[Multiple values selected]"

Dim CurrentCopyJob As CopyJob
Dim SelectedItems As String

Dim WithEvents FileBrowserTimer As EasyCopy2DLL.EasyTimer
Attribute FileBrowserTimer.VB_VarHelpID = -1

Friend Sub SetForm()
  AddUACShieldToButton cmdGo.hWnd
  Me.Show
  pJobs.BorderStyle = 0
  pJob.BorderStyle = 0
  pName.BorderStyle = 0
  pItems.BorderStyle = 0
  pTarget.BorderStyle = 0
  pOptions.BorderStyle = 0
  ResizeObjects
  SetFormCaption
End Sub

Friend Sub Start()
  PopulateJobList
  If Not lstJobs.ListCount = 0 Then
    lstJobs.ListIndex = 0
  End If
  SetSelectedCopyJob
End Sub

Friend Sub FileBrowserUnloaded()
  Set FileBrowserTimer = New EasyCopy2DLL.EasyTimer
  With FileBrowserTimer
    .Interval = 50
    .Enabled = True
  End With
End Sub

Private Sub ResizeObjects()
  On Error GoTo ResizeError
  fJobs.Move 15, 15
  pJobs.Move 60, 210
  lstJobs.Move 0, 0
  fJobs.Height = (Me.ScaleHeight - (fJobs.Top * 2))
  pJobs.Height = (fJobs.Height - 270)
  cmdNewJob.Move 0, (pJobs.ScaleHeight - cmdNewJob.Height)
  cmdDelJob.Move ((cmdNewJob.Left + cmdNewJob.Width) + 75), cmdNewJob.Top
  pJobs.Width = ((pJobs.Width - pJobs.ScaleWidth) + ((cmdDelJob.Left + cmdDelJob.Width) + cmdNewJob.Left))
  fJobs.Width = (pJobs.Width + 120)
  lstJobs.Move 0, 0, pJobs.ScaleWidth, (cmdNewJob.Top - 15)
  fJob.Move ((fJobs.Left + fJobs.Width) + 30), fJobs.Top
  fJob.Width = (Me.ScaleWidth - (fJob.Left + fJobs.Left))
  fJob.Height = fJobs.Height
  pJob.Move 60, 210, (fJob.Width - 120), (fJob.Height - 270)
  cmdGo.Move (pJob.ScaleWidth - cmdGo.Width), (pJob.ScaleHeight - cmdGo.Height)
  cmdSave.Move (cmdGo.Left - (cmdSave.Width + 45)), cmdGo.Top
  cmdAdvanced.Move (fJob.Left + 60), (cmdSave.Top + 210)
  fName.Move 0, 0, pJob.ScaleWidth
  pName.Move 60, 210, (fName.Width - 120)
  txtName.Move 0, 0, pName.ScaleWidth
  pName.Height = (txtName.Height + txtName.Top)
  fName.Height = (pName.Height + 270)
  fItems.Move fName.Left, (fName.Top + fName.Height), fName.Width
  fItems.Height = (cmdGo.Top - (fItems.Top + 15))
  pItems.Move 60, 210, (fItems.Width - 120), (fItems.Height - 270)
  fOptions.Left = 0
  fOptions.Width = pItems.ScaleWidth
  pOptions.Move 60, 210, (fOptions.Width - 120)
  fTarget.Move 0, 0, pOptions.ScaleWidth
  pTarget.Move 60, 210, (fTarget.Width - 120)
  cmdTarget.Move (pTarget.ScaleWidth - cmdTarget.Width), 0
  txtTarget.Move 0, 0, (cmdTarget.Left - 45)
  pTarget.Height = ((pTarget.Height - pTarget.ScaleHeight) + txtTarget.Height)
  fTarget.Height = (pTarget.Height + 270)
  chkSubFolders.Move 15, (fTarget.Height + 45)
  chkAttributes.Move chkSubFolders.Left, ((chkSubFolders.Top + chkSubFolders.Height) + 75)
  chkOverwrite.Move chkAttributes.Left, ((chkAttributes.Top + chkAttributes.Height) + 75)
  chkDelete.Move (pOptions.ScaleWidth - chkDelete.Width), chkAttributes.Top
  chkError.Move chkDelete.Left, chkSubFolders.Top
  chkDelError.Move (chkDelete.Left + 270), (chkOverwrite.Top - 45)
  pOptions.Height = ((pOptions.Height - pOptions.ScaleHeight) + (chkOverwrite.Top + chkOverwrite.Height))
  fOptions.Height = (pOptions.Height + 270)
  fOptions.Top = (pItems.ScaleHeight - fOptions.Height)
  lvwItems.Move 0, 0, pItems.ScaleWidth, (fOptions.Top - 15)
  On Error GoTo 0
  Exit Sub
ResizeError:
  Resume Next
End Sub

Private Sub SetFormCaption()
  Me.Caption = Replace(FORM_CAPTION, "%ver%", GetAppVersion)
End Sub

Private Function GetAppVersion() As String
  With App
    GetAppVersion = Trim(Str(.Major)) & "." & Trim(Str(.Minor)) & "." & Trim(Str(.Revision))
  End With
End Function

Private Sub PopulateJobList()
  Dim i As Long
  lstJobs.Clear
  For i = 1 To AllCopyJobs.CopyJobs
    lstJobs.AddItem AllCopyJobs.CopyJob(i).Name
  Next
End Sub

Private Sub SetSelectedCopyJob()
  Dim cJ As CopyJob
  If lstJobs.ListIndex = -1 Then
    Set CurrentCopyJob = Nothing
    UpdateCopyJobData
  Else
    Set cJ = AllCopyJobs.CopyJob((lstJobs.ListIndex + 1))
    If Not cJ Is CurrentCopyJob Then
      Set CurrentCopyJob = cJ
      UpdateCopyJobData
    End If
  End If
End Sub

Private Sub UpdateCopyJobData()
  ClearItemsList
  EnableObjects
  If CurrentCopyJob Is Nothing Then
    txtName.Text = ""
    SetSelectedItems
    UpdateSelectedItems
  Else
    txtName.Text = CurrentCopyJob.Name
    PopulateItemsList
    SetSelectedItems
  End If
End Sub

Private Sub PopulateItemsList()
  Dim i As Long
  PrepareItemsList
  For i = 1 To CurrentCopyJob.JobItems
    AddJobItemToList CurrentCopyJob.JobItem(i)
  Next
End Sub

Private Sub AddJobItemToList(NewJobItem As JobItem)
  Dim lI As ComctlLib.ListItem
  Set lI = lvwItems.ListItems.Add(, NewJobItem.Key)
  If Len(NewJobItem.SourcePath) <= 3 Then
    lI.Text = GetDiskName(Left(NewJobItem.SourcePath, 1))
    lI.SubItems(1) = MY_COMPUTER
  Else
    lI.Text = GetFileName(NewJobItem.SourcePath)
    lI.SubItems(1) = GetParentFolder(NewJobItem.SourcePath) & "\"
  End If
  lI.SmallIcon = imlItems.ListImages.Add(, , GetAssociatedIcon(NewJobItem.SourcePath, False)).Index
  NewJobItem.ListItem = lI
  NewJobItem.ListViewParent = lvwItems
End Sub

Private Sub PrepareItemsList()
  imlItems.ImageWidth = 16
  imlItems.ImageHeight = 16
  imlItems.ListImages.Add , , Me.Icon
  Set lvwItems.SmallIcons = imlItems
  lvwItems.View = lvwReport
  With lvwItems.ColumnHeaders
    .Add , "Name", "Name"
    .Add , "Path", "Source Path"
    .Add , "Size", "Size"
    .Add , "Flags", "Flags"
    .Add , "Target", "Target Path"
  End With
End Sub

Private Sub ClearItemsList()
  lvwItems.ListItems.Clear
  lvwItems.ColumnHeaders.Clear
  Set lvwItems.SmallIcons = Nothing
  imlItems.ListImages.Clear
End Sub

Private Sub EnableObjects()
  Dim dEn As Boolean
  dEn = Not CurrentCopyJob Is Nothing
  cmdDelJob.Enabled = dEn
  mnuDeleteCopyJob.Enabled = dEn
  fJob.Enabled = dEn
  fJob.ForeColor = GetForeColor(dEn)
  fName.Enabled = dEn
  fName.ForeColor = GetForeColor(dEn)
  txtName.Enabled = dEn
  lvwItems.Enabled = dEn
  cmdSave.Enabled = dEn
  cmdGo.Enabled = dEn
  If dEn Then
    If SelectedItems = "" Then
      dEn = False
    End If
  End If
  fOptions.Enabled = dEn
  fOptions.ForeColor = GetForeColor(dEn)
  fTarget.Enabled = dEn
  fTarget.ForeColor = GetForeColor(dEn)
  txtTarget.Enabled = dEn
  cmdTarget.Enabled = dEn
  chkSubFolders.Enabled = dEn
  chkAttributes.Enabled = dEn
  chkOverwrite.Enabled = dEn
  chkError.Enabled = dEn
  chkDelete.Enabled = dEn
  If Not chkDelete.Enabled Then
    chkDelError.Enabled = False
  Else
    chkDelError.Enabled = dEn
  End If
End Sub

Private Sub SetSelectedItems()
  Dim lI As ListItem
  Dim sI As String
  sI = ""
  For Each lI In lvwItems.ListItems
    If lI.Selected Then
      sI = sI & lI.Key & vbNullChar
    End If
  Next
  If Not SelectedItems = sI Then
    SelectedItems = sI
    UpdateSelectedItems
  End If
End Sub

Private Sub UpdateSelectedItems()
  EnableObjects
  ChangedByCode = True
  If SelectedItems = "" Then
    txtTarget.Text = ""
    chkSubFolders.Value = vbUnchecked
    chkAttributes.Value = vbUnchecked
    chkOverwrite.Value = vbUnchecked
    chkError.Value = vbUnchecked
    chkDelete.Value = vbUnchecked
    chkDelError.Value = vbUnchecked
  Else
    PopulateSelectedItemsData
  End If
  ChangedByCode = False
End Sub

Private Sub PopulateSelectedItemsData()
  Dim i As Long
  Dim j As Long
  Dim sI As String
  Dim tI As String
  Dim tTarget As String
  Dim tSub As CheckBoxConstants
  Dim tAttrib As CheckBoxConstants
  Dim tOverwrite As CheckBoxConstants
  Dim tIgnore As CheckBoxConstants
  Dim tDelete As CheckBoxConstants
  Dim tDelErr As CheckBoxConstants
  sI = SelectedItems
  Do Until sI = ""
    i = InStr(sI, vbNullChar)
    If i = 0 Then
      tI = sI
      sI = ""
    Else
      tI = Left(sI, (i - 1))
      sI = Right(sI, (Len(sI) - i))
    End If
    j = (j + 1)
    With CurrentCopyJob.GetJobItemFromKey(tI)
      If j = 1 Then
        tTarget = .TargetPath
        tSub = GetItemValue(.IncludeSubFolders)
        tAttrib = GetItemValue(.ResetAttributes)
        tOverwrite = GetItemValue(.Overwrite)
        tIgnore = GetItemValue(.IgnoreErrors)
        tDelete = GetItemValue(.DeleteAfterCopy)
        tDelErr = GetItemValue(.DeleteAfterError)
      Else
        If Not tTarget = .TargetPath Then
          tTarget = MULTIPLE_VALUES_SELECTED
        End If
        tSub = CompareItemValues(tSub, .IncludeSubFolders)
        tAttrib = CompareItemValues(tAttrib, .ResetAttributes)
        tOverwrite = CompareItemValues(tOverwrite, .Overwrite)
        tIgnore = CompareItemValues(tIgnore, .IgnoreErrors)
        tDelete = CompareItemValues(tDelete, .DeleteAfterCopy)
        tDelErr = CompareItemValues(tDelErr, .DeleteAfterError)
      End If
    End With
  Loop
  txtTarget.Text = tTarget
  chkSubFolders.Value = tSub
  chkAttributes.Value = tAttrib
  chkOverwrite.Value = tOverwrite
  chkError.Value = tIgnore
  chkDelete.Value = tDelete
  chkDelError.Value = tDelErr
  Select Case chkDelete.Value
    Case vbChecked, vbGrayed
      chkDelError.Enabled = chkDelete.Enabled
    Case Else
      chkDelError.Enabled = False
  End Select
End Sub

Private Function GetItemValue(BoolValue As Boolean) As CheckBoxConstants
  If BoolValue Then
    GetItemValue = vbChecked
  Else
    GetItemValue = vbUnchecked
  End If
End Function

Private Function CompareItemValues(TempValue As CheckBoxConstants, BoolValue As Boolean) As CheckBoxConstants
  If BoolValue Then
    If TempValue = vbChecked Then
      CompareItemValues = vbChecked
    Else
      CompareItemValues = vbGrayed
    End If
  Else
    If TempValue = vbUnchecked Then
      CompareItemValues = vbUnchecked
    Else
      CompareItemValues = vbGrayed
    End If
  End If
End Function

Private Function GetForeColor(IsEnabled As Boolean) As Long
  If IsEnabled Then
    GetForeColor = FORECOLOR_ENABLED
  Else
    GetForeColor = FORECOLOR_DISABLED
  End If
End Function

Private Function GetDefaultJobName() As String
  Dim i As Long
  Dim j As Long
  Dim nN As String
  Do
    i = (i + 1)
    nN = "EasyCopy Job #" & Trim(Str(i))
    For j = 1 To AllCopyJobs.CopyJobs
      If nN = AllCopyJobs.CopyJob(j).Name Then
        nN = ""
        Exit For
      End If
    Next
  Loop Until nN <> ""
  GetDefaultJobName = nN
End Function

Private Sub AddItemToCopy(Path As String)
  Dim nI As JobItem
  Set nI = CurrentCopyJob.AddJobItem(Path)
  If nI.IsFolder Then
    nI.IncludeSubFolders = True
  End If
  AddJobItemToList nI
End Sub

Private Sub SetSelectedItemsValues()
  Dim i As Long
  Dim sI As String
  Dim tI As String
  Dim jI As JobItem
  If Not ChangedByCode Then
    sI = SelectedItems
    Do Until sI = ""
      i = InStr(sI, vbNullChar)
      If i = 0 Then
        tI = sI
        sI = ""
      Else
        tI = Left(sI, (i - 1))
        sI = Right(sI, (Len(sI) - i))
      End If
      Set jI = CurrentCopyJob.GetJobItemFromKey(tI)
      If Not jI Is Nothing Then
        SetThisJobItemsValues jI
      End If
    Loop
  End If
End Sub

Private Sub SetThisJobItemsValues(ThisJobItem As JobItem)
  With ThisJobItem
    If Not txtTarget.Text = MULTIPLE_VALUES_SELECTED Then
      .TargetPath = txtTarget.Text
    End If
    Select Case chkSubFolders.Value
      Case vbChecked
        .IncludeSubFolders = True
      Case vbUnchecked
        .IncludeSubFolders = False
    End Select
    Select Case chkAttributes.Value
      Case vbChecked
        .ResetAttributes = True
      Case vbUnchecked
        .ResetAttributes = False
    End Select
    Select Case chkOverwrite.Value
      Case vbChecked
        .Overwrite = True
      Case vbUnchecked
        .Overwrite = False
    End Select
    Select Case chkError.Value
      Case vbChecked
        .IgnoreErrors = True
      Case vbUnchecked
        .IgnoreErrors = False
    End Select
    Select Case chkDelete.Value
      Case vbChecked
        .DeleteAfterCopy = True
      Case vbUnchecked
        .DeleteAfterCopy = False
    End Select
    Select Case chkDelError.Value
      Case vbChecked
        .DeleteAfterError = True
      Case vbUnchecked
        .DeleteAfterError = False
    End Select
  End With
End Sub

Private Function AllJobsSaved() As Boolean
  Dim i As Long
  Dim jS As Boolean
  jS = True
  For i = 1 To AllCopyJobs.CopyJobs
    If Not AllCopyJobs.CopyJob(i).JobSaved Then
      jS = False
      Exit For
    End If
  Next
  AllJobsSaved = jS
End Function

Private Sub RefreshCurrentJob()
  If Not CurrentCopyJob Is Nothing Then
    CurrentCopyJob.RefreshJobSize
  End If
End Sub

Private Function CanExit() As Boolean
  If Not AllJobsSaved Then
    Select Case MsgBox("One or more jobs has not been saved. Do you want to save them before exiting?", vbYesNoCancel Or vbQuestion, "Save before exit?")
      Case vbYes
        If AllCopyJobs.SaveAllCopyJobs Then
          CanExit = True
        Else
          MsgBox "EasyCopy encountered an unknown error when trying to save the jobs!" & vbCrLf & vbCrLf & "Program termination has been cancelled!", vbOKOnly Or vbExclamation, "Unknown Error"
        End If
      Case vbNo
        CanExit = True
    End Select
  Else
    CanExit = True
  End If
End Function

Private Function GetCMDPath() As String
  GetCMDPath = FixPath(App.Path) & "\" & CMD_EXENAME
End Function

Private Function DeleteIfErrorWarning() As Boolean
  Dim wM As String
  wM = "WARNING:" & vbCrLf & vbCrLf
  wM = wM & "By selecting this option, the files and folders selected will be deleted even if copying has somehow failed!" & vbCrLf
  wM = wM & "You run the risk of accidental loss of data!"
  Select Case MsgBox(wM, vbOKCancel Or vbCritical Or vbApplicationModal, "Attention")
    Case vbOK
      DeleteIfErrorWarning = True
  End Select
End Function

Private Sub chkAttributes_Click()
  SetSelectedItemsValues
End Sub

Private Sub chkDelError_Click()
  If chkDelError.Value = vbChecked Then
    If DeleteIfErrorWarning Then
      SetSelectedItemsValues
    Else
      chkDelError.Value = vbUnchecked
    End If
  End If
End Sub

Private Sub chkDelete_Click()
  SetSelectedItemsValues
  Select Case chkDelete.Value
    Case vbChecked, vbGrayed
      chkDelError.Enabled = chkDelete.Enabled
    Case Else
      chkDelError.Enabled = False
  End Select
End Sub

Private Sub chkError_Click()
  SetSelectedItemsValues
End Sub

Private Sub chkOverwrite_Click()
  SetSelectedItemsValues
End Sub

Private Sub chkSubFolders_Click()
  SetSelectedItemsValues
End Sub

Private Sub cmdAdvanced_Click()
  If Not SettingsFormLoaded Then
    Load frmSettings
    frmSettings.SetForm
    SettingsFormLoaded = True
  Else
    frmSettings.SetFocus
  End If
End Sub

Private Sub cmdDelJob_Click()
  Dim dMsg As String
  Dim lI As Integer
  Dim jRem As Boolean
  dMsg = "Are you sure you want to delete the Job named """ & CurrentCopyJob.Name & """?" & vbCrLf & vbCrLf
  dMsg = dMsg & "This action is irreversible and cannot be undone!"
  Select Case MsgBox(dMsg, vbYesNo Or vbQuestion, "Delete Copy Job")
    Case vbYes
      lI = lstJobs.ListIndex
      jRem = AllCopyJobs.RemoveCopyJob(lI + 1)
      lstJobs.RemoveItem lI
      If Not lstJobs.ListCount = 0 Then
        If lI >= lstJobs.ListCount Then
          lI = (lstJobs.ListCount - 1)
        End If
        lstJobs.ListIndex = lI
      Else
        SetSelectedCopyJob
      End If
      If Not jRem Then
        MsgBox "EasyCopy encountered an unknown error when trying to remove the local file!", vbOKOnly Or vbExclamation, "File Error"
      End If
  End Select
End Sub

Private Sub cmdGo_Click()
  If Not CurrentCopyJob Is Nothing Then
    If CurrentCopyJob.CanRun Then
      If Not LaunchCopyJob(GetCMDPath, CurrentCopyJob) Then
        MsgBox "An unknown error occurred when trying to start the selected Job!", vbOKOnly Or vbExclamation, "Unknown error!"
      End If
    Else
      MsgBox "This job cannot be run!" & vbCrLf & vbCrLf & "Make sure that all items has a valid target path.", vbOKOnly Or vbInformation, "Cannot start!"
    End If
  End If
End Sub

Private Sub cmdNewJob_Click()
  Dim jN As String
  jN = InputBox("Input a name for the new job:", "Create new Copy Job...", GetDefaultJobName)
  If Not jN = "" Then
    With AllCopyJobs.AddCopyJob(jN)
      lstJobs.AddItem .Name
      lstJobs.ListIndex = (lstJobs.ListCount - 1)
      'SetSelectedCopyJob
    End With
  End If
End Sub

Private Sub cmdSave_Click()
  If CurrentCopyJob.JobSaved Then
    MsgBox "No need to save this!", vbOKOnly Or vbInformation, "User error"
  Else
    If Not CurrentCopyJob.SaveJob Then
      MsgBox "An unknown error occurred trying to save this job!", vbOKOnly Or vbExclamation, "Error!"
    Else
      MsgBox "Job saved!", vbOKOnly Or vbInformation, "Yay!"
    End If
  End If
End Sub

Private Sub cmdTarget_Click()
  Dim nP As String
  nP = BrowseForFolder("Select target path...", Me.hWnd)
  If Not nP = "" Then
    txtTarget.Text = nP
  End If
End Sub

Private Sub FileBrowserTimer_Timer()
  FileBrowserTimer.Enabled = False
  Set FileBrowserTimer = Nothing
  Set frmFileBrowser = Nothing
End Sub

Private Sub Form_Resize()
  ResizeObjects
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not CanExit Then
    Cancel = 1
  Else
    Set AllCopyJobs = Nothing
    If Not FileBrowserTimer Is Nothing Then
      FileBrowserTimer_Timer
    End If
  End If
End Sub

Private Sub lstJobs_Click()
  SetSelectedCopyJob
End Sub

Private Sub lstJobs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    mnuDelete.Enabled = Not CurrentCopyJob Is Nothing
    PopupMenu mnuJobs
  End If
End Sub

Private Sub lvwItems_Click()
  SetSelectedItems
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As ComctlLib.ListItem)
  SetSelectedItems
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDelete
      mnuRemove_Click
    Case vbKeyF5
      RefreshCurrentJob
  End Select
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    mnuRemove.Enabled = Not SelectedItems = ""
    PopupMenu mnuItems
  End If
End Sub

Private Sub lvwItems_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Long
  If Data.GetFormat(vbCFFiles) Then
    For i = 1 To Data.Files.Count
      AddItemToCopy Data.Files.Item(i)
    Next
  End If
End Sub

Private Sub mnuAddFiles_Click()
  Dim i As Long
  Dim nF As String
  Dim tP As String
  Dim tF As String
  nF = BrowseForFile("Select the file(s) you wish to add to the current task...", Me.hWnd)
  While Len(nF) > 2
    i = InStr(nF, vbNullChar)
    If i = 0 Then
      tF = nF
      nF = ""
    Else
      tF = Left(nF, (i - 1))
      nF = Right(nF, (Len(nF) - i))
    End If
    If tP = "" Then
      tP = FixPath(tF)
      tF = ""
    Else
      AddItemToCopy tP & "\" & tF
    End If
  Wend
  If Not tP = "" Then
    If tF = "" Then
      AddItemToCopy tP
    End If
  End If
End Sub

Private Sub mnuAddFolder_Click()
  Dim nP As String
  nP = BrowseForFolder("Select the folder you wish to add to the current task...", Me.hWnd)
  If Not nP = "" Then
    AddItemToCopy nP
  End If
End Sub

Private Sub mnuDelete_Click()
  cmdDelJob_Click
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuFileBrowser_Click()
  Load frmFileBrowser
  frmFileBrowser.SetForm
End Sub

Private Sub mnuNew_Click()
  cmdNewJob_Click
End Sub

Private Sub mnuNewCopyJob_Click()
  cmdNewJob_Click
End Sub

Private Sub mnuRemove_Click()
  Dim i As Long
  For i = lvwItems.ListItems.Count To 1 Step -1
    If lvwItems.ListItems.Item(i).Selected Then
      CurrentCopyJob.RemoveJobItem CurrentCopyJob.GetJobIndexFromKey(lvwItems.ListItems(i).Key)
      lvwItems.ListItems.Remove i
    End If
  Next
  lvwItems_Click
End Sub

Private Sub txtName_Change()
  If Not CurrentCopyJob Is Nothing Then
    CurrentCopyJob.Name = txtName.Text
    lstJobs.List(lstJobs.ListIndex) = CurrentCopyJob.Name
  End If
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtTarget_Change()
  SetSelectedItemsValues
End Sub

Private Sub txtTarget_GotFocus()
  With txtTarget
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub
