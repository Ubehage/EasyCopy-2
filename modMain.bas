Attribute VB_Name = "modMain"
Option Explicit

Global AllCopyJobs As CopyJobs

Global BufferSettings As BUFFER_Settings

Global ChangedByCode As Boolean

Global UnloadedByCode As Boolean

Global SettingsFormLoaded As Boolean
Global BrowserFormLoaded As Boolean

Sub Main()
  InitCommonControls
  Randomize Timer
  Set AllCopyJobs = New CopyJobs
  AllCopyJobs.LoadAllCopyJobs
  Load frmMain
  frmMain.SetForm
  frmMain.Start
End Sub
