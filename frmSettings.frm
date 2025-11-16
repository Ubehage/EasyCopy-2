VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ubehage's EasyCopy v2 - Advanced Options"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Discard && Close"
      Height          =   435
      Left            =   2400
      TabIndex        =   21
      Top             =   5610
      Width           =   1350
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save && Close"
      Height          =   435
      Left            =   705
      TabIndex        =   20
      Top             =   5670
      Width           =   1350
   End
   Begin VB.Frame frmBuffer 
      Caption         =   "Copy Buffer"
      Height          =   4755
      Left            =   540
      TabIndex        =   0
      Top             =   780
      Width           =   8325
      Begin VB.PictureBox picBuffer 
         Height          =   3960
         Left            =   240
         ScaleHeight     =   3900
         ScaleWidth      =   7680
         TabIndex        =   1
         Top             =   375
         Width           =   7740
         Begin VB.CheckBox chkPriority 
            Caption         =   "Always write directly (do NOT use Windows Caching)"
            Height          =   225
            Left            =   285
            TabIndex        =   19
            Top             =   3570
            Width           =   4125
         End
         Begin VB.Frame frmLowHigh 
            Height          =   1770
            Left            =   210
            TabIndex        =   14
            Top             =   1650
            Width           =   4350
            Begin VB.PictureBox picLowHigh 
               Height          =   1380
               Left            =   225
               ScaleHeight     =   1320
               ScaleWidth      =   3615
               TabIndex        =   15
               Top             =   210
               Width           =   3675
               Begin VB.OptionButton optAverage 
                  Caption         =   "Always use average common buffer size"
                  Height          =   225
                  Left            =   195
                  TabIndex        =   18
                  Top             =   840
                  Width           =   3150
               End
               Begin VB.OptionButton optHighest 
                  Caption         =   "Always use highest common buffer size"
                  Height          =   225
                  Left            =   150
                  TabIndex        =   17
                  Top             =   465
                  Width           =   3150
               End
               Begin VB.OptionButton optLowest 
                  Caption         =   "Always use lowest common buffer size"
                  Height          =   225
                  Left            =   165
                  TabIndex        =   16
                  Top             =   180
                  Width           =   3150
               End
            End
         End
         Begin VB.ComboBox cmbRemovable 
            Height          =   315
            Left            =   4290
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1320
            Width           =   1740
         End
         Begin VB.ComboBox cmbNetwork 
            Height          =   315
            Left            =   4215
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   915
            Width           =   1740
         End
         Begin VB.ComboBox cmbOptical 
            Height          =   315
            Left            =   4245
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   570
            Width           =   1740
         End
         Begin VB.ComboBox cmbHardDisk 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   210
            Width           =   1740
         End
         Begin VB.TextBox txtRemovable 
            Height          =   315
            Left            =   2850
            TabIndex        =   9
            Text            =   "1234567890"
            Top             =   1230
            Width           =   1260
         End
         Begin VB.TextBox txtNetwork 
            Height          =   315
            Left            =   2760
            TabIndex        =   8
            Text            =   "1234567890"
            Top             =   855
            Width           =   1260
         End
         Begin VB.TextBox txtOptical 
            Height          =   315
            Left            =   2775
            TabIndex        =   7
            Text            =   "1234567890"
            Top             =   540
            Width           =   1260
         End
         Begin VB.TextBox txtHardDisk 
            Height          =   315
            Left            =   2865
            TabIndex        =   6
            Text            =   "1234567890"
            Top             =   225
            Width           =   1260
         End
         Begin VB.Label lblRemovable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buffer Size for Removable Disks:"
            Height          =   195
            Left            =   210
            TabIndex        =   5
            Top             =   1170
            Width           =   2325
         End
         Begin VB.Label lblNetwork 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buffer Size for Network Drives:"
            Height          =   195
            Left            =   315
            TabIndex        =   4
            Top             =   900
            Width           =   2175
         End
         Begin VB.Label lblOptical 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buffer Size for Optical Drives:"
            Height          =   195
            Left            =   255
            TabIndex        =   3
            Top             =   510
            Width           =   2070
         End
         Begin VB.Label lblHardDisk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buffer Size for Hard Disks:"
            Height          =   195
            Left            =   255
            TabIndex        =   2
            Top             =   150
            Width           =   1860
         End
      End
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   360
      Left            =   150
      TabIndex        =   22
      Top             =   60
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   635
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempBuffer As BUFFER_Settings

Dim ListIndexHD As Single
Dim ListIndexOpt As Single
Dim ListIndexNet As Single
Dim ListIndexRem As Single

Dim IsChanged As Boolean

Friend Sub SetForm()
  ChangedByCode = True
  Me.Show , frmMain
  ResizeObjects
  SetTabLabels
  SetComboContents
  ReadBufferSettings TempBuffer
  FillTextBoxes
  Select Case TempBuffer.BufferToUse
    Case BUFFER_USERS.buLowest
      optLowest.Value = 1
    Case BUFFER_USERS.byHighest
      optHighest.Value = 1
    Case BUFFER_USERS.buAverage
      optAverage.Value = 1
  End Select
  If TempBuffer.AlwaysWrite Then
    chkPriority.Value = vbChecked
  Else
    chkPriority.Value = vbUnchecked
  End If
  ChangedByCode = False
End Sub

Private Sub ResizeObjects()
  Tab1.Move 60, 60
  frmBuffer.Move Tab1.Left, ((Tab1.Top + Tab1.Height) + 15)
  picBuffer.Move 60, 210
  lblHardDisk.Move 0, 0
  txtHardDisk.Move ((lblHardDisk.Left + GetWidestLabel) + 45), 0
  txtOptical.Move txtHardDisk.Left, ((txtHardDisk.Top + txtHardDisk.Height) + 15)
  txtNetwork.Move txtOptical.Left, ((txtOptical.Top + txtOptical.Height) + 15)
  txtRemovable.Move txtNetwork.Left, ((txtNetwork.Top + txtNetwork.Height) + 15)
  cmbHardDisk.Move ((txtHardDisk.Left + txtHardDisk.Width) + 15), txtHardDisk.Top
  cmbOptical.Move cmbHardDisk.Left, txtOptical.Top
  cmbNetwork.Move cmbOptical.Left, txtNetwork.Top
  cmbRemovable.Move cmbNetwork.Left, txtRemovable.Top
  lblHardDisk.Top = (txtHardDisk.Top + ((txtHardDisk.Height - lblHardDisk.Height) / 2))
  lblOptical.Move lblHardDisk.Left, (txtOptical.Top + ((txtOptical.Height - lblOptical.Height) / 2))
  lblNetwork.Move lblOptical.Left, (txtNetwork.Top + ((txtNetwork.Height - lblNetwork.Height) / 2))
  lblRemovable.Move lblNetwork.Left, (txtRemovable.Top + ((txtRemovable.Height - lblRemovable.Height) / 2))
  frmLowHigh.Move lblRemovable.Left, (txtRemovable.Top + txtRemovable.Height)
  picLowHigh.Move 60, 210
  optLowest.Move 0, 0
  optHighest.Move optLowest.Left, ((optLowest.Top + optLowest.Height) + 45)
  optAverage.Move optHighest.Left, ((optHighest.Top + optHighest.Height) + 45)
  picLowHigh.BorderStyle = 0
  picLowHigh.Width = ((picLowHigh.Width - picLowHigh.ScaleWidth) + (GetWidestOption + (optLowest.Left * 2)))
  picLowHigh.Height = ((picLowHigh.Height - picLowHigh.ScaleHeight) + ((optAverage.Top + optAverage.Height) + optLowest.Top))
  frmLowHigh.Width = (picLowHigh.Width + 120)
  frmLowHigh.Height = (picLowHigh.Height + 270)
  chkPriority.Move frmLowHigh.Left, ((frmLowHigh.Top + frmLowHigh.Height) + 15)
  picBuffer.BorderStyle = 0
  picBuffer.Width = ((picBuffer.Width - picBuffer.ScaleWidth) + ((cmbHardDisk.Left + cmbHardDisk.Width) + lblHardDisk.Left))
  picBuffer.Height = ((picBuffer.Height - picBuffer.ScaleHeight) + ((chkPriority.Top + chkPriority.Height) + txtHardDisk.Top))
  frmBuffer.Width = (picBuffer.Width + 120)
  frmBuffer.Height = (picBuffer.Height + 270)
  Tab1.Width = frmBuffer.Width
  Me.Width = ((Me.Width - Me.ScaleWidth) + (frmBuffer.Width + (frmBuffer.Left * 2)))
  cmdCancel.Move ((frmBuffer.Left + frmBuffer.Width) - cmdCancel.Width), ((frmBuffer.Top + frmBuffer.Height) + 15)
  cmdOK.Move (cmdCancel.Left - (cmdOK.Width + 45)), cmdCancel.Top
  Me.Height = ((Me.Height - Me.ScaleHeight) + ((cmdOK.Top + cmdOK.Height) + Tab1.Top))
End Sub

Private Function GetWidestLabel() As Long
  Dim wL As Long
  wL = GetHighestValue(lblHardDisk.Width, lblOptical.Width)
  wL = GetHighestValue(wL, lblNetwork.Width)
  wL = GetHighestValue(wL, lblRemovable.Width)
  GetWidestLabel = wL
End Function

Private Function GetWidestOption() As Long
  Dim wO As Long
  wO = GetHighestValue(optLowest.Width, optHighest.Width)
  wO = GetHighestValue(wO, optAverage.Width)
  GetWidestOption = wO
End Function

Private Function GetHighestValue(Value1 As Long, Value2 As Long) As Long
  If Value1 >= Value2 Then
    GetHighestValue = Value1
  Else
    GetHighestValue = Value2
  End If
End Function

Private Sub SetTabLabels()
  With Tab1.Tabs
    .Item(1).Caption = "Global Settings"
    '.Add , , "For Current Job Only"
  End With
End Sub

Private Sub SetComboContents()
  SetSingleCombo cmbHardDisk
  SetSingleCombo cmbOptical
  SetSingleCombo cmbNetwork
  SetSingleCombo cmbRemovable
End Sub

Private Sub SetSingleCombo(ThisCombo As ComboBox)
  With ThisCombo
    .Clear
    .AddItem "B (Bytes)"
    .AddItem "KB (KiloBytes)"
    .AddItem "MB (MegaBytes)"
    .ListIndex = 1
  End With
End Sub

Private Sub FillTextBoxes()
  FillSingleTextBox txtHardDisk
  FillSingleTextBox txtOptical
  FillSingleTextBox txtNetwork
  FillSingleTextBox txtRemovable
End Sub

Private Sub FillSingleTextBox(ThisTextBox As TextBox)
  Dim v As Long
  Dim t As String
  If ThisTextBox Is txtHardDisk Then
    t = GetValueFromCombo(TempBuffer.HardDiskBuffer, cmbHardDisk)
  ElseIf ThisTextBox Is txtOptical Then
    t = GetValueFromCombo(TempBuffer.OpticalBuffer, cmbOptical)
  ElseIf ThisTextBox Is txtNetwork Then
    t = GetValueFromCombo(TempBuffer.NetworkBuffer, cmbNetwork)
  ElseIf ThisTextBox Is txtRemovable Then
    t = GetValueFromCombo(TempBuffer.RemovableBuffer, cmbRemovable)
  End If
  ThisTextBox.Text = t
End Sub

Private Function GetValueFromCombo(ThisValue As Double, ThisCombo As ComboBox) As String
  Dim v As Double
  Dim rV As String
  v = ThisValue
  If ThisCombo.ListIndex = 0 Then
    'do nothing...
  Else
    v = (v / 1024)
    If ThisCombo.ListIndex = 1 Then
      'do nothing...
    Else
      v = (v / 1024)
    End If
  End If
  GetValueFromCombo = GetRoundedNumber(v, 2)
End Function

Private Function GetRoundedNumber(ThisNumber As Double, NumDecimals As Long) As String
  Dim i As Long
  Dim rN As String
  rN = Trim(Str(ThisNumber))
  'i = InStr(rN, ".")
  'If i > 0 Then
  '  rN = Left(rN, (i + NumDecimals))
  'End If
  If Left(rN, 1) = "." Then
    rN = "0" & rN
  ElseIf Right(rN, 1) = "." Then
    rN = Left(rN, (Len(rN) - 1))
  End If
  GetRoundedNumber = rN
End Function

Private Function GetNewValueFromCombo(OldValue As String, OldIndex As Single, NewIndex As Single) As String
  Dim i As Long
  Dim v As Double
  i = OldIndex
  v = CDbl(Val(OldValue))
  Do Until i = NewIndex
    If i > NewIndex Then
      v = (v * 1024)
      i = (i - 1)
    ElseIf i < NewIndex Then
      v = (v / 1024)
      i = (i + 1)
    End If
  Loop
  GetNewValueFromCombo = GetRoundedNumber(v, 2)
End Function

Private Function GetActualValueFromCombo(ThisValue As String, ComboIndex As Single) As Double
  Dim v As Double
  v = CLng(Val(ThisValue))
  If ComboIndex = 0 Then
    'do nothing...
  Else
    v = (v * 1024)
    If ComboIndex = 1 Then
      'do nothing...
    Else
      v = (v * 1024)
    End If
  End If
  GetActualValueFromCombo = v
End Function

Private Function CanCancel() As Boolean
  Dim r As Boolean
  If Not IsChanged Then
    r = True
  Else
    Select Case MsgBox("Are you sure you want to close and discard all changes?", vbYesNo Or vbQuestion, "Discard Changes: Advanced Options")
      Case vbYes
        r = True
    End Select
  End If
  CanCancel = r
End Function

Private Sub chkPriority_Click()
  If Not ChangedByCode Then
    TempBuffer.AlwaysWrite = (chkPriority.Value = vbChecked)
    IsChanged = True
  End If
End Sub

Private Sub cmbHardDisk_Click()
  Dim i As Single
  i = cmbHardDisk.ListIndex
  If Not ChangedByCode Then
    ChangedByCode = True
    txtHardDisk.Text = GetNewValueFromCombo(txtHardDisk.Text, ListIndexHD, i)
    ChangedByCode = False
  End If
  ListIndexHD = i
End Sub

Private Sub cmbNetwork_Click()
  Dim i As Single
  i = cmbNetwork.ListIndex
  If Not ChangedByCode Then
    ChangedByCode = True
    txtNetwork.Text = GetNewValueFromCombo(txtNetwork.Text, ListIndexNet, i)
    ChangedByCode = False
  End If
  ListIndexNet = i
End Sub

Private Sub cmbOptical_Click()
  Dim i As Single
  i = cmbOptical.ListIndex
  If Not ChangedByCode Then
    ChangedByCode = True
    txtOptical.Text = GetNewValueFromCombo(txtOptical.Text, ListIndexOpt, i)
    ChangedByCode = False
  End If
  ListIndexOpt = i
End Sub

Private Sub cmbRemovable_Click()
  Dim i As Single
  i = cmbRemovable.ListIndex
  If Not ChangedByCode Then
    ChangedByCode = True
    txtRemovable.Text = GetNewValueFromCombo(txtRemovable.Text, ListIndexRem, i)
    ChangedByCode = False
  End If
  ListIndexRem = i
End Sub

Private Sub cmdCancel_Click()
  If CanCancel Then
    UnloadedByCode = True
    Unload Me
    UnloadedByCode = False
  End If
End Sub

Private Sub cmdOK_Click()
  If IsChanged Then
    SaveBufferSettings TempBuffer
    IsChanged = False
  End If
  UnloadedByCode = True
  Unload Me
  UnloadedByCode = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnloadedByCode Then
    If Not CanCancel Then
      Cancel = 1
    End If
  End If
  If Cancel = 0 Then
    SettingsFormLoaded = False
  End If
End Sub

Private Sub optAverage_Click()
  If Not ChangedByCode Then
    TempBuffer.BufferToUse = buAverage
    IsChanged = True
  End If
End Sub

Private Sub optHighest_Click()
  If Not ChangedByCode Then
    TempBuffer.BufferToUse = byHighest
    IsChanged = True
  End If
End Sub

Private Sub optLowest_Click()
  If Not ChangedByCode Then
    TempBuffer.BufferToUse = buLowest
    IsChanged = True
  End If
End Sub

Private Sub txtHardDisk_Change()
  If Not ChangedByCode Then
    TempBuffer.HardDiskBuffer = GetActualValueFromCombo(txtHardDisk.Text, cmbHardDisk.ListIndex)
    IsChanged = True
  End If
End Sub

Private Sub txtHardDisk_GotFocus()
  With txtHardDisk
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtHardDisk_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
      'do nothing...
    Case 44, 46
      If InStr(txtHardDisk.Text, ".") = 0 Then
        KeyAscii = 46
      Else
        KeyAscii = 0
      End If
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtNetwork_Change()
  If Not ChangedByCode Then
    TempBuffer.NetworkBuffer = GetActualValueFromCombo(txtNetwork.Text, cmbNetwork.ListIndex)
    IsChanged = True
  End If
End Sub

Private Sub txtNetwork_GotFocus()
  With txtNetwork
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtNetwork_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
      'do nothing...
    Case 44, 46
      If InStr(txtNetwork.Text, ".") = 0 Then
        KeyAscii = 46
      Else
        KeyAscii = 0
      End If
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtOptical_Change()
  If Not ChangedByCode Then
    TempBuffer.OpticalBuffer = GetActualValueFromCombo(txtOptical.Text, cmbOptical.ListIndex)
    IsChanged = True
  End If
End Sub

Private Sub txtOptical_GotFocus()
  With txtOptical
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtOptical_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
      'do nothing...
    Case 44, 46
      If InStr(txtOptical.Text, ".") = 0 Then
        KeyAscii = 46
      Else
        KeyAscii = 0
      End If
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub txtRemovable_Change()
  If Not ChangedByCode Then
    TempBuffer.RemovableBuffer = GetActualValueFromCombo(txtRemovable.Text, cmbRemovable.ListIndex)
    IsChanged = True
  End If
End Sub

Private Sub txtRemovable_GotFocus()
  With txtRemovable
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtRemovable_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
      'do nothing...
    Case 44, 46
      If InStr(txtRemovable.Text, ".") = 0 Then
        KeyAscii = 46
      Else
        KeyAscii = 0
      End If
    Case Else
      KeyAscii = 0
  End Select
End Sub
