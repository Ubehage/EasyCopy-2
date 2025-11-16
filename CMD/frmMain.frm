VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EasyCopy2"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7845
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShutdown 
      Caption         =   "Shut down your computer when this job is finished"
      Height          =   255
      Left            =   135
      TabIndex        =   6
      Top             =   3270
      Width           =   3810
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop copying!"
      Height          =   420
      Left            =   4455
      TabIndex        =   5
      Top             =   3210
      Width           =   1395
   End
   Begin VB.Frame fBack 
      Height          =   3045
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   4140
      Begin VB.PictureBox pBack 
         Height          =   2535
         Left            =   165
         ScaleHeight     =   2475
         ScaleWidth      =   3720
         TabIndex        =   1
         Top             =   345
         Width           =   3780
         Begin EasyCopy2CMD.Progress pFile 
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   1755
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   344
            Min             =   0
            Max             =   100
            Value           =   0
         End
         Begin EasyCopy2CMD.Progress pFiles 
            Height          =   240
            Left            =   210
            TabIndex        =   9
            Top             =   975
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   423
            Min             =   0
            Max             =   100
            Value           =   0
         End
         Begin EasyCopy2CMD.Progress pBytes 
            Height          =   210
            Left            =   240
            TabIndex        =   8
            Top             =   390
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   370
            Min             =   0
            Max             =   100
            Value           =   0
         End
         Begin VB.Label lFileName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lFileName"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   2010
            Width           =   690
         End
         Begin VB.Label lFile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current file:"
            Height          =   195
            Left            =   225
            TabIndex        =   4
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label lFiles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Files copied:"
            Height          =   195
            Left            =   165
            TabIndex        =   3
            Top             =   705
            Width           =   885
         End
         Begin VB.Label lBytes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total progress - bytes copied:"
            Height          =   195
            Left            =   195
            TabIndex        =   2
            Top             =   75
            Width           =   2085
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_CAPTION = "Ubehage's EasyCopy v2 - [Copying: %t%]"

Friend Sub SetForm()
  Me.Show
  ResizeObjects
  pBack.BorderStyle = 0
  ResizeObjects
  lBytes.Caption = ""
  lFiles.Caption = ""
  lFile.Caption = ""
  lFileName.Caption = ""
  Me.Caption = Replace(FORM_CAPTION, "%t%", CurrentCopyJob.Name)
End Sub

Friend Sub SetToPause(PauseNow As Boolean)
  pFiles.Visible = Not PauseNow
  lFile.Visible = pFiles.Visible
  pFile.Visible = lFile.Visible
End Sub

Private Sub ResizeObjects()
  lBytes.Move 0, 0
  pBytes.Move lBytes.Left, ((lBytes.Top + lBytes.Height) + 15), (15 * 600), lBytes.Height
  lFiles.Move pBytes.Left, ((pBytes.Top + pBytes.Height) + 45)
  pFiles.Move lFiles.Left, ((lFiles.Top + lFiles.Height) + 15), pBytes.Width, pBytes.Height
  lFile.Move pFiles.Left, ((pFiles.Top + pFiles.Height) + 45)
  pFile.Move lFile.Left, ((lFile.Top + lFile.Height) + 15), pFiles.Width, pFiles.Height
  lFileName.Move pFile.Left, ((pFile.Top + pFile.Height) + 15)
  pBack.Move 60, 210, ((pBack.Width - pBack.ScaleWidth) + (pBytes.Width + (pBytes.Left * 2))), ((pBack.Height - pBack.ScaleHeight) + ((lFileName.Top + lFileName.Height) + lBytes.Top))
  fBack.Move 15, 15, (pBack.Width + 120), (pBack.Height + 270)
  Me.Width = ((Me.Width - Me.ScaleWidth) + (fBack.Width + (fBack.Left * 2)))
  cmdStop.Move (Me.ScaleWidth - (cmdStop.Width + 45)), ((fBack.Top + fBack.Height) + 15)
  chkShutdown.Move (fBack.Left + 150), (cmdStop.Top + ((cmdStop.Height - chkShutdown.Height) / 2))
  Me.Height = ((Me.Height - Me.ScaleHeight) + ((cmdStop.Top + cmdStop.Height) + 45))
End Sub

Private Sub ConfirmExit()
  If CanAbort Then
    ExitNow = True
  End If
End Sub

Private Sub cmdStop_Click()
  ConfirmExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnloadedByCode Then
    Cancel = 1
    ConfirmExit
  End If
End Sub
