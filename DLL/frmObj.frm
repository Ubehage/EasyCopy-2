VERSION 5.00
Begin VB.Form frmObj 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmObj.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox pUnknown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3225
      Picture         =   "frmObj.frx":748A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   735
      Width           =   540
   End
   Begin VB.PictureBox p 
      Height          =   900
      Left            =   780
      ScaleHeight     =   840
      ScaleWidth      =   1725
      TabIndex        =   0
      Top             =   480
      Width           =   1785
   End
End
Attribute VB_Name = "frmObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SMALL_ICON_SIZE = 16
Private Const LARGE_ICON_SIZE = 32

Dim WithEvents UnloadTimer As EasyCopy2DLL.EasyTimer
Attribute UnloadTimer.VB_VarHelpID = -1

Friend Sub ResizePictureBox(LargeIcon As Boolean)
  Dim pW As Long
  Dim pH As Long
  pW = (p.Width - p.ScaleWidth)
  pH = (p.Height - p.ScaleHeight)
  If LargeIcon Then
    pW = (pW + (15 * LARGE_ICON_SIZE))
    pH = (pH + (15 * LARGE_ICON_SIZE))
  Else
    pW = (pW + (15 * SMALL_ICON_SIZE))
    pH = (pH + (15 * SMALL_ICON_SIZE))
  End If
  p.Width = pW
  p.Height = pH
  p.AutoRedraw = True
End Sub

Friend Sub SetUnloadTimer()
  If UnloadTimer Is Nothing Then
    Set UnloadTimer = New EasyCopy2DLL.EasyTimer
    UnloadTimer.Interval = 2000
  End If
  UnloadTimer.Enabled = True
End Sub

Private Sub UnloadTimer_Timer()
  UnloadTimer.Enabled = False
  Set UnloadTimer = Nothing
  Unload Me
End Sub
