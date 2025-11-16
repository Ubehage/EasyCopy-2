VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl Progress 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin ComctlLib.ProgressBar p 
      Height          =   225
      Left            =   390
      TabIndex        =   0
      Top             =   630
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Min As Double
Dim m_Max As Double
Dim m_Value As Double

Public Property Get Min() As Double
  Min = m_Min
End Property
Public Property Get Max() As Double
  Max = m_Max
End Property
Public Property Get Value() As Double
  Value = m_Value
End Property
Public Property Get FloodPercent() As Single
  FloodPercent = GetFloodPercent
End Property
Public Property Let Min(New_Min As Double)
  If Not m_Min = New_Min Then
    If New_Min < m_Max Then
      m_Min = New_Min
      If m_Min > m_Value Then
        m_Value = m_Min
      End If
      Refresh
    End If
  End If
End Property
Public Property Let Max(New_Max As Double)
  If Not m_Max = New_Max Then
    If New_Max >= m_Min Then
      m_Max = New_Max
      If m_Max < m_Value Then
        m_Value = m_Max
      End If
      Refresh
    End If
  End If
End Property
Public Property Let Value(New_Value As Double)
  If Not m_Value = New_Value Then
    If (New_Value >= m_Min And New_Value <= m_Max) Then
      m_Value = New_Value
      Refresh
    End If
  End If
End Property

Public Sub Refresh()
  p.Min = 0
  p.Max = 100
  p.Value = GetFloodPercent
End Sub

Private Function GetFloodPercent() As Single
  Dim fP As Double
  On Error GoTo FloodError
  fP = ((100 / (m_Max - m_Min)) * (m_Value - m_Min))
FloodExit:
  On Error GoTo 0
  If fP < 0 Then
    fP = 0
  ElseIf fP > 100 Then
    fP = 100
  End If
  GetFloodPercent = CSng(fP)
  Exit Function
FloodError:
  fP = 100
  Resume FloodExit
End Function

Private Sub UserControl_InitProperties()
  m_Min = 0
  m_Max = 100
  m_Value = 50
  Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Min = PropBag.ReadProperty("Min", 0)
  m_Max = PropBag.ReadProperty("Max", 100)
  m_Value = PropBag.ReadProperty("Value", 50)
  Refresh
End Sub

Private Sub UserControl_Resize()
  p.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Min", m_Min
  PropBag.WriteProperty "Max", m_Max
  PropBag.WriteProperty "Value", m_Value
End Sub
