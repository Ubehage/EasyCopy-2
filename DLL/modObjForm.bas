Attribute VB_Name = "modObjForm"
Option Explicit

Dim FormLoaders As Long

Public Sub LoadObjectForm()
  FormLoaders = (FormLoaders + 1)
  If FormLoaders = 1 Then
    Load frmObj
  End If
End Sub

Public Sub UnloadObjectForm()
  If Not FormLoaders = 0 Then
    FormLoaders = (FormLoaders - 1)
    If FormLoaders = 0 Then
      frmObj.SetUnloadTimer
    End If
  End If
End Sub
