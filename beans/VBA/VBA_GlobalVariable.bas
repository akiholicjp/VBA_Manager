Option Explicit
Option Base 0

'=-=-=-=- VBA: Global Variable Definition =-=-=-=-

Dim i_gvarmanager_isinit As Variant

Private Sub GVAR_InitializeHelper()
   '=-=-=-=- VBA: Global Variable Initialize =-=-=-=-
End Sub

Public Sub GVAR_Initialize() 'VBA: Auto_Open
   Call GVAR_InitializeHelper
   Set i_gvarmanager_isinit = Application
End Sub

Public Function GVAR_check(Optional bReload As Boolean = True) As Boolean
   On Error Resume Next
   If Not i_gvarmanager_isinit Is Application Then
      On Error GoTo 0
      If bReload Then Call GVAR_Initialize
      GVAR_check = False
   Else
      On Error GoTo 0
      GVAR_check = True
   End If
End Function
