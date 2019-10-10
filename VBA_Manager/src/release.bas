Option Explicit
' VBA: Import beans/VBA/VBA_ExportCodeModule.bas: Private
' VBA: Import beans/VBA/VBA_GetModule.bas: Private
' VBA: Import beans/VBA/VBA_GetComponents.bas: Private
' VBA: Import beans/VBA/VBA_RemoveCodeModule.bas: Private
' VBA: Import beans/Path/GetOwnPath.bas: Private

Public Sub release()
   Dim oNew As Object, oOld As Object
   Set oNew = VBA_GetModule(VBA_GetComponents(), "Z_VBA_ManagerN")
   If oNew Is Nothing Then Exit Sub
   Set oOld = VBA_GetModule(VBA_GetComponents(), "Z_VBA_Manager")
   Call VBA_RemoveCodeModule(VBA_GetComponents(), oOld)
   oNew.Name = "Z_VBA_Manager"
   Call VBA_ExportCodeModule(oNew, GetOwnPath() & "/")
End Sub
