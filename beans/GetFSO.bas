Function GetFSO() As Object
   Static G_FSO As Object
   If G_FSO Is Nothing Then Set G_FSO = CreateObject("Scripting.FileSystemObject")
   Set GetFSO = G_FSO
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetFSO()
   Dim o1 As Object, o2 As Object
   Set o1 = GetFSO()
   assert "FileSystemObject", TypeName(o1)
   Set o2 = GetFSO()
   assert "FileSystemObject", TypeName(o2)
   assert ObjPtr(o1), ObjPtr(o2)
End Sub
