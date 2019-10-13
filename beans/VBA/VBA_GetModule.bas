Function VBA_GetModule(ByRef oComps As Object, ByVal sName As String) As Object
   Dim o As Object
   Set VBA_GetModule = Nothing
   For Each o In oComps
      If o.Name = sName Then
         Set VBA_GetModule = o
         Exit For
      End If
   Next o
End Function
