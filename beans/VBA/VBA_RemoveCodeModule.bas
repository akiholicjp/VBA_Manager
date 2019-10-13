Function VBA_RemoveCodeModule(ByRef oComps As Object, ByRef o As Object) As Boolean
   If o.Type = 100 Then ' Module.Document
      With o.CodeModule
         .DeleteLines StartLine:=1, Count:=.CountOfLines
      End With
   Else
      oComps.Remove o
   End If
   VBA_RemoveCodeModule = True
End Function
