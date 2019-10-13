' VBA: Import VBA_GetCountOfDeclarationLines.bas

Sub VBA_MergeCodeModule(ByRef oBaseMod As Object, ByRef oImportMod As Object, Optional ByVal sTag As String = "")
   Dim s As String
   Dim iDeclPos As Long

   With oImportMod.CodeModule
      s = ""
      iDeclPos = VBA_GetCountOfDeclarationLines(oImportMod)
      If iDeclPos > 0 Then
         s = .Lines(1, iDeclPos)
         If sTag <> "" Then s = sTag & vbCrLf & s
      End If
   End With
   If s <> "" Then
      With oBaseMod.CodeModule
         .InsertLines VBA_GetCountOfDeclarationLines(oBaseMod) + 1, s
      End With
   End If

   With oImportMod.CodeModule
      s = ""
      If .CountOfLines - iDeclPos > 0 Then
         s = .Lines(iDeclPos + 1, .CountOfLines - iDeclPos)
         If sTag <> "" Then s = sTag & vbCrLf & s
      End If
   End With
   If s <> "" Then
      With oBaseMod.CodeModule
         .InsertLines .CountOfLines + 1, s
      End With
   End If
End Sub
